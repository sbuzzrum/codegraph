/**
 * VB6 reference resolution.
 *
 * VB6 resolves a name by scope, not by similarity, and the generic name
 * matcher has no way to know that: it will happily bind a call to a `Private`
 * procedure of another module, or pick one of two same-named candidates at
 * random. Both are edges that cannot exist in a running VB6 program.
 *
 * This resolver applies the language's own rules instead (prompt §9), and
 * where they do not determine a single target it returns null so the reference
 * stays UNRESOLVED — which the specification ranks above a plausible guess
 * (§21).
 *
 * Scope order for an unqualified name inside a procedure:
 *   1. members of the enclosing module / class / form / UserControl;
 *   2. Public procedures of standard modules (`.bas`) in the project;
 *   3. nothing — unresolved.
 *
 * A qualified name (`ModA.Process`, `c.Compute`, `.Compute` inside a `With`)
 * targets a member OF THE QUALIFIER: either a module/class named by the
 * qualifier, or the type of the variable the qualifier names. When that type
 * is `Object` or `Variant` the call is late-bound and has no static target at
 * all.
 *
 * See docs/vb6/VB6_SEMANTIC_MODEL.md; every rule is pinned by a fixture in
 * __tests__/fixtures/vb6/.
 */
import { Node } from '../types';
import { QueryBuilder } from '../db/queries';
import { UnresolvedRef, ResolvedRef } from './types';
import { LRUCache } from './lru-cache';
import { projectMemberFiles } from '../extraction/vb6-extractor';

/** Kinds that can own members — a VB6 module, class, form or UserControl. */
const CONTAINER_KINDS = new Set(['module', 'class', 'namespace', 'struct', 'interface']);

/** Kinds a type position (`As X`, `New X`, `Implements X`) can name. */
const TYPE_KINDS = new Set(['class', 'module', 'struct', 'enum', 'interface']);

/**
 * Declared types that carry no static member information. A call through one
 * of these is late-bound: VB6 itself only finds the target at run time, so
 * neither can we (prompt §15).
 */
const LATE_BOUND_TYPES = new Set(['object', 'variant']);

/** A standard module — the only container whose Public members are global. */
function isStandardModule(node: Node): boolean {
  return node.filePath.toLowerCase().endsWith('.bas');
}

function isPrivate(node: Node): boolean {
  return node.visibility === 'private';
}

/** A symbol that holds data — the kind of thing a qualifier names. */
function isData(node: Node): boolean {
  return (
    node.kind === 'variable' ||
    node.kind === 'field' ||
    node.kind === 'constant' ||
    node.kind === 'parameter'
  );
}

export class Vb6Resolver {
  // Bounded like every other resolver cache: an unbounded Map grows with each
  // distinct lookup and is how this layer has OOM'd on large codebases before.
  private byName: LRUCache<string, Node[]>;
  private byFile: LRUCache<string, Node[]>;

  /**
   * Which projects each file belongs to, from `.vbp` membership. A VB6 name
   * is scoped to its PROJECT, not to everything that happens to be indexed:
   * when several projects sit in one index — the normal case for a product
   * built from many `.vbp` — the same procedure name legitimately exists in
   * each, and without this the resolver sees an ambiguity that does not exist
   * for the caller and resolves nothing.
   *
   * Built once, lazily: null until the first reference needs it.
   */
  private projectsByFile: Map<string, Set<string>> | null = null;

  /** @param cacheLimit shared with the other resolver caches (see index.ts). */
  constructor(private queries: QueryBuilder, cacheLimit: number) {
    this.byName = new LRUCache(cacheLimit);
    this.byFile = new LRUCache(cacheLimit);
  }

  /** Drop memoised node lookups (the graph grows during resolution). */
  clearCaches(): void {
    this.byName.clear();
    this.byFile.clear();
    this.projectsByFile = null;
  }

  private projectIndex(): Map<string, Set<string>> {
    if (this.projectsByFile) return this.projectsByFile;
    const map = new Map<string, Set<string>>();
    for (const project of this.queries.getNodesByKind('module')) {
      if (project.language !== 'vb6' || !project.decorators?.includes('vb6:project')) continue;
      // Read from the node, not from `contains` edges: those edges are created
      // by the very resolution pass that asks this question, so at this point
      // most of them do not exist yet.
      for (const file of projectMemberFiles(project.docstring)) {
        let owners = map.get(file);
        if (!owners) {
          owners = new Set();
          map.set(file, owners);
        }
        owners.add(project.id);
      }
      // The `.vbp` itself belongs to its project.
      let own = map.get(project.filePath);
      if (!own) {
        own = new Set();
        map.set(project.filePath, own);
      }
      own.add(project.id);
    }
    this.projectsByFile = map;
    return map;
  }

  /** True when `file` is part of a project whose NAME is `projectName`. */
  private belongsToProjectNamed(file: string, projectName: string): boolean {
    const owners = this.projectIndex().get(file);
    if (!owners) return false;
    const wanted = projectName.toLowerCase();
    for (const id of owners) {
      const project = this.queries.getNodeById(id);
      if (project && project.name.toLowerCase() === wanted) return true;
    }
    return false;
  }

  /**
   * True when both files belong to a common project, false when they provably
   * do not, null when membership is unknown for either — a file listed in no
   * indexed `.vbp` must not be cut off from resolution.
   */
  private shareProject(a: string, b: string): boolean | null {
    const index = this.projectIndex();
    const left = index.get(a);
    const right = index.get(b);
    if (!left || !right) return null;
    for (const project of left) if (right.has(project)) return true;
    return false;
  }

  /**
   * Resolve one VB6 reference, or null to leave it unresolved.
   */
  resolve(ref: UnresolvedRef): ResolvedRef | null {
    const candidatesAll = this.lookup(ref.referenceName);
    if (candidatesAll.length === 0) {
      // No symbol of that name anywhere. For a QUALIFIED reference that is the
      // normal case, not a dead end: `txtName.Text` names a member of a
      // standard control, which is not in the graph — but `txtName` is, and
      // the link to it is worth keeping (see linkToQualifier).
      const onlyQualifier = qualifierOf(ref);
      if (!onlyQualifier || ref.referenceKind === 'type_of') return null;
      return (
        this.linkToQualifier(ref, onlyQualifier, this.findData(onlyQualifier, ref.filePath)) ??
        this.linkToChainRoot(ref)
      );
    }

    switch (ref.referenceKind) {
      // `.vbp`/`.vbg` membership: the target is a whole module/class/form.
      case 'contains':
      case 'imports':
        return this.pick(
          ref,
          candidatesAll.filter((n) => CONTAINER_KINDS.has(n.kind)),
          { scope: 'project' },
          'qualified-name'
        );

      // Type positions name a project-global type.
      case 'implements':
      case 'type_of':
      case 'instantiates':
      case 'returns': {
        const types = candidatesAll.filter((n) => TYPE_KINDS.has(n.kind));
        // A control's type is written `Library.Control`, and that library is
        // the ActiveX project that builds the OCX. When it is indexed too,
        // it says which of several same-named UserControls is meant (§12).
        const library = qualifierOf(ref);
        if (library && types.length > 1) {
          const owned = types.filter((n) => this.belongsToProjectNamed(n.filePath, library));
          if (owned.length > 0) {
            return this.pick(ref, owned, { scope: 'type', library }, 'qualified-name');
          }
        }
        return this.pick(ref, types, { scope: 'type' });
      }

      default:
        break;
    }

    // Anything a Private member of another file cannot be seen from here.
    const visible = candidatesAll.filter((n) => !isPrivate(n) || n.filePath === ref.filePath);
    const qualifier = qualifierOf(ref);
    return qualifier
      ? this.resolveQualified(ref, qualifier, visible)
      : this.resolveUnqualified(ref, visible);
  }

  /**
   * `Qualifier.Member` — the target is a member of the qualifier, never the
   * qualifier itself.
   */
  private resolveQualified(ref: UnresolvedRef, qualifier: string, visible: Node[]): ResolvedRef | null {
    // The qualifier may name a module/class/form directly (`ModA.Process`,
    // `Form1.Reload` through the default instance).
    for (const container of this.lookup(qualifier).filter((n) => CONTAINER_KINDS.has(n.kind))) {
      const members = visible.filter((n) => n.filePath === container.filePath && n.id !== container.id);
      const hit = this.pick(ref, members, { scope: 'module', qualifier }, 'qualified-name');
      if (hit) return hit;
    }

    // Otherwise it names a variable, field or control: the member belongs to
    // that variable's declared type.
    const variable = this.findData(qualifier, ref.filePath);
    const typeName = variable?.returnType;

    if (typeName && !LATE_BOUND_TYPES.has(typeName.toLowerCase())) {
      for (const type of this.lookup(typeName).filter((n) => CONTAINER_KINDS.has(n.kind))) {
        const members = visible.filter((n) => n.filePath === type.filePath && n.id !== type.id);
        const hit = this.pick(ref, members, { scope: 'type', qualifier }, 'instance-method');
        if (hit) return hit;
      }
    }

    return this.linkToQualifier(ref, qualifier, variable) ?? this.linkToChainRoot(ref);
  }

  /** Last chance for a member chain: bind to the object the chain starts from. */
  private linkToChainRoot(ref: UnresolvedRef): ResolvedRef | null {
    const root = chainRootOf(ref);
    if (!root) return null;
    return this.linkToQualifier(ref, root, this.findData(root, ref.filePath));
  }

  /**
   * The member could not be resolved — its type is a standard VB control, an
   * external COM object, or late-bound. The OBJECT it was used on is in the
   * graph, though, and losing that link too would throw away the one thing we
   * do know.
   *
   * So the reference attaches to the qualifier, carrying the member name in
   * `metadata.member`. It answers "what touches this control / this recordset",
   * which on a VB6 codebase is most of what a reader wants, and it never
   * pretends to have found the member: the edge points at the object, not at a
   * same-named property picked from somewhere else in the project (§21).
   *
   * `createEdges` demotes a `calls` landing on data to `references`, so
   * `obj.DoThing(1)` does not read as "calls obj".
   */
  private linkToQualifier(ref: UnresolvedRef, qualifier: string, data: Node | undefined): ResolvedRef | null {
    const target =
      data ??
      // A `With` target or a qualifier declared elsewhere may still name a
      // module/class/form of the project.
      this.lookup(qualifier).find((n) => CONTAINER_KINDS.has(n.kind));
    if (!target) return null;
    return {
      original: ref,
      targetNodeId: target.id,
      confidence: 0.5,
      resolvedBy: 'instance-method',
      metadata: { scope: 'object', vb6: 'member_on', member: ref.referenceName, qualifier },
    };
  }

  /** An unqualified name: own container first, then global standard modules. */
  private resolveUnqualified(ref: UnresolvedRef, visible: Node[]): ResolvedRef | null {
    const own = visible.filter((n) => n.filePath === ref.filePath);
    const inOwnContainer = this.pick(ref, own, { scope: 'module' });
    if (inOwnContainer) return inOwnContainer;

    // Only standard modules expose their Public procedures project-wide. A
    // class or form member is unreachable this way, however well the name
    // matches.
    const globals = visible.filter((n) => isStandardModule(n) && !isPrivate(n) && !CONTAINER_KINDS.has(n.kind));
    return this.pick(ref, globals, { scope: 'project' });
  }

  /**
   * Choose the single target among candidates, or null when the language does
   * not determine one. Ambiguity is deliberately NOT broken by guessing.
   */
  private pick(
    ref: UnresolvedRef,
    candidates: Node[],
    meta: Record<string, unknown>,
    resolvedBy: ResolvedRef['resolvedBy'] = 'exact-match'
  ): ResolvedRef | null {
    let pool = dedupe(candidates);
    if (pool.length === 0) return null;

    if (pool.length > 1) {
      // Candidates outside the caller's project are not candidates at all.
      const ownProject = pool.filter((n) => this.shareProject(ref.filePath, n.filePath) === true);
      if (ownProject.length > 0) pool = ownProject;
    }

    if (pool.length > 1) {
      // A call names something callable.
      if (ref.referenceKind === 'calls') {
        const callable = pool.filter((n) => n.kind === 'method' || n.kind === 'function' || n.kind === 'property');
        if (callable.length > 0) pool = callable;
      }
      // `Property Get` and `Property Let` share one name by design; the getter
      // is the readable side and the only one a call can mean.
      if (pool.length > 1 && pool.every((n) => n.kind === 'property')) {
        const getter = pool.filter((n) => n.decorators?.includes('vb6:property-get'));
        if (getter.length === 1) pool = getter;
      }
      if (pool.length > 1) return null; // genuinely ambiguous → unresolved
    }

    const target = pool[0]!;
    const extra: Record<string, unknown> = { ...meta };
    // A reference to an Event declaration is a RaiseEvent site.
    if (ref.referenceKind === 'references' && target.decorators?.includes('vb6:event')) {
      extra.vb6 = 'raises_event';
    }
    return {
      original: ref,
      targetNodeId: target.id,
      confidence: meta.scope === 'project' ? 0.9 : 0.95,
      resolvedBy,
      metadata: extra,
    };
  }

  /** Nodes of a name, case-insensitively — VB6 identifiers are not case-sensitive. */
  private lookup(name: string): Node[] {
    const key = name.toLowerCase();
    let hit = this.byName.get(key);
    if (!hit) {
      hit = this.queries.getNodesByLowerName(key).filter((n) => n.language === 'vb6');
      this.byName.set(key, hit);
    }
    return hit;
  }

  /**
   * The data symbol a qualifier names, following VB6 scope: locals and
   * parameters of the calling file first, then `Public` variables of the
   * project's standard modules.
   *
   * The second step matters more than it looks: a VB6 application keeps its
   * shared objects in `Public` module variables, and they are qualifiers all
   * over the codebase. Looking only in the calling file loses every one of
   * them.
   */
  private findData(name: string, filePath: string): Node | undefined {
    const key = name.toLowerCase();
    let nodes = this.byFile.get(filePath);
    if (!nodes) {
      nodes = this.queries.getNodesByFile(filePath);
      this.byFile.set(filePath, nodes);
    }
    const own = nodes.find((n) => n.name.toLowerCase() === key && isData(n));
    if (own) return own;

    const globals = this.lookup(name).filter(
      (n) => isData(n) && isStandardModule(n) && n.visibility !== 'private'
    );
    if (globals.length === 0) return undefined;
    if (globals.length === 1) return globals[0];

    // Several projects may declare the same global; prefer the caller's.
    const ownProject = globals.filter((n) => this.shareProject(filePath, n.filePath) === true);
    return ownProject.length === 1 ? ownProject[0] : undefined;
  }
}

/** The qualifier carried by the extractor in `candidates` (`c.Compute` → `c`). */
function qualifierOf(ref: UnresolvedRef): string | undefined {
  return qualifierAt(ref, 0);
}

/**
 * The root of a member chain, when the extractor recorded one:
 * `Adodc1.Recordset.MoveNext` puts `Recordset.MoveNext` first and
 * `Adodc1.MoveNext` second. Used as a second chance when the immediate
 * qualifier names nothing.
 */
function chainRootOf(ref: UnresolvedRef): string | undefined {
  return qualifierAt(ref, 1);
}

function qualifierAt(ref: UnresolvedRef, index: number): string | undefined {
  const candidate = ref.candidates?.[index];
  if (!candidate) return undefined;
  const dot = candidate.lastIndexOf('.');
  return dot > 0 ? candidate.slice(0, dot) : undefined;
}

function dedupe(nodes: Node[]): Node[] {
  const seen = new Set<string>();
  const out: Node[] = [];
  for (const n of nodes) {
    if (seen.has(n.id)) continue;
    seen.add(n.id);
    out.push(n);
  }
  return out;
}