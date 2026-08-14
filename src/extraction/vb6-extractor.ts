import {
  Node,
  Edge,
  ExtractionResult,
  ExtractionError,
  UnresolvedReference,
  NodeKind,
} from '../types';
import { generateNodeId } from './tree-sitter-helpers';

/**
 * Custom extractor for Visual Basic 6 / Classic VB.
 *
 * VB6 has no adequate tree-sitter grammar (the VB.NET grammar errors on
 * VB6-core constructs — `Property Get/Let/Set`, `Event`, `WithEvents`,
 * `Implements`, `Type…End Type`, `As New`, `Declare`, and the form designer
 * header), and there is no local wasm toolchain to build a VBA grammar. VB6 is
 * statement-oriented, keyword-led, with explicit block terminators
 * (`End Sub/Function/Property/Type/Enum`) and `_` line continuation, so it is
 * parsed here with a deterministic line/token scanner — the same approach
 * CodeGraph already uses for `dfm-extractor.ts`, `mybatis-extractor.ts`,
 * `liquid-extractor.ts`. See `docs/vb6/IMPLEMENTATION_DECISIONS.md` (D1).
 *
 * Handles:
 *  - `.bas` (standard module)  → `module` node
 *  - `.cls` (class module)     → `class` node
 *  - `.frm` (form)             → `class` node + designer controls
 *  - `.ctl` (user control)     → `class` node + designer controls
 *  - `.vbp` / `.vbg`           → `module`/`namespace` project node + membership
 *
 * The VB6 ontology (prompt §8) is mapped onto CodeGraph's fixed NodeKind/
 * EdgeKind sets plus `decorators` markers (`vb6:*`). See
 * `docs/vb6/VB6_SEMANTIC_MODEL.md`.
 */

// ---------------------------------------------------------------------------
// Lexical helpers
// ---------------------------------------------------------------------------

/** A logical line: source text after comment-strip + continuation join. */
interface LogicalLine {
  text: string;
  /** 1-based physical line where this logical line starts. */
  line: number;
  /** 1-based physical line where this logical line ends. */
  endLine: number;
}

/**
 * Strip a trailing `'` comment (and `Rem` statements) from one physical line,
 * respecting string literals (VB6 escapes a quote by doubling it, which a
 * simple toggle handles correctly). Returns just the code portion.
 */
function stripComment(line: string): string {
  let inStr = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      inStr = !inStr;
    } else if (ch === "'" && !inStr) {
      return line.slice(0, i);
    }
  }
  // `Rem`-form comment as a whole statement (optionally after a label).
  const m = line.match(/^(\s*(?:[A-Za-z_]\w*:\s*)?)Rem(?:\s.*)?$/i);
  if (m) return m[1]!;
  return line;
}

/**
 * Turn raw source into logical lines: strip comments, join `_` continuations,
 * drop blank lines. Positions track the original physical lines.
 */
function toLogicalLines(source: string): LogicalLine[] {
  const phys = source.split(/\r?\n/);
  const out: LogicalLine[] = [];
  let acc = '';
  let startLine = 0;
  for (let i = 0; i < phys.length; i++) {
    const code = stripComment(phys[i]!);
    const isCont = /\s_\s*$/.test(code);
    if (acc === '') startLine = i + 1;
    if (isCont) {
      acc += code.replace(/\s_\s*$/, ' ');
      continue;
    }
    acc += code;
    const text = acc.trim();
    if (text !== '') out.push({ text, line: startLine, endLine: i + 1 });
    acc = '';
  }
  if (acc.trim() !== '') out.push({ text: acc.trim(), line: startLine, endLine: phys.length });
  return out;
}

// VB6 keywords that must never be treated as a call target.
const KEYWORDS = new Set(
  [
    'if', 'then', 'else', 'elseif', 'end', 'endif', 'for', 'next', 'do', 'loop',
    'while', 'wend', 'select', 'case', 'with', 'set', 'let', 'dim', 'const',
    'redim', 'static', 'public', 'private', 'friend', 'global', 'sub', 'function',
    'property', 'get', 'exit', 'goto', 'gosub', 'on', 'error', 'resume', 'return',
    'call', 'raiseevent', 'print', 'option', 'declare', 'implements', 'type',
    'enum', 'event', 'me', 'true', 'false', 'nothing', 'null', 'empty', 'and',
    'or', 'not', 'xor', 'mod', 'is', 'new', 'as', 'byval', 'byref', 'optional',
    'paramarray', 'to', 'step', 'each', 'in', 'preserve', 'stop', 'end sub',
    'withevents', 'attribute', 'begin', 'beginproperty', 'endproperty', 'like',
    'typeof', 'addressof', 'erase', 'lset', 'rset', 'mid', 'input', 'output',
    'append', 'random', 'binary', 'lock', 'unlock', 'open', 'close', 'write',
    'get', 'put', 'seek', 'width', 'spc', 'tab',
  ],
);

// Intrinsic VB6 types (not call targets when bare).
const TYPES = new Set(
  [
    'integer', 'long', 'string', 'boolean', 'double', 'single', 'variant',
    'object', 'byte', 'currency', 'date', 'decimal', 'collection', 'any',
    'longlong', 'longptr',
  ],
);

/**
 * Global objects supplied by the VB6 runtime. A member access on one of these
 * (`Debug.Print`, `Err.Number`, `App.Path`) targets the runtime, never a symbol
 * of the project, so it must not become a reference — otherwise every
 * `Debug.Print` leaves a dangling `Print` in `unresolved_refs`.
 */
const INTRINSIC_OBJECTS = new Set(
  ['debug', 'err', 'app', 'screen', 'printer', 'clipboard', 'forms', 'vb', 'vba'],
);

function isIntrinsicType(name: string): boolean {
  return TYPES.has(name.toLowerCase());
}

/**
 * Blank out the contents of string literals, keeping the quotes and the
 * original length so every offset stays valid. Without this an identifier
 * inside a literal (`"call Foo(1) now"`) is scanned as real code.
 */
function maskStrings(line: string): string {
  let out = '';
  let inStr = false;
  for (const ch of line) {
    if (ch === '"') {
      inStr = !inStr;
      out += ch;
    } else {
      out += inStr ? ' ' : ch;
    }
  }
  return out;
}

/** `Me.Items(3)` → `Items`; used to name the target of a `With` block. */
function lastSegment(expr: string): string {
  const m = expr.trim().match(/([A-Za-z_]\w*)\s*(\([^)]*\))?\s*$/);
  return m ? m[1]! : expr.trim();
}

/** Index of the previous non-blank character, or -1. */
function prevNonSpace(s: string, from: number): number {
  let i = from;
  while (i >= 0 && /\s/.test(s[i]!)) i--;
  return i;
}

/** Index of the next non-blank character, or -1. */
function nextNonSpace(s: string, from: number): number {
  let i = from;
  while (i < s.length && /\s/.test(s[i]!)) i++;
  return i < s.length ? i : -1;
}

/**
 * Parameters declared in a procedure signature `(ByVal a As Long, b() As Foo)`.
 *
 * The declared type matters as much as the name: a parameter is very often the
 * qualifier of a member call (`pItem.Refresh`), and without its type that call
 * has nothing to resolve against.
 */
function parameterDecls(signature: string): Array<{ name: string; type?: string }> {
  const paren = signature.match(/\(([\s\S]*)\)/);
  if (!paren) return [];
  const out: Array<{ name: string; type?: string }> = [];
  for (const part of splitDecls(paren[1]!)) {
    const m = part
      .trim()
      .replace(/^(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(?:ParamArray\s+)?/i, '')
      .match(/^([A-Za-z_]\w*)\s*(?:\([^)]*\))?\s*(?:As\s+(?:New\s+)?([A-Za-z_][\w.]*))?/i);
    if (m) out.push({ name: m[1]!, type: m[2] ? simpleType(m[2]) : undefined });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Core VB6 code extractor
// ---------------------------------------------------------------------------

type ContainerKind = 'module' | 'class';

interface Container {
  id: string;
  name: string;
  kind: NodeKind;
}

export class Vb6Extractor {
  private filePath: string;
  private source: string;
  private language: 'vb6' = 'vb6';
  private nodes: Node[] = [];
  private edges: Edge[] = [];
  private unresolved: UnresolvedReference[] = [];
  private errors: ExtractionError[] = [];
  private stack: Container[] = [];

  /**
   * Every name declared as data in this file (module fields, procedure locals,
   * parameters), lower-cased. VB6 spells an array access and a call the same
   * way — `Items(3)` — so the only way to tell them apart is knowing what was
   * declared. Statement scanning is therefore deferred until the whole file has
   * been read (see `pending`), because declarations may follow the procedures
   * that use them.
   */
  private declaredVars = new Set<string>();

  /** Statements queued for the deferred scan, in source order. */
  private pending: { text: string; fromId: string; line: number; procId: string }[] = [];

  /** Target names of the enclosing `With` blocks, innermost last. */
  private withStack: string[] = [];

  /**
   * @param containerKind  'module' for .bas, 'class' for .cls/.frm/.ctl.
   * @param containerDecorators  e.g. ['vb6:form'] for a form's top node.
   */
  constructor(
    filePath: string,
    source: string,
    private containerKind: ContainerKind = 'module',
    private containerDecorators: string[] = [],
  ) {
    this.filePath = filePath;
    this.source = source;
  }

  private baseName(): string {
    const b = this.filePath.split(/[\\/]/).pop() || this.filePath;
    return b.replace(/\.[^.]+$/, '');
  }

  private qname(name: string): string {
    const parts = this.stack.filter((c) => c.kind !== 'file').map((c) => c.name);
    parts.push(name);
    return parts.join('::');
  }

  private mkNode(
    kind: NodeKind,
    name: string,
    line: number,
    endLine: number,
    extra: Partial<Node> = {},
  ): Node {
    const node: Node = {
      id: generateNodeId(this.filePath, kind, this.qname(name), line),
      kind,
      name,
      qualifiedName: this.qname(name),
      filePath: this.filePath,
      language: this.language,
      startLine: line,
      endLine,
      startColumn: 0,
      endColumn: 0,
      updatedAt: Date.now(),
      ...extra,
    };
    this.nodes.push(node);
    const parent = this.stack[this.stack.length - 1];
    if (parent) this.edges.push({ source: parent.id, target: node.id, kind: 'contains' });
    return node;
  }

  private addRef(
    fromNodeId: string,
    name: string,
    kind: UnresolvedReference['referenceKind'],
    line: number,
    meta?: Record<string, unknown>,
  ): void {
    // The qualifier decides what a member reference means (`c.Compute` is a
    // method of c's type, `ModA.Compute` a procedure of that module), and
    // resolution happens later, reading the reference back from the database —
    // where `metadata` has no column. `candidates` does persist, and its
    // meaning is exactly "qualified name this might resolve to", so the
    // qualified form travels there.
    const qualifier = meta?.qualifier;
    const chainRoot = meta?.chainRoot;
    const candidates: string[] = [];
    if (typeof qualifier === 'string' && qualifier !== '') candidates.push(`${qualifier}.${name}`);
    if (typeof chainRoot === 'string' && chainRoot !== '') candidates.push(`${chainRoot}.${name}`);
    this.unresolved.push({
      fromNodeId,
      referenceName: name,
      referenceKind: kind,
      line,
      column: 0,
      filePath: this.filePath,
      language: this.language,
      candidates: candidates.length > 0 ? candidates : undefined,
      metadata: meta,
    });
  }

  /** The top module/class/form container node id (set after extract()). */
  containerId = '';

  /** Full extraction entry-point (creates file + container nodes). */
  extract(): ExtractionResult {
    const startTime = Date.now();
    try {
      const lines = toLogicalLines(this.source);
      // File node.
      const fileId = generateNodeId(this.filePath, 'file', this.filePath, 1);
      const fileNode: Node = {
        id: fileId,
        kind: 'file',
        name: this.filePath.split(/[\\/]/).pop() || this.filePath,
        qualifiedName: this.filePath,
        filePath: this.filePath,
        language: this.language,
        startLine: 1,
        endLine: this.source.split(/\r?\n/).length,
        startColumn: 0,
        endColumn: 0,
        updatedAt: Date.now(),
      };
      this.nodes.push(fileNode);
      this.stack.push({ id: fileId, name: fileNode.name, kind: 'file' });

      // Container (module/class/form/usercontrol) name from Attribute VB_Name.
      let containerName = this.baseName();
      for (const l of lines) {
        const m = l.text.match(/^Attribute\s+VB_Name\s*=\s*"([^"]*)"/i);
        if (m) { containerName = m[1]!; break; }
      }
      const container = this.mkNode(this.containerKind, containerName, 1, fileNode.endLine, {
        visibility: 'public',
        decorators: this.containerDecorators.length ? this.containerDecorators : undefined,
      });
      this.stack.push({ id: container.id, name: container.name, kind: container.kind });
      this.containerId = container.id;

      this.parseBody(lines, container.id);
    } catch (error) {
      this.errors.push({
        message: `VB6 extraction error: ${error instanceof Error ? error.message : String(error)}`,
        severity: 'error',
        code: 'parse_error',
        filePath: this.filePath,
      });
    }
    return {
      nodes: this.nodes,
      edges: this.edges,
      unresolvedReferences: this.unresolved,
      errors: this.errors,
      durationMs: Date.now() - startTime,
    };
  }

  /**
   * Parse the module/class body. `moduleId` is the top container node id, used
   * as the `from` for module-level references (Implements, module-level New).
   */
  private parseBody(lines: LogicalLine[], moduleId: string): void {
    // Each block mode carries the node it opened, so its `endLine` can be
    // closed on the matching `End …`. Without that a procedure spans a single
    // line and every consumer — codegraph_node, explore — shows its signature
    // instead of its body.
    type Mode =
      | { kind: 'module' }
      | { kind: 'proc'; end: RegExp; procId: string; node: Node }
      | { kind: 'type'; structId: string; node: Node }
      | { kind: 'enum'; enumId: string; node: Node };
    let mode: Mode = { kind: 'module' };

    for (const l of lines) {
      const t = l.text;

      // ---- inside a Sub/Function/Property body -----------------------------
      if (mode.kind === 'proc') {
        if (mode.end.test(t)) {
          mode.node.endLine = l.endLine;
          this.popTo(moduleId);
          mode = { kind: 'module' };
          continue;
        }
        // Local declarations become symbols right away (they need the
        // procedure on the scope stack); everything else is scanned later,
        // once every declaration in the file is known.
        if (!this.emitLocalDecls(t, l.line, l.endLine, mode.procId)) {
          this.pending.push({ text: t, fromId: mode.procId, line: l.line, procId: mode.procId });
        }
        continue;
      }

      // ---- inside a Type … End Type ---------------------------------------
      if (mode.kind === 'type') {
        if (/^End\s+Type\b/i.test(t)) { mode.node.endLine = l.endLine; this.popTo(moduleId); mode = { kind: 'module' }; continue; }
        const mm = t.match(/^([A-Za-z_]\w*)\s*(\([^)]*\))?\s+As\s+([A-Za-z_][\w.]*)/i);
        if (mm) {
          this.mkNode('field', mm[1]!, l.line, l.endLine, { signature: mm[3], returnType: simpleType(mm[3]!) });
        }
        continue;
      }

      // ---- inside an Enum … End Enum --------------------------------------
      if (mode.kind === 'enum') {
        if (/^End\s+Enum\b/i.test(t)) { mode.node.endLine = l.endLine; this.popTo(moduleId); mode = { kind: 'module' }; continue; }
        const mm = t.match(/^\[?([A-Za-z_]\w*)\]?\s*(?:=\s*(.+))?$/);
        if (mm) this.mkNode('enum_member', mm[1]!, l.line, l.endLine, { signature: mm[2]?.trim() });
        continue;
      }

      // ---- module scope ----------------------------------------------------
      // Procedures: Sub / Function
      let m = t.match(/^(?:(Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function)\s+([A-Za-z_]\w*)\s*(\([^)]*\))?(?:\s+As\s+([A-Za-z_][\w.]*(?:\([^)]*\))?))?/i);
      if (m) {
        const vis = visFrom(m[1]);
        const returnType = m[5] ? simpleType(m[5]) : undefined;
        const proc = this.mkNode('method', m[3]!, l.line, l.endLine, {
          visibility: vis,
          signature: t,
          returnType,
        });
        if (returnType && !isIntrinsicType(returnType)) this.addRef(proc.id, returnType, 'returns', l.line);
        // A Function assigns its result through its own name; that is not a
        // recursive call, and the name is not data either.
        this.stack.push({ id: proc.id, name: proc.name, kind: 'method' });
        this.emitParameters(m[4] ?? '', l.line);
        mode = { kind: 'proc', end: new RegExp(`^End\\s+${m[2]}\\b`, 'i'), procId: proc.id, node: proc };
        continue;
      }

      // Property Get/Let/Set
      m = t.match(/^(?:(Public|Private|Friend)\s+)?(?:Static\s+)?Property\s+(Get|Let|Set)\s+([A-Za-z_]\w*)\s*(\([^)]*\))?(?:\s+As\s+([A-Za-z_][\w.]*))?/i);
      if (m) {
        const accessor = m[2]!.toLowerCase();
        const prop = this.mkNode('property', m[3]!, l.line, l.endLine, {
          visibility: visFrom(m[1]),
          signature: t,
          decorators: [`vb6:property-${accessor}`],
          returnType: m[5] ? simpleType(m[5]) : undefined,
        });
        this.stack.push({ id: prop.id, name: prop.name, kind: 'property' });
        this.emitParameters(m[4] ?? '', l.line);
        mode = { kind: 'proc', end: /^End\s+Property\b/i, procId: prop.id, node: prop };
        continue;
      }

      // Declare (external / P-Invoke)
      m = t.match(
        /^(?:(Public|Private|Global)\s+)?Declare\s+(?:PtrSafe\s+)?(Function|Sub)\s+([A-Za-z_]\w*)\s+Lib\s+"([^"]*)"(?:\s+Alias\s+"([^"]*)")?\s*(\([^)]*\))?(?:\s+As\s+([A-Za-z_][\w.]*))?/i
      );
      if (m) {
        this.mkNode('method', m[3]!, l.line, l.endLine, {
          visibility: visFrom(m[1]),
          signature: t,
          decorators: ['vb6:declare'],
          returnType: m[7] ? simpleType(m[7]) : undefined,
          docstring: `Declare Lib "${m[4]}"${m[5] ? ` Alias "${m[5]}"` : ''}`,
        });
        for (const p of parameterDecls(m[6] ?? '')) this.declaredVars.add(p.name.toLowerCase());
        continue;
      }

      // Event declaration
      m = t.match(/^(?:(Public|Private)\s+)?Event\s+([A-Za-z_]\w*)\s*(\([^)]*\))?/i);
      if (m) {
        this.mkNode('field', m[2]!, l.line, l.endLine, {
          visibility: visFrom(m[1]),
          signature: t,
          decorators: ['vb6:event'],
        });
        continue;
      }

      // Implements
      m = t.match(/^Implements\s+([A-Za-z_][\w.]*)/i);
      if (m) {
        this.addRef(moduleId, simpleType(m[1]!)!, 'implements', l.line);
        continue;
      }

      // Type … End Type  (UDT → struct)
      m = t.match(/^(?:(Public|Private)\s+)?Type\s+([A-Za-z_]\w*)/i);
      if (m) {
        const st = this.mkNode('struct', m[2]!, l.line, l.endLine, { visibility: visFrom(m[1]) });
        this.stack.push({ id: st.id, name: st.name, kind: 'struct' });
        mode = { kind: 'type', structId: st.id, node: st };
        continue;
      }

      // Enum … End Enum
      m = t.match(/^(?:(Public|Private)\s+)?Enum\s+([A-Za-z_]\w*)/i);
      if (m) {
        const en = this.mkNode('enum', m[2]!, l.line, l.endLine, { visibility: visFrom(m[1]) });
        this.stack.push({ id: en.id, name: en.name, kind: 'enum' });
        mode = { kind: 'enum', enumId: en.id, node: en };
        continue;
      }

      // Const (module level)
      m = t.match(/^(?:(Public|Private|Global)\s+)?Const\s+(.+)$/i);
      if (m) {
        for (const decl of splitDecls(m[2]!)) {
          const cm = decl.match(/^\[?([A-Za-z_]\w*)\]?\s*(?:As\s+([A-Za-z_][\w.]*))?\s*=\s*(.+)$/i);
          if (cm) this.mkNode('constant', cm[1]!, l.line, l.endLine, {
            // A module-level Const with no modifier is Private, like Dim.
            visibility: visFrom(m[1], 'private'), signature: decl.trim(), returnType: cm[2] ? simpleType(cm[2]) : undefined,
          });
        }
        continue;
      }

      // Module-level variable / field: Public|Private|Global|Dim [WithEvents] decls
      m = t.match(/^(Public|Private|Global|Dim)\s+(WithEvents\s+)?(.+)$/i);
      if (m) {
        const withEvents = !!m[2];
        this.emitVarDecls(m[3]!, l.line, l.endLine, {
          visibility: visFrom(m[1], 'private'),
          withEvents,
          asField: true,
        });
        continue;
      }

      // Option / Attribute / VERSION / Begin(designer) / lone End → ignore here.
    }

    // Deferred scan: now every declaration in the file is known, so an
    // `Items(3)` can be told from a `Helper(3)`.
    let currentProc = '';
    for (const stmt of this.pending) {
      if (stmt.procId !== currentProc) {
        this.withStack = []; // a With block never spans procedures
        currentProc = stmt.procId;
      }
      this.scanStatement(stmt.text, stmt.fromId, stmt.line);
    }
    this.pending = [];
  }

  /**
   * Emit nodes for a procedure-local declaration (`Dim`/`Static`/`Const`/
   * `ReDim`). Returns true when the line was a declaration and needs no
   * further scanning.
   */
  private emitLocalDecls(t: string, line: number, endLine: number, procId: string): boolean {
    const m = t.match(/^(Dim|Static|Const|ReDim)(?:\s+Preserve)?\s+(WithEvents\s+)?(.+)$/i);
    if (!m) return false;

    // `ReDim` resizes an existing array — the symbol already exists.
    const isReDim = m[1]!.toLowerCase() === 'redim';
    const isConst = m[1]!.toLowerCase() === 'const';

    for (const decl of splitDecls(m[3]!)) {
      const dm = decl
        .trim()
        .match(/^([A-Za-z_]\w*)\s*(\([^)]*\))?\s*(?:As\s+(New\s+)?([A-Za-z_][\w.]*))?/i);
      if (!dm) continue;
      const name = dm[1]!;
      this.declaredVars.add(name.toLowerCase());
      const typeName = dm[4] ? simpleType(dm[4]) : undefined;

      if (!isReDim) {
        const node = this.mkNode(isConst ? 'constant' : 'variable', name, line, endLine, {
          signature: decl.trim(),
          returnType: typeName,
          isStatic: m[1]!.toLowerCase() === 'static' || undefined,
          decorators: m[2] ? ['vb6:withevents'] : undefined,
        });
        if (typeName && !isIntrinsicType(typeName)) this.addRef(node.id, typeName, 'type_of', line);
      }
      // `As New` instantiates at first use, from the procedure.
      if (dm[3] && typeName && !isIntrinsicType(typeName)) {
        this.addRef(procId, typeName, 'instantiates', line);
      }
    }
    return true;
  }

  /**
   * Emit a `parameter` node per declared parameter, contained by the
   * procedure. They are symbols of the language (§8) and, more practically,
   * the qualifier of a member call is very often one of them.
   */
  private emitParameters(signature: string, line: number): void {
    for (const p of parameterDecls(signature)) {
      this.declaredVars.add(p.name.toLowerCase());
      const node = this.mkNode('parameter', p.name, line, line, {
        returnType: p.type,
        signature: p.type ? `${p.name} As ${p.type}` : p.name,
      });
      if (p.type && !isIntrinsicType(p.type)) this.addRef(node.id, p.type, 'type_of', line);
    }
  }

  /** Pop the scope stack back down to (and keeping) the module container. */
  private popTo(moduleId: string): void {
    while (this.stack.length && this.stack[this.stack.length - 1]!.id !== moduleId) {
      this.stack.pop();
    }
  }

  /**
   * Emit variable/field nodes from a declaration list `a As Long, b As New Foo`.
   * `asField` distinguishes module-level fields from procedure locals.
   */
  private emitVarDecls(
    declList: string,
    line: number,
    endLine: number,
    opts: { visibility?: Node['visibility']; withEvents?: boolean; asField: boolean },
  ): void {
    for (const decl of splitDecls(declList)) {
      const dm = decl.match(/^([A-Za-z_]\w*)\s*(\([^)]*\))?\s*(?:As\s+(New\s+)?([A-Za-z_][\w.]*))?/i);
      if (!dm) continue;
      const name = dm[1]!;
      this.declaredVars.add(name.toLowerCase());
      const typeName = dm[4] ? simpleType(dm[4]) : undefined;
      const decorators: string[] = [];
      if (opts.withEvents) decorators.push('vb6:withevents');
      const node = this.mkNode(opts.asField ? 'field' : 'variable', name, line, endLine, {
        visibility: opts.visibility,
        signature: decl.trim(),
        returnType: typeName,
        decorators: decorators.length ? decorators : undefined,
      });
      // Intrinsic types are not symbols of the project: emitting a reference to
      // `Long` only leaves a row that can never resolve.
      if (typeName && !isIntrinsicType(typeName)) {
        this.addRef(node.id, typeName, 'type_of', line);
        if (dm[3]) this.addRef(node.id, typeName, 'instantiates', line); // As New
      }
    }
  }

  /** Scan one statement of a procedure body for calls and references. */
  private scanStatement(rawText: string, fromId: string, line: number): void {
    // A late-bound ProgID lives inside a string literal, so read it before the
    // literals are masked. The target is external and stays unresolved by
    // design (prompt §15) — but the ProgID itself is worth recording.
    for (const m of rawText.matchAll(/\b(?:CreateObject|GetObject)\s*\(\s*"([^"]+)"/gi)) {
      const arg = m[1]!;
      // GetObject's first argument may be a file path, which is not a ProgID.
      if (/[\\/:]/.test(arg) || !arg.includes('.')) continue;
      this.addRef(fromId, arg, 'instantiates', line, { vb6: 'late_binding', progId: arg });
    }

    const t = maskStrings(rawText);

    // A line label (`ErrHandler:`) is a jump target, not a call.
    if (/^[A-Za-z_]\w*\s*:\s*$/.test(t)) return;

    // With blocks: remember the target so `.Member` lines can be attributed.
    let m = t.match(/^With\s+(.+)$/i);
    if (m) {
      this.withStack.push(lastSegment(m[1]!));
      return;
    }
    if (/^End\s+With\b/i.test(t)) {
      this.withStack.pop();
      return;
    }

    // RaiseEvent Name(...)
    m = t.match(/^RaiseEvent\s+([A-Za-z_]\w*)/i);
    if (m) {
      this.addRef(fromId, m[1]!, 'references', line, { vb6: 'raises_event' });
      return;
    }

    // `Set x = New Foo` / `x = New Foo`
    for (const nm of t.matchAll(/\bNew\s+([A-Za-z_][\w.]*)/gi)) {
      const ty = simpleType(nm[1]!);
      if (!isIntrinsicType(ty)) this.addRef(fromId, ty, 'instantiates', line);
    }

    this.scanIdentifiers(t, fromId, line);
  }

  /**
   * Walk the identifiers of one statement and emit the references they imply.
   *
   * The rules follow VB6 syntax rather than guessing from shape:
   *  - `A.B` targets B — the member — and never A, which is only the qualifier;
   *  - `.B` inside a `With` targets B on the With target;
   *  - a name followed by `(` is a call, unless it was declared as data, in
   *    which case it is an array access;
   *  - a name that opens the statement and is not being assigned to is an
   *    implicit call (`DoWork`, `Log Msg:="x"`);
   *  - a bare name inside an expression is a variable read and is not emitted.
   */
  private scanIdentifiers(t: string, fromId: string, line: number): void {
    // `Call Foo(1)` — everything after `Call` is one invocation.
    const callPrefix = t.match(/^Call\s+/i);
    const body = callPrefix ? t.slice(callPrefix[0].length) : t;
    const forcedCall = !!callPrefix;

    const ident = /[A-Za-z_]\w*/g;
    let chainStart = -1; // offset where the current dotted chain begins
    let chainRoot: string | undefined; // the identifier that opens that chain
    let match: RegExpExecArray | null;

    while ((match = ident.exec(body)) !== null) {
      const name = match[0];
      const start = match.index;
      const end = start + name.length;

      const prevIdx = prevNonSpace(body, start - 1);
      const prevCh = prevIdx >= 0 ? body[prevIdx] : '';
      const isMember = prevCh === '.';

      const nextIdx = nextNonSpace(body, end);
      const nextCh = nextIdx >= 0 ? body[nextIdx] : '';
      const invoked = nextCh === '(';
      const assigned = nextCh === '=' && body[nextIdx + 1] !== '=';

      if (!isMember) {
        // Head of a new chain. A qualifier (`A` in `A.B`) is never a target.
        chainStart = start;
        chainRoot = name;
        if (nextCh === '.') continue;

        const lower = name.toLowerCase();
        if (KEYWORDS.has(lower) || TYPES.has(lower)) continue;

        if (invoked) {
          const kind = this.declaredVars.has(lower) ? 'references' : 'calls';
          this.addRef(fromId, name, kind, line, kind === 'calls' ? { vb6: 'call' } : { vb6: 'array_access' });
          continue;
        }
        // Statement-leading name with no parentheses: an implicit call, unless
        // it is the assignment target or a known variable.
        if ((start === 0 || forcedCall) && !assigned && !this.declaredVars.has(lower)) {
          this.addRef(fromId, name, 'calls', line, { vb6: 'implicit_call' });
        }
        continue;
      }

      // --- member of a chain ------------------------------------------------
      // Who is the member accessed on? Either the identifier before the dot, or
      // the enclosing `With` when the dot opens the expression.
      const beforeDot = prevNonSpace(body, prevIdx - 1);
      let qualifiedByExpr = beforeDot >= 0 && /[\w)\]]/.test(body[beforeDot]!);
      let qualifier: string | undefined;
      if (qualifiedByExpr) {
        const qm = body.slice(0, beforeDot + 1).match(/([A-Za-z_]\w*)\s*$/);
        qualifier = qm?.[1];
        // `If .Recordcount > 0 Then` — the word before the dot is a KEYWORD, so
        // the dot does not belong to it: this is a `With` member that happens
        // to follow a keyword. Attributing it to `If` both invents a qualifier
        // and loses the real one.
        if (qualifier && KEYWORDS.has(qualifier.toLowerCase())) {
          qualifier = undefined;
          qualifiedByExpr = false;
        }
      }
      if (!qualifiedByExpr) {
        qualifier = this.withStack[this.withStack.length - 1];
        if (qualifier === undefined) continue; // a dot with no owner: skip
        chainStart = prevIdx; // the chain starts at the dot itself
        chainRoot = qualifier;
      }

      // Members of the VB6 runtime objects are not project symbols.
      if (qualifier && INTRINSIC_OBJECTS.has(qualifier.toLowerCase())) continue;

      // `Me` is the current type, so the member is unqualified in practice.
      const meta: Record<string, unknown> = {};
      if (qualifier && qualifier.toLowerCase() !== 'me') meta.qualifier = qualifier;
      // `Adodc1.Recordset.MoveNext`: the immediate qualifier (`Recordset`) is
      // itself a member of something external, so it names no symbol. The ROOT
      // of the chain (`Adodc1`) usually does, and it is the object the whole
      // expression acts on — worth keeping as a second chance for resolution.
      if (chainRoot && qualifier && chainRoot.toLowerCase() !== qualifier.toLowerCase()) {
        meta.chainRoot = chainRoot;
      }

      const isChainHead = chainStart === 0 || forcedCall;
      if (invoked || (isChainHead && !assigned)) {
        this.addRef(fromId, name, 'calls', line, { ...meta, vb6: 'member_call' });
      } else {
        this.addRef(fromId, name, 'references', line, { ...meta, vb6: 'member_ref' });
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Small pure helpers
// ---------------------------------------------------------------------------

/**
 * VB6 visibility from the declaration keyword.
 *
 * Two rules that are easy to get wrong and both matter for scope: a procedure
 * with NO modifier is **Public**, and a module-level `Dim` is **Private**
 * (unlike `Public`/`Global`). `fallback` is what an absent keyword means for
 * the construct being declared.
 */
function visFrom(kw: string | undefined, fallback: Node['visibility'] = 'public'): Node['visibility'] {
  if (!kw) return fallback;
  const k = kw.toLowerCase();
  if (k === 'private' || k === 'dim') return 'private';
  if (k === 'friend') return 'internal';
  return 'public'; // Public / Global
}

/** Reduce a dotted/qualified type to its simple name (`VB.Timer` → `Timer`). */
function simpleType(t: string): string {
  return t.replace(/\(.*$/, '').split('.').pop()!.trim();
}

/**
 * Split a declaration list on top-level commas (not inside parens), e.g.
 * `a As Long, b(1 To 3) As Integer, c As New Foo`.
 */
function splitDecls(s: string): string[] {
  const out: string[] = [];
  let depth = 0, cur = '';
  for (const ch of s) {
    if (ch === '(') depth++;
    else if (ch === ')') depth--;
    if (ch === ',' && depth === 0) { out.push(cur); cur = ''; }
    else cur += ch;
  }
  if (cur.trim()) out.push(cur);
  return out;
}

// ---------------------------------------------------------------------------
// Form / UserControl composite extractor (designer + code body)
// ---------------------------------------------------------------------------

/**
 * `.frm` / `.ctl` are composite files: a designer header (`VERSION … Begin …
 * End`) describing the control tree, followed by the code body. The code body
 * is parsed by `Vb6Extractor` (which harmlessly ignores the designer lines —
 * none match a declaration form); this class additionally parses the designer
 * to emit the contained controls (`field` nodes, `vb6:control`), control arrays
 * (`vb6:controlarray`), and OCX component references. Modeled on
 * `dfm-extractor.ts`.
 */
export class Vb6FormExtractor {
  constructor(private filePath: string, private source: string) {}

  extract(): ExtractionResult {
    const ext = (this.filePath.match(/\.[^.]+$/)?.[0] || '').toLowerCase();
    const isCtl = ext === '.ctl';
    const decorators = [isCtl ? 'vb6:usercontrol' : 'vb6:form'];

    const code = new Vb6Extractor(this.filePath, this.source, 'class', decorators);
    const res = code.extract();

    try {
      this.parseDesigner(res, code.containerId);
    } catch (error) {
      res.errors.push({
        message: `VB6 designer parse error: ${error instanceof Error ? error.message : String(error)}`,
        severity: 'warning',
        code: 'designer_parse_error',
        filePath: this.filePath,
      });
    }
    return res;
  }

  /** Parse the `Begin … End` designer tree, appending controls to the form. */
  private parseDesigner(res: ExtractionResult, formId: string): void {
    const container = res.nodes.find((n) => n.id === formId);
    const formName = container?.name || 'Form';
    const lines = this.source.split(/\r?\n/);

    // `Object = "{GUID}#ver#0"; "FILE.ocx"` → the OCX component this form uses.
    // Modelled as an import node so the CLSID and the file are queryable, with
    // an edge from the form (prompt §8, §15).
    for (let i = 0; i < lines.length; i++) {
      const om = lines[i]!.match(/^\s*Object\s*=\s*"?\{([^}]*)\}[^;]*;\s*"?([^\s";]+)/i);
      if (om) {
        const file = om[2]!;
        const compNode = mkComponentNode(this.filePath, file, i + 1, 'vb6:ocx-reference', {
          clsid: om[1],
          file,
        });
        res.nodes.push(compNode);
        res.edges.push({
          source: formId,
          target: compNode.id,
          kind: 'references',
          metadata: { vb6: 'ocx_component' },
        });
      }
      if (/^\s*(Attribute|Option)\b/i.test(lines[i]!)) break; // designer header is over
    }

    // Control tree. Stack of node ids; the root Begin maps to the form itself.
    const stack: string[] = [formId];
    const names: string[] = [formName];
    let rootSeen = false;
    let lastControlIdx = -1;
    const controls: Node[] = [];

    for (let i = 0; i < lines.length; i++) {
      const raw = lines[i]!;
      const t = raw.trim();
      const beginM = t.match(/^Begin\s+([A-Za-z_][\w.]*)\s+([A-Za-z_]\w*)/i);
      if (beginM) {
        const typeName = beginM[1]!;
        const instName = beginM[2]!;
        if (!rootSeen) { rootSeen = true; stack.push(formId); names.push(instName); continue; }
        const isOcx = !/^VB\./i.test(typeName) && typeName.includes('.');
        const qn = names.slice(1).filter(Boolean).join('::') + '::' + instName;

        // A control array is written as several Begin blocks sharing one name
        // inside the same container. They are ONE symbol with many indices, so
        // the first block owns the node and the rest only mark it.
        const existing = controls.find(
          (c) => c.qualifiedName === qn && c.name === instName
        );
        if (existing) {
          if (!existing.decorators?.includes('vb6:controlarray')) {
            existing.decorators = [...(existing.decorators || []), 'vb6:controlarray'];
          }
          lastControlIdx = controls.indexOf(existing);
          stack.push(existing.id);
          names.push(instName);
          continue;
        }
        const node: Node = {
          id: generateNodeId(this.filePath, 'field', qn, i + 1),
          kind: 'field',
          name: instName,
          qualifiedName: qn,
          filePath: this.filePath,
          language: 'vb6',
          startLine: i + 1,
          endLine: i + 1,
          startColumn: 0,
          endColumn: 0,
          signature: typeName,
          returnType: simpleType(typeName),
          decorators: isOcx ? ['vb6:control', 'vb6:ocx'] : ['vb6:control'],
          updatedAt: Date.now(),
        };
        res.nodes.push(node);
        res.edges.push({ source: stack[stack.length - 1]!, target: node.id, kind: 'contains' });
        // A control instance IS OF the control's type (INSTANCE_OF, §8): a
        // `type_of` reference, so resolution treats it as a type position and
        // can reach a UserControl declared elsewhere in the project.
        //
        // The designer writes the type as `Library.Control`
        // (`MyEditOCX.MyEdit`), and that library IS the name of the ActiveX
        // project that builds the OCX. Keeping it as the qualifier is what
        // lets resolution pick the right UserControl when several projects
        // define the same control name — the client→OCX link of §12.
        const library = typeName.includes('.') ? typeName.slice(0, typeName.lastIndexOf('.')) : undefined;
        this.unref(
          res,
          node.id,
          simpleType(typeName),
          i + 1,
          { vb6: isOcx ? 'ocx_type' : 'control_type', ...(library ? { qualifier: library } : {}) },
          'type_of'
        );
        controls.push(node);
        lastControlIdx = controls.length - 1;
        stack.push(node.id);
        names.push(instName);
        continue;
      }
      if (/^End\s*$/i.test(t)) {
        if (stack.length > 1) { stack.pop(); names.pop(); }
        if (stack.length === 1) break; // closed the root Begin — designer done
        continue;
      }
      // Control-array member: `Index = N` inside the current control.
      const idxM = t.match(/^Index\s*=\s*(-?\d+)/i);
      if (idxM && lastControlIdx >= 0 && stack[stack.length - 1] === controls[lastControlIdx]!.id) {
        const c = controls[lastControlIdx]!;
        if (!c.decorators?.includes('vb6:controlarray')) {
          c.decorators = [...(c.decorators || []), 'vb6:controlarray'];
        }
        const seen = c.docstring?.match(/Index = (.+)$/)?.[1];
        c.docstring = `Control array, Index = ${seen ? `${seen}, ${idxM[1]}` : idxM[1]}`;
      }
    }
  }

  private unref(
    res: ExtractionResult,
    fromId: string,
    name: string,
    line: number,
    meta: Record<string, unknown>,
    kind: UnresolvedReference['referenceKind'] = 'references'
  ): void {
    const qualifier = meta.qualifier;
    res.unresolvedReferences.push({
      fromNodeId: fromId,
      referenceName: name,
      referenceKind: kind,
      candidates: typeof qualifier === 'string' && qualifier !== '' ? [`${qualifier}.${name}`] : undefined,
      line,
      column: 0,
      filePath: this.filePath,
      language: 'vb6',
      metadata: meta,
    });
  }
}

// ---------------------------------------------------------------------------
// Project files (.vbp / .vbg)
// ---------------------------------------------------------------------------

/**
 * Marker prefixing the JSON list of member files on a project node's
 * `docstring`. Resolution needs to know which files make up a project before
 * any `contains` edge exists — those edges are produced BY resolution — so the
 * membership travels on the node itself. Read with `projectMemberFiles()`.
 */
export const VB6_PROJECT_FILES = 'vb6:files=';

/** The member files recorded on a project node, or [] if it carries none. */
export function projectMemberFiles(docstring: string | undefined): string[] {
  if (!docstring || !docstring.startsWith(VB6_PROJECT_FILES)) return [];
  try {
    const parsed = JSON.parse(docstring.slice(VB6_PROJECT_FILES.length));
    return Array.isArray(parsed) ? parsed.filter((x): x is string => typeof x === 'string') : [];
  } catch {
    return [];
  }
}

/**
 * `.vbp` (project) and `.vbg` (project group). We extract project scope: the
 * project node, its member files (Form/Module/Class/UserControl/Designer), and
 * external references (Reference/Object → COM/OCX). The goal is scope + code
 * resolution, not build/deploy metadata (prompt §10).
 */
export class Vb6ProjectExtractor {
  constructor(private filePath: string, private source: string) {}

  extract(): ExtractionResult {
    const startTime = Date.now();
    const nodes: Node[] = [];
    const edges: Edge[] = [];
    const unresolved: UnresolvedReference[] = [];
    const errors: ExtractionError[] = [];
    const lines = this.source.split(/\r?\n/);
    const isGroup = /\.vbg$/i.test(this.filePath);

    const fileId = generateNodeId(this.filePath, 'file', this.filePath, 1);
    nodes.push(mkFile(this.filePath, lines.length, fileId));

    if (isGroup) {
      // .vbg: StartupProject / Project=<relpath> lines → referenced projects.
      const projName = this.filePath.split(/[\\/]/).pop()!.replace(/\.[^.]+$/, '');
      const group = mkContainer('namespace', projName, this.filePath, fileId, lines.length, ['vb6:projectgroup']);
      nodes.push(group.node); edges.push(group.edge);
      for (let i = 0; i < lines.length; i++) {
        const m = lines[i]!.match(/^\s*(?:StartupProject|Project)\s*=\s*(.+?)\s*$/i);
        if (m) unresolved.push(ref(group.node.id, memberName(m[1]!), 'imports', i + 1, this.filePath, { vb6: 'project_reference' }));
      }
    } else {
      // .vbp: Name=, Type=, member files, references.
      let projName = this.filePath.split(/[\\/]/).pop()!.replace(/\.[^.]+$/, '');
      let projType = 'Exe';
      for (const raw of lines) {
        const nm = raw.match(/^\s*Name\s*=\s*"?([^"\r\n]+?)"?\s*$/i);
        if (nm) { projName = nm[1]!; break; }
      }
      for (const raw of lines) { const tm = raw.match(/^\s*Type\s*=\s*(\w+)/i); if (tm) { projType = tm[1]!; break; } }
      const proj = mkContainer('module', projName, this.filePath, fileId, lines.length, ['vb6:project']);
      proj.node.signature = `VB6 ${projType} project`;
      // Files this project is made of, as index-relative paths. A VB6 name is
      // scoped to its project, so resolution needs this membership — and it
      // cannot wait for the `contains` edges below, which are themselves
      // produced by the resolution pass that needs the answer. Carrying it on
      // the node makes project scope available from the first reference on.
      const memberFiles: string[] = [];
      nodes.push(proj.node); edges.push(proj.edge);

      for (let i = 0; i < lines.length; i++) {
        const raw = lines[i]!;
        // Member files: Form=..., Module=Name; path, Class=Name; path, UserControl=..., Designer=...
        let m = raw.match(/^\s*(Form|Module|Class|UserControl|Designer|RelatedDoc|UserDocument|PropertyPage)\s*=\s*(.+?)\s*$/i);
        if (m) {
          unresolved.push(ref(proj.node.id, memberName(m[2]!), 'contains', i + 1, this.filePath, { vb6: 'project_member', memberKind: m[1] }));
          const file = memberFilePath(this.filePath, m[2]!);
          if (file) memberFiles.push(file);
          continue;
        }
        // Components: `Object={GUID}#maj.min#lcid; FILE.ocx` (early-bound OCX).
        m = raw.match(/^\s*Object\s*=\s*"?\{([^}]*)\}[^;]*;\s*"?([^\s";]+)/i);
        if (m) {
          const node = mkComponentNode(this.filePath, m[2]!, i + 1, 'vb6:ocx-reference', {
            clsid: m[1],
            file: m[2],
          });
          nodes.push(node);
          edges.push({ source: proj.node.id, target: node.id, kind: 'imports', metadata: { vb6: 'ocx_component' } });
          continue;
        }
        // COM type-library references: `Reference=*\G{GUID}#maj.min#lcid#path#Description`.
        if (/^\s*Reference\s*=/i.test(raw)) {
          const parts = raw.split('#');
          const guid = raw.match(/\{([0-9A-Fa-f-]+)\}/)?.[1];
          const desc = parts.length >= 2 ? parts[parts.length - 1]!.trim() : '';
          const libPath = parts.length >= 2 ? parts[parts.length - 2]!.trim() : '';
          const name = desc || memberName(libPath) || guid || 'COMReference';
          const node = mkComponentNode(this.filePath, name, i + 1, 'vb6:com-reference', {
            ...(guid ? { clsid: guid } : {}),
            ...(libPath ? { typelib: libPath } : {}),
          });
          nodes.push(node);
          edges.push({ source: proj.node.id, target: node.id, kind: 'imports', metadata: { vb6: 'com_reference' } });
        }
      }

      // Membership, readable straight off the node (see VB6_PROJECT_FILES).
      if (memberFiles.length > 0) {
        proj.node.docstring = `${VB6_PROJECT_FILES}${JSON.stringify(memberFiles)}`;
      }
    }

    return { nodes, edges, unresolvedReferences: unresolved, errors, durationMs: Date.now() - startTime };
  }
}

/**
 * Resolve a `.vbp` member entry to the file it names, relative to the index
 * root: `Module1; sub\\Module1.bas` in `app/App.vbp` → `app/sub/Module1.bas`.
 * Returns undefined when the entry carries no file name.
 */
function memberFilePath(vbpPath: string, entry: string): string | undefined {
  const parts = entry.split(';');
  const raw = (parts.length > 1 ? parts[1]! : parts[0]!).trim().replace(/"/g, '');
  if (raw === '' || !/\.[A-Za-z]+$/.test(raw)) return undefined;
  const rel = raw.replace(/\\/g, '/');
  const dir = vbpPath.replace(/\\/g, '/').replace(/\/[^/]*$/, '');
  const joined = dir === vbpPath.replace(/\\/g, '/') ? rel : `${dir}/${rel}`;
  // Normalise `a/./b` and `a/x/../b` without touching the platform's fs.
  const out: string[] = [];
  for (const seg of joined.split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return out.join('/');
}

function mkFile(filePath: string, endLine: number, id: string): Node {
  return {
    id, kind: 'file', name: filePath.split(/[\\/]/).pop() || filePath, qualifiedName: filePath,
    filePath, language: 'vb6', startLine: 1, endLine, startColumn: 0, endColumn: 0, updatedAt: Date.now(),
  };
}

/**
 * A COM type library or OCX component the project (or a form) depends on.
 * Kept as an `import` node so the CLSID / type library path survive as
 * queryable data instead of being lost in an unresolvable reference.
 */
function mkComponentNode(
  filePath: string,
  rawName: string,
  line: number,
  decorator: 'vb6:com-reference' | 'vb6:ocx-reference',
  metadata: Record<string, unknown>
): Node {
  const name = rawName.replace(/\.(ocx|dll|tlb|olb|exe)$/i, '');
  return {
    id: generateNodeId(filePath, 'import', `${decorator}:${name}`, line),
    kind: 'import',
    name,
    qualifiedName: name,
    filePath,
    language: 'vb6',
    startLine: line,
    endLine: line,
    startColumn: 0,
    endColumn: 0,
    decorators: [decorator],
    docstring: Object.entries(metadata)
      .filter(([, v]) => v !== undefined)
      .map(([k, v]) => `${k}: ${v}`)
      .join(', ') || undefined,
    updatedAt: Date.now(),
  };
}

function mkContainer(kind: NodeKind, name: string, filePath: string, fileId: string, endLine: number, decorators: string[]) {
  const node: Node = {
    id: generateNodeId(filePath, kind, name, 1), kind, name, qualifiedName: name, filePath,
    language: 'vb6', startLine: 1, endLine, startColumn: 0, endColumn: 0, decorators, visibility: 'public', updatedAt: Date.now(),
  };
  return { node, edge: { source: fileId, target: node.id, kind: 'contains' as const } };
}

function ref(fromId: string, name: string, kind: UnresolvedReference['referenceKind'], line: number, filePath: string, meta: Record<string, unknown>): UnresolvedReference {
  return { fromNodeId: fromId, referenceName: name, referenceKind: kind, line, column: 0, filePath, language: 'vb6', metadata: meta };
}

/** `Name; relative\path\File.frm` or `File.bas` → base member name. */
function memberName(spec: string): string {
  const semi = spec.split(';');
  const nameOrPath = (semi.length > 1 ? semi[0]! : semi[0]!).trim();
  // For `Module1; Module1.bas` the first segment IS the module name.
  if (semi.length > 1) return nameOrPath.replace(/"/g, '');
  return nameOrPath.split(/[\\/]/).pop()!.replace(/\.[^.]+$/, '').replace(/"/g, '');
}

// ---------------------------------------------------------------------------
// Dispatcher
// ---------------------------------------------------------------------------

/** Entry point used by `extractFromSource` for every VB6 file. */
export function extractVb6File(filePath: string, source: string): ExtractionResult {
  const ext = (filePath.match(/\.[^.]+$/)?.[0] || '').toLowerCase();
  if (ext === '.vbp' || ext === '.vbg') return new Vb6ProjectExtractor(filePath, source).extract();
  if (ext === '.frm' || ext === '.ctl') return new Vb6FormExtractor(filePath, source).extract();
  const kind: ContainerKind = ext === '.bas' ? 'module' : 'class';
  return new Vb6Extractor(filePath, source, kind).extract();
}
