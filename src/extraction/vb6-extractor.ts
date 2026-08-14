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
    this.unresolved.push({
      fromNodeId,
      referenceName: name,
      referenceKind: kind,
      line,
      column: 0,
      filePath: this.filePath,
      language: this.language,
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
    type Mode =
      | { kind: 'module' }
      | { kind: 'proc'; end: RegExp; procId: string }
      | { kind: 'type'; structId: string }
      | { kind: 'enum'; enumId: string };
    let mode: Mode = { kind: 'module' };

    for (const l of lines) {
      const t = l.text;

      // ---- inside a Sub/Function/Property body -----------------------------
      if (mode.kind === 'proc') {
        if (mode.end.test(t)) {
          this.popTo(moduleId);
          mode = { kind: 'module' };
          continue;
        }
        this.scanStatement(t, mode.procId, l.line);
        continue;
      }

      // ---- inside a Type … End Type ---------------------------------------
      if (mode.kind === 'type') {
        if (/^End\s+Type\b/i.test(t)) { this.popTo(moduleId); mode = { kind: 'module' }; continue; }
        const mm = t.match(/^([A-Za-z_]\w*)\s*(\([^)]*\))?\s+As\s+([A-Za-z_][\w.]*)/i);
        if (mm) {
          this.mkNode('field', mm[1]!, l.line, l.endLine, { signature: mm[3], returnType: simpleType(mm[3]!) });
        }
        continue;
      }

      // ---- inside an Enum … End Enum --------------------------------------
      if (mode.kind === 'enum') {
        if (/^End\s+Enum\b/i.test(t)) { this.popTo(moduleId); mode = { kind: 'module' }; continue; }
        const mm = t.match(/^\[?([A-Za-z_]\w*)\]?\s*(?:=\s*(.+))?$/);
        if (mm) this.mkNode('enum_member', mm[1]!, l.line, l.endLine, { signature: mm[2]?.trim() });
        continue;
      }

      // ---- module scope ----------------------------------------------------
      // Procedures: Sub / Function
      let m = t.match(/^(?:(Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function)\s+([A-Za-z_]\w*)\s*(\([^)]*\))?(?:\s+As\s+([A-Za-z_][\w.]*(?:\([^)]*\))?))?/i);
      if (m) {
        const vis = visFrom(m[1]);
        const proc = this.mkNode('method', m[3]!, l.line, l.endLine, {
          visibility: vis,
          signature: t,
          returnType: m[5] ? simpleType(m[5]) : undefined,
        });
        if (m[5]) this.addRef(proc.id, simpleType(m[5])!, 'returns', l.line);
        this.stack.push({ id: proc.id, name: proc.name, kind: 'method' });
        mode = { kind: 'proc', end: new RegExp(`^End\\s+${m[2]}\\b`, 'i'), procId: proc.id };
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
        mode = { kind: 'proc', end: /^End\s+Property\b/i, procId: prop.id };
        continue;
      }

      // Declare (external / P-Invoke)
      m = t.match(/^(?:(Public|Private|Global)\s+)?Declare\s+(?:PtrSafe\s+)?(Function|Sub)\s+([A-Za-z_]\w*)\s+Lib\s+"([^"]*)"(?:\s+Alias\s+"([^"]*)")?/i);
      if (m) {
        this.mkNode('method', m[3]!, l.line, l.endLine, {
          visibility: visFrom(m[1]),
          signature: t,
          decorators: ['vb6:declare'],
          docstring: `Declare Lib "${m[4]}"${m[5] ? ` Alias "${m[5]}"` : ''}`,
        });
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
        mode = { kind: 'type', structId: st.id };
        continue;
      }

      // Enum … End Enum
      m = t.match(/^(?:(Public|Private)\s+)?Enum\s+([A-Za-z_]\w*)/i);
      if (m) {
        const en = this.mkNode('enum', m[2]!, l.line, l.endLine, { visibility: visFrom(m[1]) });
        this.stack.push({ id: en.id, name: en.name, kind: 'enum' });
        mode = { kind: 'enum', enumId: en.id };
        continue;
      }

      // Const (module level)
      m = t.match(/^(?:(Public|Private|Global)\s+)?Const\s+(.+)$/i);
      if (m) {
        for (const decl of splitDecls(m[2]!)) {
          const cm = decl.match(/^\[?([A-Za-z_]\w*)\]?\s*(?:As\s+([A-Za-z_][\w.]*))?\s*=\s*(.+)$/i);
          if (cm) this.mkNode('constant', cm[1]!, l.line, l.endLine, {
            visibility: visFrom(m[1]), signature: decl.trim(), returnType: cm[2] ? simpleType(cm[2]) : undefined,
          });
        }
        continue;
      }

      // Module-level variable / field: Public|Private|Global|Dim [WithEvents] decls
      m = t.match(/^(Public|Private|Global|Dim)\s+(WithEvents\s+)?(.+)$/i);
      if (m) {
        const withEvents = !!m[2];
        this.emitVarDecls(m[3]!, l.line, l.endLine, {
          visibility: visFrom(m[1]),
          withEvents,
          asField: true,
        });
        continue;
      }

      // Option / Attribute / VERSION / Begin(designer) / lone End → ignore here.
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
      const typeName = dm[4] ? simpleType(dm[4]) : undefined;
      const decorators: string[] = [];
      if (opts.withEvents) decorators.push('vb6:withevents');
      const node = this.mkNode(opts.asField ? 'field' : 'variable', name, line, endLine, {
        visibility: opts.visibility,
        signature: decl.trim(),
        returnType: typeName,
        decorators: decorators.length ? decorators : undefined,
      });
      if (typeName) this.addRef(node.id, typeName, 'type_of', line);
      if (dm[3]) this.addRef(node.id, typeName!, 'instantiates', line); // As New
    }
  }

  /** Scan one statement line inside a procedure body for calls/refs. */
  private scanStatement(t: string, fromId: string, line: number): void {
    // Local declarations (Dim/Static/Const/ReDim) — capture As New instantiation.
    let m = t.match(/^(?:Dim|Static|Const|ReDim(?:\s+Preserve)?)\s+(WithEvents\s+)?(.+)$/i);
    if (m) {
      for (const decl of splitDecls(m[2]!)) {
        const dm = decl.match(/^([A-Za-z_]\w*)\s*(\([^)]*\))?\s*As\s+(New\s+)?([A-Za-z_][\w.]*)/i);
        if (dm && dm[3]) this.addRef(fromId, simpleType(dm[4]!)!, 'instantiates', line); // As New Foo
        else if (dm && dm[4]) this.addRef(fromId, simpleType(dm[4]!)!, 'type_of', line);
      }
      return;
    }

    // RaiseEvent Name(...)
    m = t.match(/^RaiseEvent\s+([A-Za-z_]\w*)/i);
    if (m) { this.addRef(fromId, m[1]!, 'references', line, { vb6: 'raises_event' }); return; }

    // `Set x = New Foo`  /  `x = New Foo`
    for (const nm of t.matchAll(/\bNew\s+([A-Za-z_][\w.]*)/gi)) {
      this.addRef(fromId, simpleType(nm[1]!)!, 'instantiates', line);
    }

    // `Call Foo` / `Call obj.Bar(...)`
    m = t.match(/^Call\s+(?:[A-Za-z_]\w*\.)*([A-Za-z_]\w*)/i);
    if (m) { this.addRef(fromId, m[1]!, 'calls', line); return; }

    // Explicit `name(` call sites (incl. member calls obj.Method(...)).
    let sawCall = false;
    for (const cm of t.matchAll(/(\.)?\b([A-Za-z_]\w*)\s*\(/g)) {
      const isMember = !!cm[1];
      const name = cm[2]!;
      if (!isMember && (KEYWORDS.has(name.toLowerCase()) || TYPES.has(name.toLowerCase()))) continue;
      this.addRef(fromId, name, 'calls', line, isMember ? { vb6: 'member_call' } : undefined);
      sawCall = true;
    }
    if (sawCall) return;

    // Bare implicit call: `Helper total` / `DoEvents` (no `=`, no `(`).
    if (!/[=(]/.test(t)) {
      const bm = t.match(/^([A-Za-z_]\w*)\b/);
      if (bm && !KEYWORDS.has(bm[1]!.toLowerCase()) && !TYPES.has(bm[1]!.toLowerCase())) {
        this.addRef(fromId, bm[1]!, 'calls', line, { vb6: 'implicit_call' });
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Small pure helpers
// ---------------------------------------------------------------------------

function visFrom(kw?: string): Node['visibility'] | undefined {
  if (!kw) return undefined;
  const k = kw.toLowerCase();
  if (k === 'private') return 'private';
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

    // `Object = "{GUID}#ver#0"; "FILE.ocx"` → OCX component reference from the form.
    for (let i = 0; i < lines.length; i++) {
      const om = lines[i]!.match(/^\s*Object\s*=\s*"?\{[^}]*\}[^;]*;\s*"?([^\s";]+)/i);
      if (om) {
        this.unref(res, formId, om[1]!.replace(/\.(ocx|dll)$/i, ''), i + 1, { vb6: 'ocx_component' });
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
        this.unref(res, node.id, simpleType(typeName), i + 1, isOcx ? { vb6: 'ocx_type' } : { vb6: 'control_type' });
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
        c.decorators = [...(c.decorators || []), 'vb6:controlarray'];
        c.docstring = `Control array element, Index = ${idxM[1]}`;
      }
    }
  }

  private unref(res: ExtractionResult, fromId: string, name: string, line: number, meta: Record<string, unknown>): void {
    res.unresolvedReferences.push({
      fromNodeId: fromId,
      referenceName: name,
      referenceKind: 'references',
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
      nodes.push(proj.node); edges.push(proj.edge);

      for (let i = 0; i < lines.length; i++) {
        const raw = lines[i]!;
        // Member files: Form=..., Module=Name; path, Class=Name; path, UserControl=..., Designer=...
        let m = raw.match(/^\s*(Form|Module|Class|UserControl|Designer|RelatedDoc|UserDocument|PropertyPage)\s*=\s*(.+?)\s*$/i);
        if (m) {
          unresolved.push(ref(proj.node.id, memberName(m[2]!), 'contains', i + 1, this.filePath, { vb6: 'project_member', memberKind: m[1] }));
          continue;
        }
        // Components: `Object={GUID}#maj.min#lcid; FILE.ocx` (early-bound OCX).
        m = raw.match(/^\s*Object\s*=\s*"?\{[^}]*\}[^;]*;\s*"?([^\s";]+)/i);
        if (m) { unresolved.push(ref(proj.node.id, m[1]!.replace(/\.(ocx|dll)$/i, ''), 'imports', i + 1, this.filePath, { vb6: 'ocx_component' })); continue; }
        // COM type-library references: `Reference=*\G{GUID}#maj.min#lcid#path#Description`.
        if (/^\s*Reference\s*=/i.test(raw)) {
          const parts = raw.split('#');
          const guid = raw.match(/\{([0-9A-Fa-f-]+)\}/)?.[1];
          const desc = parts.length >= 2 ? parts[parts.length - 1]!.trim() : '';
          const libPath = parts.length >= 2 ? parts[parts.length - 2]!.trim() : '';
          const name = desc || memberName(libPath) || guid || 'COMReference';
          unresolved.push(ref(proj.node.id, name, 'imports', i + 1, this.filePath, {
            vb6: 'com_reference', ...(guid ? { clsid: guid } : {}), ...(libPath ? { typelib: libPath } : {}),
          }));
        }
      }
    }

    return { nodes, edges, unresolvedReferences: unresolved, errors, durationMs: Date.now() - startTime };
  }
}

function mkFile(filePath: string, endLine: number, id: string): Node {
  return {
    id, kind: 'file', name: filePath.split(/[\\/]/).pop() || filePath, qualifiedName: filePath,
    filePath, language: 'vb6', startLine: 1, endLine, startColumn: 0, endColumn: 0, updatedAt: Date.now(),
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
