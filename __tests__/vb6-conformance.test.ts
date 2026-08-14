/**
 * VB6 conformance suite — the independent oracle for the VB6 frontend.
 *
 * Each fixture under `__tests__/fixtures/vb6/NN_name/` is a self-contained
 * mini VB6 project (`.vbp` + sources) plus an `expected.json` oracle. The
 * runner indexes the fixture through the real pipeline (extraction →
 * resolution → SQLite) and compares the produced graph against the oracle.
 *
 * The oracle states the CORRECT VB6 semantics, not today's behaviour. A
 * fixture whose semantics the engine does not implement yet is marked
 * `"status": "known-gap"` with a `"gap"` reason: the runner then asserts the
 * gap is STILL there, so the suite stays green while tracking exactly what is
 * missing — and fails loudly ("promote to pass") the moment a gap is fixed.
 * That is what keeps the metrics in docs/vb6/VB6_CONFORMANCE.md reproducible.
 *
 * Matching is subset-based: every `expect` entry must find at least one match
 * in the graph, every `forbid` entry must find none. Omitted fields are
 * wildcards, so a fixture only pins what it is actually about.
 *
 * Set `VB6_CONFORMANCE_REPORT=1` to (re)write docs/vb6/VB6_CONFORMANCE.md.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as os from 'node:os';
import { CodeGraph } from '../src';

const FIXTURES_DIR = path.join(__dirname, 'fixtures', 'vb6');

// ---------------------------------------------------------------------------
// Oracle schema
// ---------------------------------------------------------------------------

interface NodeExpect {
  kind?: string;
  name?: string;
  file?: string;
  visibility?: string;
  returnType?: string;
  qualifiedName?: string;
  /** Every listed decorator must be present on the node. */
  decorators?: string[];
  /**
   * Exact number of matches required. Omitted means "at least one" — set it
   * when the count itself is the semantics (a control array is ONE symbol with
   * many indices, not N duplicate symbols).
   */
  count?: number;
}

interface EdgeExpect {
  kind?: string;
  from?: string;
  to?: string;
  fromFile?: string;
  toFile?: string;
  fromKind?: string;
  toKind?: string;
  provenance?: string;
  /** Every listed key must match the edge's JSON metadata. */
  metadata?: Record<string, unknown>;
  /** Exact number of matching edges required. Omitted means "at least one". */
  count?: number;
}

interface UnresolvedExpect {
  name?: string;
  kind?: string;
  from?: string;
}

interface Oracle {
  description: string;
  status: 'pass' | 'known-gap';
  gap?: string;
  expect?: { nodes?: NodeExpect[]; edges?: EdgeExpect[]; unresolved?: UnresolvedExpect[] };
  forbid?: { nodes?: NodeExpect[]; edges?: EdgeExpect[]; unresolved?: UnresolvedExpect[] };
}

// ---------------------------------------------------------------------------
// Graph snapshot
// ---------------------------------------------------------------------------

interface GraphNode {
  id: string;
  kind: string;
  name: string;
  qualifiedName: string;
  file: string;
  visibility: string | null;
  returnType: string | null;
  decorators: string[];
}

interface GraphEdge {
  kind: string;
  from: string;
  to: string;
  fromFile: string;
  toFile: string;
  fromKind: string;
  toKind: string;
  provenance: string | null;
  metadata: Record<string, unknown>;
}

interface GraphUnresolved {
  name: string;
  kind: string;
  from: string;
  status: string;
}

interface Snapshot {
  nodes: GraphNode[];
  edges: GraphEdge[];
  unresolved: GraphUnresolved[];
  parseErrors: number;
}

function parseJson(raw: unknown): any {
  if (typeof raw !== 'string' || raw === '') return undefined;
  try {
    return JSON.parse(raw);
  } catch {
    return undefined;
  }
}

/** Copy a fixture into a temp dir and index it through the real pipeline. */
async function indexFixture(fixtureDir: string, tmpDir: string): Promise<Snapshot> {
  for (const entry of fs.readdirSync(fixtureDir)) {
    if (entry === 'expected.json') continue;
    fs.cpSync(path.join(fixtureDir, entry), path.join(tmpDir, entry), { recursive: true });
  }

  const cg = await CodeGraph.init(tmpDir, { silent: true });
  const result = await cg.indexAll();
  const db = (cg as any).db.db;

  const nodes: GraphNode[] = db
    .prepare(
      `SELECT id, kind, name, qualified_name qn, file_path file, visibility, return_type rt, decorators
       FROM nodes WHERE language = 'vb6'`
    )
    .all()
    .map((r: any) => ({
      id: r.id,
      kind: r.kind,
      name: r.name,
      qualifiedName: r.qn,
      file: path.basename(r.file),
      visibility: r.visibility ?? null,
      returnType: r.rt ?? null,
      decorators: parseJson(r.decorators) ?? [],
    }));

  const edges: GraphEdge[] = db
    .prepare(
      `SELECT e.kind, s.name sname, s.kind skind, s.file_path sfile,
              t.name tname, t.kind tkind, t.file_path tfile,
              e.provenance, e.metadata
       FROM edges e
       JOIN nodes s ON s.id = e.source
       JOIN nodes t ON t.id = e.target
       WHERE s.language = 'vb6' OR t.language = 'vb6'`
    )
    .all()
    .map((r: any) => ({
      kind: r.kind,
      from: r.sname,
      to: r.tname,
      fromFile: path.basename(r.sfile),
      toFile: path.basename(r.tfile),
      fromKind: r.skind,
      toKind: r.tkind,
      provenance: r.provenance ?? null,
      metadata: parseJson(r.metadata) ?? {},
    }));

  const unresolved: GraphUnresolved[] = db
    .prepare(
      `SELECT u.reference_name name, u.reference_kind kind, u.status, n.name fromName
       FROM unresolved_refs u
       LEFT JOIN nodes n ON n.id = u.from_node_id
       WHERE u.language = 'vb6'`
    )
    .all()
    .map((r: any) => ({ name: r.name, kind: r.kind, from: r.fromName ?? '', status: r.status }));

  cg.destroy();

  const parseErrors = (result.errors ?? []).filter((e: any) => e.severity === 'error').length;
  return { nodes, edges, unresolved, parseErrors };
}

// ---------------------------------------------------------------------------
// Matching
// ---------------------------------------------------------------------------

function nodeMatches(n: GraphNode, e: NodeExpect): boolean {
  if (e.kind !== undefined && n.kind !== e.kind) return false;
  if (e.name !== undefined && n.name !== e.name) return false;
  if (e.file !== undefined && n.file !== e.file) return false;
  if (e.visibility !== undefined && n.visibility !== e.visibility) return false;
  if (e.returnType !== undefined && n.returnType !== e.returnType) return false;
  if (e.qualifiedName !== undefined && n.qualifiedName !== e.qualifiedName) return false;
  if (e.decorators !== undefined && !e.decorators.every((d) => n.decorators.includes(d))) return false;
  return true;
}

function edgeMatches(g: GraphEdge, e: EdgeExpect): boolean {
  if (e.kind !== undefined && g.kind !== e.kind) return false;
  if (e.from !== undefined && g.from !== e.from) return false;
  if (e.to !== undefined && g.to !== e.to) return false;
  if (e.fromFile !== undefined && g.fromFile !== e.fromFile) return false;
  if (e.toFile !== undefined && g.toFile !== e.toFile) return false;
  if (e.fromKind !== undefined && g.fromKind !== e.fromKind) return false;
  if (e.toKind !== undefined && g.toKind !== e.toKind) return false;
  if (e.provenance !== undefined && g.provenance !== e.provenance) return false;
  if (e.metadata !== undefined) {
    for (const [k, v] of Object.entries(e.metadata)) {
      if (JSON.stringify(g.metadata[k]) !== JSON.stringify(v)) return false;
    }
  }
  return true;
}

function unresolvedMatches(u: GraphUnresolved, e: UnresolvedExpect): boolean {
  if (e.name !== undefined && u.name !== e.name) return false;
  if (e.kind !== undefined && u.kind !== e.kind) return false;
  if (e.from !== undefined && u.from !== e.from) return false;
  return true;
}

interface Evaluation {
  /** Expected things the graph is missing (false negatives). */
  missing: string[];
  /** Forbidden things the graph contains (false positives). */
  violations: string[];
  expectedCount: number;
  foundCount: number;
}

function evaluate(oracle: Oracle, snap: Snapshot): Evaluation {
  const missing: string[] = [];
  const violations: string[] = [];
  let expectedCount = 0;
  let foundCount = 0;

  for (const e of oracle.expect?.nodes ?? []) {
    expectedCount++;
    const hits = snap.nodes.filter((n) => nodeMatches(n, e)).length;
    const ok = e.count !== undefined ? hits === e.count : hits > 0;
    if (ok) foundCount++;
    else missing.push(`node ${JSON.stringify(e)} — found ${hits}`);
  }
  for (const e of oracle.expect?.edges ?? []) {
    expectedCount++;
    const hits = snap.edges.filter((g) => edgeMatches(g, e)).length;
    const ok = e.count !== undefined ? hits === e.count : hits > 0;
    if (ok) foundCount++;
    else missing.push(`edge ${JSON.stringify(e)} — found ${hits}`);
  }
  for (const e of oracle.expect?.unresolved ?? []) {
    expectedCount++;
    if (snap.unresolved.some((u) => unresolvedMatches(u, e))) foundCount++;
    else missing.push(`unresolved ${JSON.stringify(e)}`);
  }

  for (const e of oracle.forbid?.nodes ?? []) {
    const hit = snap.nodes.find((n) => nodeMatches(n, e));
    if (hit) violations.push(`forbidden node ${JSON.stringify(e)} — found ${hit.kind} ${hit.qualifiedName}`);
  }
  for (const e of oracle.forbid?.edges ?? []) {
    const hit = snap.edges.find((g) => edgeMatches(g, e));
    if (hit) violations.push(`forbidden edge ${JSON.stringify(e)} — found ${hit.from} --${hit.kind}--> ${hit.to}`);
  }
  for (const e of oracle.forbid?.unresolved ?? []) {
    const hit = snap.unresolved.find((u) => unresolvedMatches(u, e));
    if (hit) violations.push(`forbidden unresolved ${JSON.stringify(e)} — found ${hit.name} [${hit.kind}]`);
  }

  return { missing, violations, expectedCount, foundCount };
}

// ---------------------------------------------------------------------------
// Suite
// ---------------------------------------------------------------------------

interface Row {
  fixture: string;
  status: Oracle['status'];
  gap?: string;
  expected: number;
  found: number;
  missing: number;
  violations: number;
  unresolved: number;
  parseErrors: number;
  outcome: 'pass' | 'fail' | 'known-gap' | 'ready-to-promote';
}

const fixtureNames = fs.existsSync(FIXTURES_DIR)
  ? fs
      .readdirSync(FIXTURES_DIR, { withFileTypes: true })
      .filter((d) => d.isDirectory())
      .map((d) => d.name)
      .sort()
  : [];

const rows: Row[] = [];

describe('VB6 conformance suite', () => {
  let tmpRoot: string;

  beforeAll(() => {
    tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'vb6-conformance-'));
  });

  afterAll(() => {
    fs.rmSync(tmpRoot, { recursive: true, force: true });
    if (process.env.VB6_CONFORMANCE_REPORT === '1') writeReport(rows);
  });

  it('has fixtures', () => {
    expect(fixtureNames.length).toBeGreaterThan(0);
  });

  for (const name of fixtureNames) {
    const fixtureDir = path.join(FIXTURES_DIR, name);
    const oracle: Oracle = JSON.parse(fs.readFileSync(path.join(fixtureDir, 'expected.json'), 'utf8'));

    it(`${name} — ${oracle.description}`, async () => {
      const tmpDir = fs.mkdtempSync(path.join(tmpRoot, `${name}-`));
      const snap = await indexFixture(fixtureDir, tmpDir);
      const ev = evaluate(oracle, snap);
      const clean = ev.missing.length === 0 && ev.violations.length === 0;

      rows.push({
        fixture: name,
        status: oracle.status,
        gap: oracle.gap,
        expected: ev.expectedCount,
        found: ev.foundCount,
        missing: ev.missing.length,
        violations: ev.violations.length,
        unresolved: snap.unresolved.length,
        parseErrors: snap.parseErrors,
        outcome: oracle.status === 'pass' ? (clean ? 'pass' : 'fail') : clean ? 'ready-to-promote' : 'known-gap',
      });

      expect(snap.parseErrors, 'fixture must parse without errors').toBe(0);

      if (oracle.status === 'pass') {
        const report = [...ev.missing, ...ev.violations];
        expect(report, `${name} does not match its oracle`).toEqual([]);
        return;
      }

      // known-gap: the gap must still be there. If it isn't, the engine
      // improved — promote the fixture to "pass" so it guards the new behaviour.
      expect(
        clean,
        `${name} is marked known-gap ("${oracle.gap}") but now matches its oracle fully — ` +
          `set "status": "pass" in ${path.relative(process.cwd(), path.join(fixtureDir, 'expected.json'))}`
      ).toBe(false);
    }, 60_000);
  }
});

/** Write the reproducible metrics table required by the prompt (§20). */
function writeReport(all: Row[]): void {
  const sorted = [...all].sort((a, b) => a.fixture.localeCompare(b.fixture));
  const total = sorted.length;
  const passing = sorted.filter((r) => r.outcome === 'pass').length;
  const gaps = sorted.filter((r) => r.outcome === 'known-gap').length;
  const failing = sorted.filter((r) => r.outcome === 'fail').length;
  const promote = sorted.filter((r) => r.outcome === 'ready-to-promote').length;
  const expected = sorted.reduce((s, r) => s + r.expected, 0);
  const found = sorted.reduce((s, r) => s + r.found, 0);
  const fn = sorted.reduce((s, r) => s + r.missing, 0);
  const fp = sorted.reduce((s, r) => s + r.violations, 0);
  const unresolved = sorted.reduce((s, r) => s + r.unresolved, 0);
  const parseErrors = sorted.reduce((s, r) => s + r.parseErrors, 0);

  const lines: string[] = [
    '# VB6 conformance results',
    '',
    '> Generated by `VB6_CONFORMANCE_REPORT=1 npx vitest run __tests__/vb6-conformance.test.ts`.',
    '> Do not edit by hand.',
    '',
    '## Totals',
    '',
    '| Metric | Value |',
    '|---|---|',
    `| Fixtures | ${total} |`,
    `| Passing | ${passing} |`,
    `| Known gaps (semantics not implemented yet) | ${gaps} |`,
    `| Failing (regression — oracle says pass) | ${failing} |`,
    `| Ready to promote (gap closed) | ${promote} |`,
    `| Expected assertions | ${expected} |`,
    `| Assertions satisfied | ${found} |`,
    `| False negatives (expected, missing) | ${fn} |`,
    `| False positives (forbidden, present) | ${fp} |`,
    `| Unresolved references | ${unresolved} |`,
    `| Parse errors | ${parseErrors} |`,
    '',
    '## Per fixture',
    '',
    '| Fixture | Outcome | Expected | Found | FN | FP | Unresolved | Gap |',
    '|---|---|---|---|---|---|---|---|',
    ...sorted.map(
      (r) =>
        `| ${r.fixture} | ${r.outcome} | ${r.expected} | ${r.found} | ${r.missing} | ${r.violations} | ${r.unresolved} | ${r.gap ?? ''} |`
    ),
    '',
  ];

  const out = path.join(__dirname, '..', 'docs', 'vb6', 'VB6_CONFORMANCE.md');
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, lines.join('\n'));
}
