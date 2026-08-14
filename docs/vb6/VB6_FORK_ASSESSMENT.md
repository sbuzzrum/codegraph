# VB6 fork assessment

Assessment of the VB6 support implemented in this personal fork
(`github.com/sbuzzrum/codegraph`), against the criteria of §25 of the
specification.

Date: 2026-08-14.

---

# Verdict: **READY FOR USE WITH RESERVATIONS**

The engine is complete, conformant and free of regressions, and it can be used
to index VB6 projects today. The reservation is not about a missing feature:
it is that **the engine has never been run against a large real VB6 codebase**,
so its behaviour at that scale is unmeasured. That is the one thing the
fixtures cannot tell us, and the specification itself (§24) puts real-world
validation after — not instead of — the conformance suite.

Read `VB6_LIMITATIONS.md` before relying on the graph; the limits are
deliberate and each says what happens instead.

---

## Criterion by criterion

### Generality of the code — **good**

Nothing in the implementation is specific to any one codebase. The extractors
work from the VB6 language and file formats; the resolver from the language's
scope rules; the event synthesizer from the naming convention VB6 itself uses.
No customer name, path or workaround appears anywhere in the code, the tests
or the documentation (§23 verified by search: zero occurrences).

The fixtures are written from the specification's list of constructs, not from
any existing project.

### Consistency with CodeGraph's architecture — **good**

The support enters the pipeline where every other language does:

- one branch in `extractFromSource`, like `svelte`/`vue`/`dfm`;
- one branch in `resolveOne`, like the CFML component-path resolver;
- one synthesis pass next to the other dynamic-dispatch synthesizers.

No parallel subsystem (§18). The fixed `NodeKind`/`EdgeKind` sets were **not**
extended: the VB6 ontology of §8 is mapped onto them with `vb6:*` decorators
and edge metadata (D2). The `provenance` domain was not widened either — VB6
follows the house convention that static extraction leaves it unset and only
synthesizers mark an edge.

Three edits touch shared code, all additive and all justified:
`candidates` propagation (a field that was written and read but dropped before
reaching any resolver — a latent bug, fixed), an optional `metadata` on
`ResolvedRef`, and a rendering branch for VB6 event bindings in the MCP tools.

### Licence compatibility — **no exposure**

No third-party code was vendored, no grammar was forked, no dependency was
added: `package.json` and `package-lock.json` are untouched. Everything is
original code in the fork, under the repository's existing licence. ProLeap
was used only as a conceptual reference for VB6 semantics, never copied and
never at runtime.

### Automated tests — **good**

56 conformance fixtures, 132 assertions, all passing: 0 false negatives, 0
false positives, 0 parse errors. The oracle is machine-readable and states
correct VB6 semantics rather than current behaviour, and its `forbid` half
tests specifically for edges that must *not* exist — which is what a name
matcher gets wrong and what §21 ranks above recall.

Six fixtures beyond the specification's 50 cover cases found while building:
private homonyms, cross-binding of same-named events, line continuation,
strings and comments, array indexing, and body ranges.

### Regressions — **none**

`npm test`: 164 files, 2958 tests passing, 0 failures. No existing language is
affected (§19).

### Documentation — **complete**

`VB6_SUPPORT.md` (what works, how to use it), `VB6_ARCHITECTURE.md` (how it is
built), `VB6_SEMANTIC_MODEL.md` (what the graph contains, every rule naming the
fixture that pins it), `VB6_LIMITATIONS.md`, `VB6_CONFORMANCE.md` (generated),
`TEST_RESULTS.md`, `OPEN_QUESTIONS.md`, `IMPLEMENTATION_DECISIONS.md`,
`PROGRESS_STATE.md`.

### Reliability of the VB6 graph — **good on what is covered**

The graph reflects VB6 semantics rather than name similarity: visibility is
enforced, qualified calls resolve to the member and not the qualifier,
ambiguity produces no edge, late binding produces no edge, and synthesized
event bindings are marked as synthesized with their wiring site.

The measurable effect: on the sample project, unresolved references fell from
44 to 9 as the extractor and resolver were fixed, and the nine that remain are
external by nature.

### MCP integration quality — **good**

`query`, `callers`, `callees`, `impact`, `node` and `explore` all work on the
VB6 graph. `node` returns whole procedure bodies with the caller/callee trail;
`explore` leads with the event bindings as dynamic hops, each labelled and
carrying its wiring site (`@Form1.frm:28`), and returns verbatim source. VB6 is
named in the MCP server instructions so an agent does not assume it is
uncovered.

### Residual limitations — **known and documented**

External type libraries are not read (so standard control and external COM
types have no node), conditional compilation is not evaluated (all branches
stay), `Friend` is treated as `Public`, default properties are not applied,
plain local reads are not edges, and parameters are not nodes. Each is in
`VB6_LIMITATIONS.md` with its consequence and rationale.

---

## What the reservation means in practice

**You can** index a VB6 project and rely on: the symbol inventory, containment
and project membership, call and reference edges within the indexed source,
event flow including `WithEvents` and OCX handlers, and the absence of
invented edges.

**You should not yet** assume: that every file in a decades-old codebase parses
cleanly, that indexing time and memory are acceptable at scale, or that the
resolution rate measured here (fixtures) carries over to production code.

**To lift the reservation** — the sequence §24 prescribes:

1. index a real VB6 repository;
2. record parse errors, unresolved breakdown by category, node/edge counts and
   wall-clock time;
3. for anything that turns out wrong, write a minimal fixture first, fix the
   engine, re-run the suite, then re-validate the repository — never patch for
   the repository;
4. spot-check the precision of synthesized event edges, as the dynamic-dispatch
   playbook requires of every synthesizer;
5. run the agent A/B flow validation on small, medium and large VB6 repos.

If those pass without engine changes, the verdict becomes READY FOR USE. If
they surface general defects, the fixtures grow and the engine is fixed — the
process the whole build has followed so far.
