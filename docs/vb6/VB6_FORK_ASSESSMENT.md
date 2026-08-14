# VB6 fork assessment

Assessment of the VB6 support implemented in this personal fork
(`github.com/sbuzzrum/codegraph`), against the criteria of §25 of the
specification.

Date: 2026-08-14 (updated after real-world validation).

---

# Verdict: **READY FOR USE**

The engine is complete, conformant, free of regressions, and now validated
against a production VB6 codebase: 2,163 files, 141k nodes, 321k edges, **zero
files that failed to yield symbols**, indexed in under ten seconds.

That validation exposed nine real defects — all of them general, none of them
specific to that codebase. Each became a generic fixture, then a fix, then a
re-validation, in the order §24 prescribes. The conformance suite is 66/66.

Use it, having read `VB6_LIMITATIONS.md` first. The limits are deliberate and
each states what happens instead of guessing — and one of them shapes
expectations enough to repeat here: **member access on standard controls and
COM objects does not resolve**, because those types live in type libraries this
engine does not read. On a real VB6 application that is the majority of the
unresolved set. What you get is a dense, trustworthy map of *structure* — who
calls what, who handles which event, what belongs to which project — not of
data flow through control properties.

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

66 conformance fixtures, all passing: 0 false negatives, 0 false positives, 0
parse errors. The oracle is machine-readable and states
correct VB6 semantics rather than current behaviour, and its `forbid` half
tests specifically for edges that must *not* exist — which is what a name
matcher gets wrong and what §21 ranks above recall.

Sixteen fixtures beyond the specification's 50 cover cases found while building
and while validating: private homonyms, cross-binding of same-named events,
line continuation, strings and comments, array indexing, body ranges, and —
from the real codebase — project-scope isolation, implicit visibility,
cross-module array access, member access on an external type, picking a
UserControl by the OCX library that declares it, parameters as qualifiers,
global object qualifiers, member chains, and a keyword read as a qualifier.

### Regressions — **none**

`npm test`: 164 files, 2966 tests passing, 0 failures. No existing language is
affected (§19).

### Documentation — **complete**

`VB6_SUPPORT.md` (what works, how to use it), `VB6_ARCHITECTURE.md` (how it is
built), `VB6_SEMANTIC_MODEL.md` (what the graph contains, every rule naming the
fixture that pins it), `VB6_LIMITATIONS.md`, `VB6_CONFORMANCE.md` (generated),
`TEST_RESULTS.md`, `OPEN_QUESTIONS.md`, `IMPLEMENTATION_DECISIONS.md`,
`PROGRESS_STATE.md`.

### Reliability of the VB6 graph — **good, and now measured on real code**

The graph reflects VB6 semantics rather than name similarity: visibility is
enforced, names are scoped to their project, qualified calls resolve to the
member and not the qualifier, ambiguity produces no edge, late binding
produces no edge, and synthesized event bindings are marked as synthesized
with their wiring site.

On the production codebase: every file parsed, 12,229 event bindings were
recovered, and three rounds of fixes raised resolution from 28.5% to **71.2%**.
The measure that matters more is **internal call-graph coverage — 82.9%**, or
**92.7%** once the calls VB6 itself cannot reach unqualified are excluded.
What remains unresolved is names with no symbol in the project: VB6 intrinsics,
COM methods, and the type names of standard controls. Full numbers in
`TEST_RESULTS.md`.

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
and plain local reads are not edges. Each is in `VB6_LIMITATIONS.md` with its
consequence and rationale.

---

## What to expect in practice

**Rely on:** the symbol inventory, containment and project membership, call and
reference edges within the indexed source, event flow including `WithEvents`
and OCX handlers, the enforcement of `Private`/project scope, and the absence
of invented edges.

**Do not expect:** members of standard VB controls or COM objects to exist as
symbols (a reference binds to the object, carrying the member name),
conditional-compilation branches to be filtered, or default properties to be
applied. Index the WHOLE tree rather than one `.vbp` at a time — measured:
isolation does not improve resolution and loses the types that sibling ActiveX
projects provide. Files belonging to no indexed project resolve less well, by
design.

**Known open items**, none of them blocking, all in `OPEN_QUESTIONS.md`.
Shipping the standard VB6 control types as built-in nodes was the highest-value
one until member access was made to bind to the object; it is now a
nice-to-have. The agent A/B flow validation the repository expects per language
has not been run for VB6.

## Honesty note on the validation

The codebase used for validation is proprietary and is **not** part of this
fork: it was copied read-only, indexed outside the repository, and only
aggregate metrics are reported. No name, path, snippet or workaround from it
appears anywhere in the code, the fixtures or the documentation — the three
fixtures it motivated are written from scratch with neutral names, and each
would have been just as valid had the defect been found anywhere else.
