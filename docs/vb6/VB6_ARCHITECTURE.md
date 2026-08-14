# VB6 support — architecture

How the VB6 frontend is put together and where it hooks into CodeGraph. For
*what* the graph contains see `VB6_SEMANTIC_MODEL.md`; for what it cannot do
see `VB6_LIMITATIONS.md`.

## Shape

```
.bas .cls .frm .ctl .vbp .vbg
        │
        ▼
extractFromSource            src/extraction/tree-sitter.ts   ← language dispatch
        │
        ▼
Vb6Extractor / Vb6FormExtractor / Vb6ProjectExtractor
        │                    src/extraction/vb6-extractor.ts
        ▼
nodes + edges + unresolved references  →  SQLite
        │
        ▼
Vb6Resolver                  src/resolution/vb6-resolver.ts  ← scope rules
        │
        ▼
synthesizeVb6EventBindings   src/resolution/vb6-event-synthesizer.ts
        │
        ▼
graph → search / callers / callees / impact / explore / MCP
```

Nothing here is a parallel subsystem: extraction, resolution and synthesis are
the pipeline's own stages, entered the same way every other language enters
them.

## No tree-sitter grammar

VB6 is parsed by a deterministic line/token scanner rather than a grammar.
That is a measured decision, recorded as D1 in `IMPLEMENTATION_DECISIONS.md`:
the bundled VB.NET grammar errors on VB6-*core* constructs (`Property
Get/Let/Set`, `Event`, `WithEvents`, `Implements`, `Type…End Type`, `As New`,
`Declare`), and no local toolchain can build a VBA grammar to wasm.

The language registration reflects that: `vb6` is in `LANGUAGES` and in the
extension map of `grammars.ts`, has **no** wasm entry and **no** entry in the
tree-sitter `EXTRACTORS` map, and is dispatched by its own branch in
`extractFromSource` — the same route `svelte`, `vue` and `dfm` take.

The approach suits the language: VB6 is statement-oriented, keyword-led, with
explicit block terminators and `_` line continuation.

## Extraction

`src/extraction/vb6-extractor.ts` holds three extractors behind one entry
point, `extractVb6File`:

| Extractor | Files | Produces |
|---|---|---|
| `Vb6Extractor` | `.bas` `.cls` (and the code body of `.frm`/`.ctl`) | module/class node, procedures, properties, events, types, enums, constants, variables, references |
| `Vb6FormExtractor` | `.frm` `.ctl` | the form/UserControl plus its designer control tree; delegates the code body to `Vb6Extractor` |
| `Vb6ProjectExtractor` | `.vbp` `.vbg` | the project (or project group) node, its member files, its COM/OCX references |

### Two passes over a file

Declarations are collected first and statements are scanned afterwards, once
the whole file is known. This is not an optimisation — it is required for
correctness. VB6 spells an array access and a call identically (`Items(3)` /
`Helper(3)`), and the only thing that distinguishes them is whether the name
was declared as data. Declarations may appear *after* the procedures that use
them, so a single forward pass cannot decide.

### Statement scanning

The scanner works from VB6 syntax, not from the shape of a line:

- a dotted chain `A.B` targets the member `B`; the qualifier `A` is never a
  target, and travels with the reference so resolution can use it;
- `.Member` inside a `With` block is attributed to the block's target;
- an assignment is `name =` (never `:=`), which is what lets a call with named
  arguments be recognised as a call;
- string literals are blanked before scanning, so an identifier inside a
  literal is not code;
- line labels are jump targets;
- members of the VB6 runtime objects (`Debug`, `Err`, `App`, …) and intrinsic
  types are not project symbols and are never emitted.

## Resolution

`src/resolution/vb6-resolver.ts` is reached from `resolveOne` and has **no
fallthrough**: if the VB6 rules do not name a single target, the reference
stays unresolved. Falling back to the generic name matcher would reintroduce
precisely the bindings the resolver exists to prevent — a `Private` procedure
reached from another module, or one of two homonyms picked at random.

Scope order for an unqualified name: the enclosing module/class/form, then
`Public` procedures of standard modules (`.bas`). A qualified name resolves
against its qualifier — a module/class of that name, or the declared type of
the variable named. `Object`/`Variant` means late-bound, which has no static
target.

### How context survives to resolution

Resolution reads references back from SQLite, and `unresolved_refs` has no
`metadata` column. Rather than migrate the shared schema, the qualifier
travels in **`candidates`**, whose meaning is already "qualified name this may
resolve to". Everything else is derived from the graph: a `RaiseEvent` site is
recognised because its target is a node marked `vb6:event`; late binding
because the variable's `returnType` is `Object`.

`ResolvedRef` gained an optional `metadata` so a language resolver can record
how it decided; it is merged into the edge under the standard keys.

## Event bindings

`src/resolution/vb6-event-synthesizer.ts` runs after resolution, alongside the
other dynamic-dispatch synthesizers. VB6 wires events by naming convention
(`cmdOk_Click`), so the producer → handler hop exists nowhere in the source.

A handler binds only when its producer really exists in the same file, which
is what keeps `IWorker_Run` (an `Implements` method) and two controls sharing
an event name from being wired to each other. Names with several underscores
are tried at every split, and the split whose producer exists wins.

Every edge it creates is `provenance: 'heuristic'` with `synthesizedBy`,
`binding`, `event` and `registeredAt` in metadata — surfaced by
`codegraph_explore` and the `codegraph_node` trail as a dynamic hop with its
wiring site, never as a static call.

## Files

| File | Role |
|---|---|
| `src/types.ts` | `vb6` in `LANGUAGES` |
| `src/extraction/grammars.ts` | extension map, display name, supported/grammar-loaded lists |
| `src/extraction/tree-sitter.ts` | dispatch branch in `extractFromSource` |
| `src/extraction/vb6-extractor.ts` | the three extractors |
| `src/resolution/vb6-resolver.ts` | scope rules |
| `src/resolution/vb6-event-synthesizer.ts` | event bindings |
| `src/resolution/index.ts` | dispatch in `resolveOne`, synthesis hook, `candidates` propagation |
| `src/mcp/tools.ts` | renders a VB6 event binding as a dynamic hop |
| `__tests__/vb6-conformance.test.ts` | the oracle runner |
| `__tests__/fixtures/vb6/` | the fixtures |
