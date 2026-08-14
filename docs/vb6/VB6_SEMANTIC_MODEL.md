# VB6 semantic model

What a VB6 codebase turns into inside CodeGraph's graph. Every rule here is
pinned by a fixture in `__tests__/fixtures/vb6/` — the fixture is the
authority, this document explains it.

CodeGraph's `NodeKind` and `EdgeKind` sets are **fixed** (`src/types.ts`). The
VB6 ontology of the specification (§8) is therefore *mapped* onto them and the
VB6-specific detail is carried in `decorators` (`vb6:*` markers) and edge
`metadata`, rather than by widening the core sets. Adding a kind is reserved
for a concept that cannot be expressed at all — none so far.

## Nodes

| VB6 concept | NodeKind | Marker | Fixture |
|---|---|---|---|
| Project (`.vbp`) | `module` | `vb6:project` | 45 |
| Project group (`.vbg`) | `namespace` | `vb6:projectgroup` | 46 |
| Standard module (`.bas`) | `module` | — | 01 |
| Class module (`.cls`) | `class` | — | 04 |
| Form (`.frm`) | `class` | `vb6:form` | 14 |
| UserControl (`.ctl`) | `class` | `vb6:usercontrol` | 19 |
| Control instance (designer) | `field` | `vb6:control` (+ `vb6:ocx`, `vb6:controlarray`) | 15, 18, 23 |
| `Sub` / `Function` | `method` | — | 01 |
| `Property Get/Let/Set` | `property` | `vb6:property-get` / `-let` / `-set` | 06, 07, 08 |
| `Event` declaration | `field` | `vb6:event` | 20 |
| `Declare` (external) | `method` | `vb6:declare` | 49 |
| Module-level variable | `field` | — | 11, 12 |
| Procedure local | `variable` | — | 13 |
| `Const` | `constant` | — | 11 |
| `Enum` / members | `enum` / `enum_member` | — | 09 |
| `Type … End Type` | `struct` (members `field`) | — | 10 |
| COM reference (`Reference=`) | `import` | `vb6:com-reference` | 40 |
| OCX component (`Object=`) | `import` | `vb6:ocx-reference` | 40, 47 |

A `WithEvents` field additionally carries `vb6:withevents` (fixture 22).

## Edges

| VB6 relation | EdgeKind | Notes |
|---|---|---|
| CONTAINS / DECLARES | `contains` | file → container → member; project → member file (from `.vbp`) |
| PROJECT_CONTAINS | `contains` | project node → module/class/form, fixture 45 |
| PROJECT_REFERENCES | `imports` | `.vbg` group → project, fixture 46 |
| CALLS | `calls` | qualified and unqualified calls |
| READS / WRITES | `references` | variable and property access |
| INSTANTIATES | `instantiates` | `New`, `As New` |
| RETURNS | `returns` | function result type (user-defined types only) |
| IMPLEMENTS | `implements` | `Implements IFoo`, fixture 25 |
| USES_CONTROL / INSTANCE_OF | `references` / `type_of` | control instance → control type |
| USES_OCX / REFERENCES_COM | `references` / `imports` | fixtures 40, 47 |
| DECLARES_EVENT | `contains` | class → `Event` node, fixture 20 |
| RAISES_EVENT | `references` | `metadata.vb6 = "raises_event"`, fixture 21 |
| HANDLES_EVENT / HANDLES_OCX_EVENT | `calls` | producer → handler, see below |

### Event bindings

VB6 dispatches events by naming convention, so the edge that connects a
producer to its handler is **synthesized**, not parsed. It is therefore always
recorded as such: `provenance = 'heuristic'`, plus

```jsonc
metadata: {
  synthesizedBy: 'vb6-event-binding',
  binding: 'control_event' | 'form_event' | 'withevents' | 'ocx_event',
  event: 'Click'            // the event name, as written after the underscore
}
```

The edge direction is producer → handler (`calls`), so a call-graph traversal
walks the flow the way it happens at runtime:

| Source | Handler | `binding` | Fixture |
|---|---|---|---|
| control `field` | `Control_Event` | `control_event` | 16, 39, 52 |
| form `class` | `Form_Load` … | `form_event` | 17 |
| `WithEvents` field | `Var_Event` | `withevents` | 22 |
| OCX control `field` | `Ocx_Event` | `ocx_event` | 23 |

For `WithEvents` the binding is emitted **twice** when the producing class is
indexed: once from the field (always possible) and once from the `Event`
declaration node itself, which is what makes the end-to-end path in the
specification (§13) traversable:

```
ReadBarcode --references(raises_event)--> ItemRead --calls(withevents)--> mReader_ItemRead
```

## Provenance

Every non-syntactic relation records how it was obtained. The specification's
vocabulary (§14) maps onto CodeGraph's `edges.provenance` column plus edge
metadata:

| §14 concept | Representation |
|---|---|
| STATIC | `provenance = 'static'` |
| LEXICAL_SCOPE / MODULE_SCOPE / PROJECT_SCOPE | `provenance = 'static'`, `metadata.scope = 'lexical' \| 'module' \| 'project'` |
| TYPE_RESOLUTION | `metadata.scope = 'type'` |
| EVENT_BINDING / WITHEVENTS_BINDING | `provenance = 'heuristic'` + `metadata.binding` (above) |
| TYPELIB | `metadata.typelib`, `metadata.clsid` |
| SYNTHESIZED | `metadata.synthesizedBy` |
| HEURISTIC | `provenance = 'heuristic'` |
| UNRESOLVED | no edge — the reference stays in `unresolved_refs` |

**UNRESOLVED is a result, not a failure.** Where VB6 semantics cannot
determine a target — a `Private` procedure seen from another module, a
late-bound `Object` member, a `CreateObject` ProgID — the correct output is no
edge at all. Fixtures 41, 43, 50, 51 and 55 exist to enforce that: they forbid
the plausible-looking edge.

## Scope rules the resolver must honour

Name similarity is never sufficient on its own (§9). In decreasing priority,
for an unqualified name inside a procedure:

1. procedure locals and parameters;
2. members of the enclosing class / form / UserControl (fixture 32);
3. `Public` procedures of standard modules in the same project (fixture 31);
4. otherwise unresolved.

Class and form members are **never** reachable unqualified from outside their
type (fixture 31), and `Private`/`Friend` members are not reachable from
another module at all (fixtures 02, 51).

For a qualified name `Qualifier.Member` the target is the member *of the
qualifier*, never the qualifier itself (fixtures 26, 29, 30).
