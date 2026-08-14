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
| Parameter | `parameter` | — | 62 |
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
| USES_CONTROL / INSTANCE_OF | `type_of` | a control instance IS OF its control type, fixture 48 |
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

CodeGraph's `provenance` column has a closed domain — `'tree-sitter' | 'scip' |
'heuristic'` — and the house convention is that **static extraction leaves it
unset**; only synthesizers mark an edge. VB6 follows that convention rather
than widening the domain, and carries its own detail in `metadata`.

| §14 concept | Representation |
|---|---|
| STATIC | `provenance` unset (the edge came straight from the source) |
| LEXICAL_SCOPE / MODULE_SCOPE / PROJECT_SCOPE | `metadata.scope = 'lexical' \| 'module' \| 'project'` |
| TYPE_RESOLUTION | `metadata.scope = 'type'` |
| EVENT_BINDING / WITHEVENTS_BINDING | `provenance = 'heuristic'` + `metadata.binding` (above) |
| TYPELIB | `metadata.typelib`, `metadata.clsid` |
| SYNTHESIZED | `metadata.synthesizedBy` |
| HEURISTIC | `provenance = 'heuristic'` |
| UNRESOLVED | no edge — the reference stays in `unresolved_refs` |
| directory proximity (orphan files) | `provenance = 'heuristic'`, `metadata.scope = 'directory'` |

**UNRESOLVED is a result, not a failure.** Where VB6 semantics cannot
determine a target — a `Private` procedure seen from another module, a
`CreateObject` ProgID, a call to something that does not exist — the correct
output is no edge at all. Fixtures 41, 50, 51 and 55 enforce that: they forbid
the plausible-looking edge.

The one refinement: when the MEMBER cannot be resolved but the OBJECT it was
used on can, the reference binds to the object (below). That is not a guess —
the object is a symbol we have — and it never names a member we did not find.

## Scope rules the resolver must honour

Name similarity is never sufficient on its own (§9). In decreasing priority,
for an unqualified name inside a procedure:

1. procedure locals and parameters;
2. members of the enclosing class / form / UserControl (fixture 32);
3. `Public` procedures of standard modules in the same project (fixture 31);
4. otherwise unresolved.

The same chain applies when resolving a **qualifier**: locals and parameters of
the calling file first, then `Public` variables of the project's standard
modules (fixture 63). Looking only in the calling file loses every shared
object a VB6 application keeps in a module-level `Public` variable.

Class and form members are **never** reachable unqualified from outside their
type (fixture 31), and `Private`/`Friend` members are not reachable from
another module at all (fixtures 02, 51).

For a qualified name `Qualifier.Member` the target is the member *of the
qualifier*, never the qualifier itself (fixtures 26, 29, 30). The qualifier
may name a module/class/form directly, or a variable whose declared type owns
the member. When that type is `Object` or `Variant` the call is late-bound and
has no static target at all (fixture 43).

### How the qualifier survives to resolution

Resolution reads references back **from the database**, where
`unresolved_refs` has no `metadata` column — so the qualifier cannot travel
there. It travels in `candidates` instead, whose meaning is already "qualified
name this might resolve to": the extractor writes `["c.Compute"]` and the
resolver reads the qualifier back off it. No schema change was needed.

Everything else the resolver needs is derived from the graph rather than
carried: that a reference is a `RaiseEvent` site is known because its target
is a node marked `vb6:event`; that a call is late-bound is known because the
variable's `returnType` is `Object`.

## Member access on a type that is not in the graph

Most of what VB6 code does is touch controls and COM objects — `txtName.Text`,
`rs.Fields`, `MSComm1.Output` — whose types live in type libraries this engine
does not read. The member cannot be resolved, and binding it to a same-named
property found elsewhere in the project would be a fabrication.

What IS known is the object. So the reference attaches to the qualifier and
carries the member name:

```jsonc
kind: 'references',            // never `calls`: it lands on data
metadata: {
  vb6: 'member_on',
  member: 'Text',              // the member that could not be resolved
  qualifier: 'txtName',
  scope: 'object'
}
```

This answers "what touches this control", which is most of what a reader of a
VB6 application wants, without ever claiming to have found the member. The
same applies to late-bound calls (`Dim o As Object`): the edge points at `o`,
never at a guessed target.

### Chains

`Adodc1.Recordset.MoveNext`: the immediate qualifier (`Recordset`) is itself a
member of something external and names no symbol, so the reference falls back
to the **root** of the chain — `Adodc1`, a control that is in the graph. The
extractor records both, in `candidates`, and the resolver tries them in order.
Fixture 64.

Several members of one chain on one line collapse into a single edge: edge
identity includes line and column, by design.

### Duplicate types

Real codebases carry copies of the same class or UserControl in several
folders. When a name resolves to more than one type, candidates are ranked by
the caller's own project first, then any project at all, then files belonging
to no project — a stray copy never wins over the one that ships. Fixture 66.

## Control types name their library

A designer writes a control's type as `Library.Control`
(`MyEditOCX.MyEdit`), and that library is the name of the ActiveX project that
builds the OCX. The library travels as the reference's qualifier, and when
that project is indexed too it decides which of several same-named
UserControls is meant — the client → OCX link of §12. Fixture 61.

## What the extractor deliberately does not emit

- **Bare variable reads inside expressions.** Tracking every local read would
  multiply edges for almost no retrieval value; only members, invocations and
  statement-leading calls are emitted.
- **Members of the runtime objects** (`Debug`, `Err`, `App`, `Screen`,
  `Printer`, `Clipboard`, `Forms`): they target VB6 itself, not the project.
- **Intrinsic types** (`Long`, `String`, `Variant`, …) as reference targets:
  they can never resolve to a project symbol, so emitting them only pollutes
  `unresolved_refs` and the §20 metrics.

Intrinsic *functions* (`CStr`, `MsgBox`, `UBound`, …) **are** emitted as calls
and stay unresolved, because VB6 lets a project redefine them; dropping them
would silently lose a real call.
