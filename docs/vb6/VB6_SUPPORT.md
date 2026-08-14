# Visual Basic 6 support

CodeGraph indexes Visual Basic 6 / Classic VB projects: standard modules,
class modules, forms, UserControls, and the project files that tie them
together. Nothing extra to install — `.bas`, `.cls`, `.frm`, `.ctl`, `.vbp`
and `.vbg` are picked up like any other supported language.

```bash
codegraph init /path/to/vb6/project      # index
codegraph query   frmMain                # find symbols
codegraph node    ReadBarcode            # source + who calls it + what it calls
codegraph callers Compute
codegraph impact  ItemRead
codegraph explore ReadBarcode ItemRead Pack1_ItemRead
```

The same graph is what the MCP server exposes to an agent.

## What is in the graph

**Symbols.** Standard modules, class modules, forms, UserControls; `Sub`,
`Function`, `Property Get/Let/Set`, `Event` and `Declare` declarations;
`Type…End Type` and its fields; `Enum` and its members; module-level variables
and constants; procedure locals; the controls placed on a form or UserControl,
including control arrays; the project and project-group nodes read from
`.vbp`/`.vbg`; COM type-library and OCX component references, carrying CLSID
and type library path.

**Relations.** Containment (project → file → type → member), calls (qualified
and unqualified), variable and property references, `New`/`As New`
instantiation, `Implements`, control instance → control type, project
membership and references, `RaiseEvent`, and the event bindings below.

Visibility is recorded, and it is enforced: `Private` means private, so a call
from another module does not resolve to it.

## Events

VB6 wires an event to its handler by name — `cmdOk_Click` runs when `cmdOk`
raises `Click`, and nothing in the source says so. CodeGraph reconnects that,
producer → handler, for:

| Producer | Handler | Example |
|---|---|---|
| a control on a form | `Control_Event` | `cmdOk` → `cmdOk_Click` |
| the form or UserControl itself | `Form_Event` | `Form1` → `Form_Load` |
| a `WithEvents` field | `Field_Event` | `mReader` → `mReader_ItemRead` |
| an OCX control | `Ocx_Event` | `MSComm1` → `MSComm1_OnComm` |

For a `WithEvents` field the binding is also drawn from the `Event`
declaration in the producing class, which makes the whole path traversable:

```
ReadBarcode --references (raises_event)--> ItemRead --calls (withevents)--> mReader_ItemRead
```

These edges are **synthesized from a naming convention**, so they are marked
as such: they show up as dynamic hops with the wiring site attached
(`@Form1.frm:28`), never as static calls. A handler is bound only when its
producer actually exists in the same file.

## Members of controls and COM objects

`txtName.Text`, `rs.Fields`, `MSComm1.Output` — the member belongs to a type
that is not in the graph, so it cannot be resolved. The reference binds to the
**object** instead, carrying the member name on the edge:

```
Apply --references--> txtName        (member: "Text")
```

So "what touches this control", "which procedures use this recordset" and
"where is this OCX driven from" all work. What does not work is querying the
member itself as a symbol — there is no node for `TextBox.Text`.

The same applies to late binding: a call through a variable declared
`As Object` binds to the variable, never to a guessed target.

## What "unresolved" means here

Where VB6 semantics do not determine a target, CodeGraph draws **no edge**.
That is deliberate and it is the main difference from a name-matching
approach:

- a `Private` procedure is not reachable from another module;
- an unqualified name never reaches a class or form member from outside it;
- when two same-named `Private` procedures are both out of scope, neither is
  chosen;
- a call through a variable declared `As Object` never reaches a guessed
  member (it binds to the variable, as above);
- a `CreateObject("X.Y")` ProgID is recorded, not resolved to a local type.

A missing edge here means "VB6 itself cannot tell", which is more useful than
a plausible edge that is wrong.

Calls to VB6's own intrinsic functions (`MsgBox`, `CStr`, `Format`, …) also
stay unresolved: a project may legally define its own procedure with that
name, so they are not filtered out on the assumption that they are built-ins.

## Limits worth knowing before you start

External type libraries are not read, so the types of standard VB controls
(`CommandButton`, `Timer`) and of external COM objects resolve to nothing —
the reference is recorded, the target is not in the graph. A UserControl whose
ActiveX project is indexed as source DOES link end to end, even when several
projects define the same control name: the designer names the library, and
that decides. Conditional
compilation is not evaluated: every `#If` branch stays in the graph. `Friend`
is treated as `Public`, which is correct within a single project.

The full list, with what each one costs, is in `VB6_LIMITATIONS.md`.

## Conformance

The behaviour above is pinned by a conformance suite of self-contained VB6
mini-projects, each with a machine-readable oracle — see `VB6_CONFORMANCE.md`
for the current numbers and `__tests__/fixtures/vb6/` for the fixtures. To
reproduce:

```bash
VB6_CONFORMANCE_REPORT=1 npx vitest run __tests__/vb6-conformance.test.ts
```
