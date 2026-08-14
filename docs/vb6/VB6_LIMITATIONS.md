# VB6 support — limitations

What the VB6 frontend does not do, why, and what it costs. A limitation listed
here is a deliberate boundary, not an unknown: each says what happens instead,
so nothing shows up in the graph as a false certainty.

## External type libraries are not read

CodeGraph does not open `.tlb`/`.olb` files or read the registry, so any type
that lives outside the indexed source has no node.

**Consequence.** The types of standard VB controls (`CommandButton`, `Frame`,
`Timer`, `TextBox`) and of external COM objects stay unresolved: the control
instance exists, its `type_of` reference exists, the target does not. COM and
OCX references from `.vbp`/`.frm` *are* nodes and do carry CLSID and type
library path — what is missing is what is inside the library.

**Why.** Type libraries are a Windows-binary format and reading them would tie
indexing to a platform and to whatever happens to be registered on the
machine. A graph that differs by machine is worse than one that is honestly
incomplete.

**When an OCX is indexed as source** (a `.ctl` in the same project or a
sibling project), the link does work end to end: instance → UserControl type →
its public members.

## Conditional compilation is not evaluated

`#If` / `#ElseIf` / `#Else` / `#End If` / `#Const` are recognised as
directives and never mistaken for code, but no branch is removed.

**Consequence.** Symbols and calls from *every* branch are in the graph,
including branches that would not compile in a given configuration. Callers of
a procedure used only in a debug branch will show it.

**Why.** The active configuration lives in the IDE's project settings and on
the build command line, not in the source. Silently keeping one branch would
mean guessing, and dropping all of them would lose real code.

## `Friend` is treated as `Public`

**Consequence.** A `Friend` member is reachable from anywhere in the indexed
project. Since `Friend` means exactly "public within this project", this is
correct for single-project indexing; it is over-permissive only when several
projects are indexed together and one reaches into another's `Friend` member.

## Late binding has no target, by design

A call through a variable declared `As Object` or `As Variant`, or through
`CreateObject`/`GetObject`, produces no `calls` edge. The ProgID of a
`CreateObject("Excel.Application")` is recorded as an unresolved reference so
the dependency is visible.

**Why.** VB6 itself only finds the target at run time. Binding it to a
same-named member of an unrelated class would be a guess presented as a fact.

## Default properties are not applied

VB6 lets `Text1 = "x"` mean `Text1.Text = "x"` through a type's default
member. Default members are declared in type libraries (see above), so they
are not resolved.

**Consequence.** An assignment through a default property is not recorded as a
reference to that property.

## Variable reads inside expressions are not edges

Only members, invocations, statement-leading calls and `With` members produce
references. Reading a local in the middle of an expression does not.

**Why.** Tracking every read would multiply edges for very little retrieval
value. `READS`/`WRITES` in the specification's sense is therefore partial:
property and member access is covered, plain local reads are not.

## `Property Get` and `Property Let` share a name

They are separate nodes, distinguished by `vb6:property-get` / `-let` / `-set`
decorators, but they share a name and a qualified name. When a reference could
mean either, the getter is chosen. A write through `Property Let` may
therefore be attributed to the getter.

## A line scanner, not an AST

VB6 is parsed by a deterministic line/token scanner (decision D1 — no adequate
tree-sitter grammar exists and no local toolchain can build one). Comments,
string literals, line continuations, labels and block terminators are handled,
and the conformance suite covers the constructs the specification lists, but a
scanner has less structural certainty than a real parser on unusual code.

**Mitigation.** Every construct is pinned by a fixture, and adversarial
fixtures exist specifically for the cases where a scanner is most likely to
guess wrong (homonyms, labels, strings, array indexing, late binding).
Anything found in the wild should become a fixture before it becomes a fix.

## Project membership decides scope, and it is only as good as the `.vbp`

A VB6 name is scoped to its project, so resolution filters candidates by the
`.vbp` the calling file belongs to. A file listed in **no** indexed `.vbp` —
orphans, backups, modules kept outside the build — has unknown membership, and
for those files the project filter cannot apply: if several same-named
candidates exist across the index, the reference stays unresolved.

Measured on a real 2,163-file codebase: 79.7% of source files were covered by
`.vbp` membership, and the projects themselves had drifted — they declare 864
member files that no longer exist anywhere in the tree.

For a file in no project the language rule cannot apply, so resolution falls
back to **directory proximity**, under three gates: only when the caller
belongs to no indexed project, only when exactly one candidate sits in the
same directory, and the edge is marked `provenance: 'heuristic'` with
`metadata.scope = 'directory'`. Visibility still applies, so it can never
reach a `Private` procedure of another file.

**Consequence.** Calls from orphan files do resolve, but through a weaker
signal that is labelled as such. Treat those edges as strong hints rather than
parsed facts — they are the one place in the VB6 graph where proximity, not
semantics, chose the target.

**Index the whole tree, not one project at a time.** This was measured, because
the opposite seemed likely: indexing a single `.vbp` in isolation was expected
to remove the ambiguity that duplicate copies of a shared UserControl create.
It does not. Three projects of different sizes were indexed both ways and the
resolution rate was identical to within 0.1 point — the project-scope filter
already discards out-of-project candidates — while isolation *loses* the types
that live in sibling ActiveX projects: on one project, resolved `type_of`
edges fell from 31 to 5 because the UserControls it uses are built by another
`.vbp` in the tree.

The whole tree gives project scope AND the client → OCX link. One project at a
time gives only the first.

## Member access on external types resolves to the OBJECT, not the member

`txtName.Text` cannot resolve `Text` — the type is not in the graph — so the
reference attaches to `txtName` instead, carrying `metadata.member = 'Text'`.

**Consequence.** "What touches this control" works and is accurate. "Who reads
the Text property specifically" does not: you get the member name on the edge,
but there is no node for it, so it cannot be queried as a symbol. The same
holds for late-bound calls, which bind to the variable and never to a guessed
target.

This is a deliberate trade — the alternative was either no edge at all (which
threw away the object link too) or inventing nodes for members we have not
seen. On a real codebase it turned 91,268 references that were previously lost
into usable edges.
