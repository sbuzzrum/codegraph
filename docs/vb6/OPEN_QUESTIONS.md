# VB6 support — open questions

Decisions that are made but not settled, and things that need evidence from a
real codebase before they can be answered. Distinct from
`VB6_LIMITATIONS.md`, which lists boundaries that are deliberate and closed.

## 1. How much does the missing type-library layer actually cost?

Standard VB control types and external COM types are not in the graph, so
`Timer`, `CommandButton`, `ADODB.Connection` and friends stay unresolved. On
the fixtures this is a handful of references; on a real application that leans
on third-party OCX controls it could be the majority of the unresolved set.

**What would answer it:** the unresolved breakdown from indexing a real
project. If external types dominate, reading type libraries — or shipping a
small table of the standard VB6 controls — becomes worth its cost.

## 2. Should the standard VB6 control types ship as built-in nodes?

A cheap subset of the above: `VB.CommandButton`, `VB.TextBox`, `VB.Timer`, …
are the same in every VB6 installation and could be synthesised as external
symbols so control instances have a target.

**Against:** they are not code in the project; inventing nodes for them
inflates the graph and blurs "indexed" with "known".
**For:** it removes the most common unresolved category and lets an agent see
that `tmrTick` is a Timer.

Undecided; needs (1) first.

## 3. Multi-project resolution across a `.vbg`

A project group is indexed and its projects are linked, but resolution does
not use project membership to decide scope: a `Public` procedure of a standard
module is global to everything indexed, not to *its project*. With two
projects in one index that share a module name, the wrong one can win.

**What would answer it:** whether real groups actually contain same-named
public procedures in sibling projects. If they do, resolution has to filter
candidates by the project that contains the calling file, which means walking
`.vbp` membership at resolution time.

## 4. `Property Let` versus `Property Get` on a write

A reference that could be either currently resolves to the getter. Telling
them apart needs the assignment context, which the extractor knows and does
not currently record.

**Cost of fixing:** one more distinction carried through to resolution.
**Cost of not fixing:** a write shows as a read in callers/impact.

Worth doing if property-heavy code proves common.

## 5. Are procedure locals worth their node count?

Locals became nodes to satisfy the semantic model (§8 lists `Variable`). They
are numerous and rarely queried; on a large codebase they may be a significant
share of all nodes with little retrieval value.

**What would answer it:** the node-count breakdown on a real project, and
whether any query actually reaches them.

## 6. Does the event-binding heuristic hold up at scale?

A handler binds only when a producer of that name exists in the same file,
which is precise on the fixtures. Unknown: how often real code has a procedure
that merely *looks* like a handler (`Data_Refresh` where `Data` is a control
that does not raise `Refresh`), and how often a genuine handler is missed
because its producer is declared elsewhere.

**What would answer it:** a precision spot-check of synthesized edges on a
real project, the way the dynamic-dispatch playbook prescribes for every other
synthesizer.

## 7. Should VB6 get a `codegraph_explore` flow test?

Every language/framework in the repo is expected to be validated with the
agent A/B methodology (small/medium/large repos, ≥3 flow prompts). VB6 has
none: the canonical flow would be *user action → event handler → business
logic → data access*. This needs a real VB6 repo of each size, which is the
same blocker as (1).
