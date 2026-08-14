# VB6 support — open questions

Decisions that are made but not settled, and things that need evidence from a
real codebase before they can be answered. Distinct from
`VB6_LIMITATIONS.md`, which lists boundaries that are deliberate and closed.

## 1. How much does the missing type-library layer actually cost? — **answered**

Measured on a real 2,163-file codebase: **62% of all unresolved references are
member access on an object whose type is not in the graph** (`.Text`, `.Value`,
`.Caption` on controls; `.Fields`, `.Recordset` on COM objects), plus another
31k `type_of` references naming standard control types.

So yes: external types dominate the unresolved set, decisively.

**What was done about it:** rather than importing those types, the reference
now binds to the object the member was used on. That lifted the resolution
rate from 35.7% to 55.1% and made "what touches this control" answerable,
which is the question the type-library data was mostly wanted for. See
question 2 for what remains genuinely missing.

## 2. Should the standard VB6 control types ship as built-in nodes? — **much less urgent**

This was the highest-value item until member access was made to bind to the
OBJECT rather than being dropped. That change recovered 91,268 references and
answered the question people actually ask — *what touches this control* —
without inventing a single node.

What a built-in table would still add: the ability to query a control's
members as symbols (`who reads TextBox.Text`), and a target for the ~31k
`type_of` references naming standard control types. Both are real but neither
is load-bearing, and the table would have to be written and maintained while
covering only VB's own controls — third-party OCX types, which are abundant in
real codebases, would still be external.

**Verdict for now: not worth it.** Revisit only if querying control members as
symbols turns out to matter in practice.

## 3. Multi-project resolution across a `.vbg` — **answered and implemented**

Real groups do contain same-named public procedures in sibling projects, and
abundantly: one name had 21 definitions across 111 projects in the measured
codebase. Without project scope the resolver saw ambiguity everywhere and
resolved nothing.

Resolution now filters candidates by `.vbp` membership, which lifted the
resolution rate from 28.5% to 35.7%. What remains open is the **20% of files
that belong to no indexed `.vbp`**: see the limitation of the same name, and
the measured reason a directory fallback was rejected.

## 4. `Property Let` versus `Property Get` on a write

A reference that could be either currently resolves to the getter. Telling
them apart needs the assignment context, which the extractor knows and does
not currently record.

**Cost of fixing:** one more distinction carried through to resolution.
**Cost of not fixing:** a write shows as a read in callers/impact.

Worth doing if property-heavy code proves common.

## 5. Are procedure locals worth their node count? — **measured, still open**

On the real codebase, `variable` nodes are 28,644 of 141,016 — **20% of the
graph** — and they exist mostly so that an array access can be told from a
call. That purpose is real but it does not need them to be *nodes*; a per-file
set of declared names would do.

Still open because the cost is tolerable and removing them would lose a
`Variable` the semantic model claims to have. Revisit if graph size becomes a
constraint.

## 6. Does the event-binding heuristic hold up at scale? — **partly answered**

On the real codebase it produced 12,229 bindings and left 38% of
underscore-named methods unbound — the precision gate refusing to guess where
no producer of that name exists in the file. That is the intended shape.

Still open: the **false-negative** side. Some of that 38% are genuine handlers
whose producer the designer section did not yield (a control added at run time
with `Load`, or a `WithEvents` field declared in a different file). Nothing
measured yet says how large that share is; a sample would need reading the
proprietary source, which is out of scope here.

## 7. Should VB6 get a `codegraph_explore` flow test?

Every language/framework in the repo is expected to be validated with the
agent A/B methodology (small/medium/large repos, ≥3 flow prompts). VB6 has
none: the canonical flow would be *user action → event handler → business
logic → data access*. This needs a real VB6 repo of each size, which is the
same blocker as (1).
