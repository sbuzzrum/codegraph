# VB6 — Decisioni implementative

Registro cronologico delle decisioni tecniche vincolanti. Ogni decisione cita l'evidenza.

---

## D1 — Grammar: **NON** riusare la grammar tree-sitter VB.NET. Extractor custom line-based.

**Data:** 2026-08-13

**Evidenza (misurata):** parse di 5 snippet VB6 rappresentativi con `tree-sitter-vbnet.wasm`
(script `scratchpad/vb6-grammar-check*.mjs`). Anche dopo un `preParse` che neutralizza le righe
metadata VB6 (`Attribute`, `VERSION/BEGIN..END`, `#If`), la grammar produce errori su costrutti
**core** del linguaggio, non solo sul rumore:

| Costrutto VB6 | Esito grammar vbnet |
|---|---|
| `Property Get/Let/Set name() As T` | ERROR (VB.NET usa `Property name As T` con blocchi Get/Set) |
| `Public Event ItemRead(...)` | ERROR |
| `Private WithEvents mTimer As Timer` | ERROR |
| `Implements IFoo` (module level) | ERROR |
| `Dim o As New CItem` | ERROR |
| `Optional ByRef y As Long = 0` (param) | ERROR |
| `Public Const MAX As Long = 10` (module) | ERROR |

Conclusione: VB6 e VB.NET divergono troppo (property accessors, UDT `Type…End Type` vs `Structure`,
`Declare`, header designer). La grammar vbnet **non è adeguata** (prompt §6).

**Toolchain:** nessun `tree-sitter` CLI, nessun `emcc`/emscripten, nessun `docker` sulla macchina
(solo gcc/make). Non è possibile compilare localmente una grammar tree-sitter VBA→wasm in modo
affidabile e riproducibile.

**Decisione:** implementare un **extractor VB6 custom, line/token-based**, senza tree-sitter — lo
stesso pattern già adottato in CodeGraph per formati senza grammar adeguata: `dfm-extractor.ts`,
`mybatis-extractor.ts`, `liquid-extractor.ts`, parti markup di `razor-extractor.ts`. VB6 è
statement-oriented, keyword-led, con terminatori di blocco espliciti (`End Sub/Function/Property/Type/Enum`)
e continuazione riga `_`: un parser a righe raggiunge alta fedeltà, è deterministico e testabile
contro le 50 fixture di conformità.

**Conseguenze architetturali:**
- La lingua `vb6` è registrata in `LANGUAGES` (types.ts) e nella mappa estensioni di `grammars.ts`,
  ma **NON** ha una entry wasm né una entry nella mappa tree-sitter `EXTRACTORS` (languages/index.ts).
- Il dispatch avviene in `extractFromSource` (tree-sitter.ts) con un ramo dedicato, come svelte/vue/dfm.
- Tre extractor:
  - `Vb6Extractor` (`vb6-extractor.ts`) — corpo codice `.bas/.cls/.frm/.ctl` (dichiarazioni, chiamate, eventi…).
  - `Vb6FormExtractor` — sezione designer `.frm/.ctl` (controlli, control array, OCX ref); delega il corpo
    codice a `Vb6Extractor` (modello composito dfm/vue). *(può essere lo stesso file/classe)*
  - `Vb6ProjectExtractor` — `.vbp/.vbg` (membership progetto, reference).
- ProLeap ANTLR resta come **oracle** concettuale per casi dubbi, non a runtime.

**Rischio/limite:** un parser custom non ha la robustezza AST di tree-sitter su codice molto irregolare;
mitigato da: continuazione riga, gestione stringhe/commenti, e suite di conformità + casi adversarial.
Documentato in `VB6_LIMITATIONS.md`.

---

## D2 — Mapping ontologia VB6 su NodeKind/EdgeKind fissi (no fork del core)

**Data:** 2026-08-13

NodeKind/EdgeKind sono set fissi (`src/types.ts`). L'ontologia VB6 del prompt §8 è **mappata** su
di essi + `metadata.vb6Kind` e `provenance`, senza introdurre nuovi kind. Tabella completa in
`VB6_SEMANTIC_MODEL.md`. Nuovi kind solo se un concetto risultasse irriducibile (nessuno finora).

Provenance edge (prompt §14) registrata in `metadata.provenance`:
`static | lexical_scope | module_scope | project_scope | type_resolution | event_binding |
withevents_binding | typelib | synthesized | heuristic | unresolved`.
Regola (§21): in ambiguità, **preferire `unresolved`** a un edge falso.

---

## D3 — Designer `.frm/.ctl` come extractor composito

**Data:** 2026-08-13

La sezione designer (`VERSION … / Begin … End`) non è codice VB e viene parsata da un parser
line-based dedicato (modello `dfm-extractor.ts`): estrae Form/UserControl, controlli (tipo, nome
istanza, nesting, control array `Index`), riferimenti OCX. Il resto del file (dopo l'header) è codice
e va a `Vb6Extractor`. Gli handler eventi `Control_Event` sono collegati in fase di resolution.
