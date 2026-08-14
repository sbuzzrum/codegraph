# VB6 Support — Piano di lavoro e stato

> Documento vivo. Fonte autoritativa dei requisiti: `PROMPT_CODEGRAPH_VB6_FORK_PERSONALE.md`
> (nella dir padre del repo). Questo file traccia **piano**, **decisioni** e **stato**.

Ultimo aggiornamento: 2026-08-14 — fase corrente: **2–3 in corso** (lingua `vb6`
registrata e committata; extractor custom implementato per code/designer/progetto;
**mancano** suite di conformità/test e l'intero resolver — Fase 4).

Committato e pushato su `origin/main`:
- `2d59953` — `docs(vb6): architecture audit, phased plan and implementation decisions`
- `530b584` — `feat(vb6): register language and add custom extractors`

---

## 0. Stato oggettivo del repo (audit iniziale)

- Fork configurato: `origin` → `github.com/sbuzzrum/codegraph`, `upstream` → `colbymchenry/codegraph`
  (push upstream disabilitato). Branch `main`, working tree pulito, allineato a `origin/main`.
- Tooling: Node 24 (engines repo: `>=20 <25`), `gh` autenticato come `sbuzzrum`, build `tsc` + copy-assets,
  test `vitest`.
- **Nessun lavoro VB6 preesistente** (0 file, nessuna `docs/vb6/` prima di questo documento).

### Architettura CodeGraph rilevata (punti di aggancio reali)

| Concetto | Dove | Note per VB6 |
|---|---|---|
| Elenco lingue | `src/types.ts` → const `LANGUAGES` (→ `type Language`) | aggiungere `'vb6'` |
| NodeKind / EdgeKind | `src/types.ts` (**set FISSI**) | **mappare** l'ontologia VB6, non estenderla se evitabile |
| Grammar WASM ↔ lingua | `src/extraction/grammars.ts` (mappa `GRAMMAR_WASM`) | `tree-sitter-vbnet.wasm` già presente |
| Estensione ↔ lingua | `src/extraction/grammars.ts` (mappa ext) | `.vb`→vbnet; aggiungere `.bas/.cls/.frm/.ctl` |
| Config per-lingua (dichiarativa) | `src/extraction/languages/<lang>.ts` + `languages/index.ts` (barrel `EXTRACTORS`) | `vbnet.ts` è il template più vicino |
| Interfaccia config | `src/extraction/tree-sitter-types.ts` (`LanguageExtractor`) | definisce classTypes/methodTypes/callTypes/visitNode… |
| Extractor formati compositi | `src/extraction/*-extractor.ts` (`dfm-extractor.ts` per form Delphi, `vue`, `svelte`, `razor`) | **template per `.frm`/`.ctl`** (codice + designer) |
| Provenance edge | edge sintetizzati: `provenance:'heuristic'` + `metadata.synthesizedBy` | riusare per binding eventi/WithEvents |
| Risoluzione | `src/resolution/` (`ReferenceResolver`, `name-matcher`, `frameworks/`) | scope/qualified calls VB6 |
| Sintesi dispatch dinamico | `src/resolution/callback-synthesizer.ts`, `frameworks/` | modello per event binding VB6 |
| MCP | `src/mcp/server-instructions.ts` + `tools.ts` | esporre lingue supportate |

### NodeKind disponibili (fissi)
`file, module, class, struct, interface, trait, protocol, function, method, property, field,
variable, constant, enum, enum_member, type_alias, namespace, parameter, import, export, route,
component, union`

### EdgeKind disponibili (fissi)
`contains, calls, imports, exports, extends, implements, references, type_of, returns,
instantiates, overrides, decorates`

---

## 1. Decisione architetturale cardine — mapping ontologia VB6

Il prompt (§8) elenca nodi/relazioni VB6 ricchi. NodeKind/EdgeKind sono fissi. **Strategia: mappare + provenance/metadata**, evitando fork del core. Nuovi EdgeKind solo se un concetto è irriducibile (da decidere in fase 4, documentato in `IMPLEMENTATION_DECISIONS.md`).

**Nodi VB6 → NodeKind**
| VB6 | NodeKind | Distinzione |
|---|---|---|
| Project / ProjectGroup | `module` (o `namespace`) | `metadata.vb6Kind='project'/'projectGroup'` |
| Module (.bas) | `module` | |
| Class (.cls) | `class` | |
| Form (.frm) | `class` | `metadata.vb6Kind='form'` |
| UserControl (.ctl) | `class` | `metadata.vb6Kind='userControl'` |
| Control (istanza nel designer) | `field` | `metadata.vb6Kind='control'`, `controlType`, `index?` |
| Sub / Function | `method`/`function` | container-dipendente |
| PropertyGet/Let/Set | `property` | `metadata.accessor` |
| Event (`Event X`) | `field` (come vbnet) | `metadata.vb6Kind='event'` |
| Variable/Constant/Enum/EnumMember | omonimi | |
| UserDefinedType (`Type`) | `struct` | |
| External/COM/OCX reference | `import` | `metadata.vb6Kind='comReference'/'ocxReference'/'declare'` |

**Relazioni VB6 → EdgeKind (+ provenance)**
| VB6 | EdgeKind | provenance |
|---|---|---|
| CONTAINS / DECLARES | `contains` | static |
| CALLS | `calls` | static / lexical/module/project scope |
| READS / WRITES | `references` | static |
| INSTANTIATES (`New`, `As New`) | `instantiates` | static/type_resolution |
| RETURNS | `returns` | static |
| IMPLEMENTS | `implements` | static |
| USES_CONTROL / INSTANCE_OF | `references` / `type_of` | static |
| USES_OCX / REFERENCES_COM | `references` | typelib / heuristic |
| DECLARES_EVENT | `contains` | static |
| RAISES_EVENT | `references` (o `calls`) | static |
| HANDLES_EVENT / HANDLES_OCX_EVENT | `references` | event_binding / withevents_binding / synthesized / heuristic |
| PROJECT_CONTAINS | `contains` | static (da .vbp) |
| PROJECT_REFERENCES | `imports` | static (da .vbp/.vbg) |

**Provenance VB6** (`metadata.provenance` / `synthesizedBy`), mappata sui concetti del prompt §14:
`static, lexical_scope, module_scope, project_scope, type_resolution, event_binding,
withevents_binding, typelib, synthesized, heuristic, unresolved`.
Regola d'oro (§21): in caso di ambiguità **preferire `unresolved`** a un edge falso.

---

## 2. Decisione grammar VB6

- **Baseline:** riuso `tree-sitter-vbnet.wasm` già presente (VB6 e VB.NET condividono gran parte della sintassi di dichiarazioni/procedure/chiamate).
- **Da verificare in Fase 2** su fixture reali VB6 la copertura di: header `.frm`/`.ctl` (`VERSION`, `Object=`, blocco `Begin … End`), `Property Get/Let/Set`, `Event`, `RaiseEvent`, `WithEvents`, `Implements`, control array, `#If/#Const`, line continuation `_`.
- **Sezione designer dei `.frm/.ctl` NON è codice VB** → gestita da **extractor composito dedicato** (modello `dfm-extractor.ts`): si separa l'header designer (parser custom line-based) dal corpo codice (grammar tree-sitter). Deciso a priori: la designer section non passa per tree-sitter-vbnet.
- Se la grammar vbnet fallisce su costrutti VB6 core del corpo codice: valutare grammatiche tree-sitter VBA/VB6 esterne o estendere quella vendorizzata; ProLeap ANTLR come **oracle** per casi dubbi (non come runtime). Ogni scelta documentata in `VB6_FORK_ASSESSMENT.md` + licenze.

---

## 3. Piano a fasi

### Fase 1 — Audit & fondamenta *(quasi completa)*
- [x] Audit architettura extractor/registry/resolver/provenance/MCP.
- [x] `docs/vb6/IMPLEMENTATION_DECISIONS.md` (mapping ontologia, scelta grammar D1, scope model). Committato.
- [ ] `docs/vb6/VB6_SEMANTIC_MODEL.md` (nodi/relazioni/provenance definitivi). **Ancora da scrivere.**
- **Commit:** `docs(vb6): architecture audit and semantic model` → parzialmente fatto (`2d59953`, manca SEMANTIC_MODEL).

### Fase 2 — Grammar & fixture di conformità *(in corso)*
- [x] Registrare lingua `vb6` (types.ts + grammars.ts: estensioni `.bas/.cls/.frm/.ctl/.vbp/.vbg`, display
      name, liste supported/grammar-loaded; dispatch in `tree-sitter.ts`). **NB:** nessuna entry wasm (D1).
      Committato in `530b584`.
- [ ] Suite conformità §7: 50 mini-progetti in `__tests__/fixtures/vb6/NN_*/` con oracolo machine-readable
      (`expected.json`: nodi+edge attesi). Test runner che confronta graph prodotto vs atteso. **0 fixture, 0 test oggi.**
- [ ] ~~Test grammar~~ → **N/A** (nessuna grammar tree-sitter, D1); sostituito da: test unitari dell'extractor sui costrutti core.
- **Commit:** `feat(vb6): register language and grammar` fatto (`530b584`); `test(vb6): add conformance fixtures and oracle runner` **da fare.**

### Fase 3 — Extractor dichiarazioni & progetto *(implementato, non ancora testato)*
> Realizzato come extractor custom line-based in `src/extraction/vb6-extractor.ts` (non `languages/vb6.ts`), per D1.
- [x] `Vb6Extractor`: Sub/Function/Property (Get/Let/Set)/Event/Enum/Type(UDT)/Const/Variable/Declare,
      visibilità Public/Private/Friend/Static, return type; chiamate (`calls`), `references`, `RaiseEvent`,
      member/implicit call e `As New` (`instantiates`) con hint `metadata.vb6` (raises_event/member_call/…).
- [x] `Vb6ProjectExtractor` — `.vbp`/`.vbg`: nodo progetto + membership/reference.
- [x] `Vb6FormExtractor` — `.frm`/`.ctl`: Form/UserControl + controlli (tipo, nome istanza, nesting,
      control array `Index`, OCX ref) dalla designer section `Begin…End`; corpo codice delegato a `Vb6Extractor`.
- [ ] **Validazione:** nessun test ancora esercita questi extractor; correttezza da confermare in Fase 2/5.
- **Commit:** confluiti in `530b584` (`feat(vb6): register language and add custom extractors`).

### Fase 4 — Resolver, scope, eventi, COM
- [ ] Scope resolution (§9): procedure/module/class/form/project scope, Public/Private/Friend/Static,
      qualified vs unqualified, moduli globali, omonimi in scope diversi.
- [ ] Costrutti (§16): `With`, `As New`/`New`, `Implements`, `Call`/chiamata implicita, named args, Optional,
      ByRef/ByVal, ParamArray, line continuation, error handler; default members/default form instance.
- [ ] Event model (§13): handler `Control_Event`, eventi Form, `Event`/`RaiseEvent`/`WithEvents`,
      binding consumer↔producer con provenance corretta (static/event_binding/withevents_binding/synthesized/heuristic/unresolved).
- [ ] COM/OCX (§15): CLSID/ProgID/typelib quando disponibili; early vs late binding; late binding **non**
      inventa target.
- [ ] Conditional compilation (§17): riconosci `#If/#ElseIf/#Else/#End If/#Const`; non eliminare rami silenziosamente.
- **Commit:** `feat(vb6): resolve scope and qualified calls`, `feat(vb6): resolve events/WithEvents/RaiseEvent`, `feat(vb6): model COM/OCX references`

### Fase 5 — Integrazione, adversarial, MCP, regressioni
- [ ] Casi adversarial §21 (omonimi, private/public, control array, late binding, default member…):
      il resolver preferisce `unresolved` al falso positivo.
- [ ] MCP: `server-instructions.ts` espone VB6 tra le lingue supportate; verifica search/callers/callees/impact/explore su graph VB6.
- [ ] Metriche §20 riproducibili (fixture pass, simboli/edge attesi vs trovati, FP/FN, unresolved, parse errors, regressioni).
- [ ] **Regressioni:** `npm test` verde su tutte le lingue esistenti (nessuna regressione).
- **Commit:** `test(vb6): adversarial resolution cases`, `feat(mcp): surface vb6`, `docs(vb6): conformance metrics`

### Fase 6 — Documentazione, assessment, push
- [ ] Docs §22: `VB6_SUPPORT.md`, `VB6_ARCHITECTURE.md`, `VB6_LIMITATIONS.md`, `VB6_CONFORMANCE.md`,
      `TEST_RESULTS.md`, `OPEN_QUESTIONS.md` (+ SEMANTIC_MODEL/IMPLEMENTATION_DECISIONS già iniziati).
- [ ] `VB6_FORK_ASSESSMENT.md` con giudizio: READY / READY WITH RESERVATIONS / NOT READY.
- [ ] Push su `origin` (mai upstream, mai PR upstream). CHANGELOG `[Unreleased]` se opportuno.
- **Commit:** `docs(vb6): support, limitations and fork assessment`

---

## 4. Vincoli permanenti (dal prompt)
- Codice **general-purpose**, zero riferimenti Itineris (nomi/path/workaround). §23.
- Lavoro solo su branch `main` del fork, commit piccoli e logici, push solo su `origin`. §4–5.
- Mai riscrivere storia upstream, mai PR upstream, mai credenziali/codice proprietario nel fork. §28.
- Nessun supporto VB6 è "completo" se rompe lingue esistenti. §19.
- Repo reali VB6 solo dopo che conformità + regressioni passano. §24.

## 5. Log decisioni & avanzamento
- 2026-08-13: audit architettura completato; piano scritto; mapping ontologia proposto. In attesa di via libera per Fase 1 (doc semantico) → Fase 2.
- 2026-08-14: lingua `vb6` registrata (types.ts/grammars.ts/tree-sitter.ts) ed extractor custom
  implementato (`vb6-extractor.ts`: code + designer + progetto). Typecheck `tsc --noEmit` verde.
  Committato in due unità logiche (`2d59953` docs, `530b584` feat) e **pushato su `origin/main`**.
  **Prossimi passi:** (1) `VB6_SEMANTIC_MODEL.md`; (2) suite di conformità + test runner (Fase 2);
  (3) resolver scope/eventi/COM (Fase 4). Nulla è ancora testato: gli extractor sono da validare.
