# VB6 Support — Piano di lavoro e stato

> Documento vivo. Fonte autoritativa dei requisiti: `PROMPT_CODEGRAPH_VB6_FORK_PERSONALE.md`
> (nella dir padre del repo). Questo file traccia **piano**, **decisioni** e **stato**.

Ultimo aggiornamento: 2026-08-14 — **tutte le fasi completate, §24 inclusa**.
Motore VB6 completo (extractor + resolver + binding eventi), MCP verificato,
documentazione e assessment prodotti, **validazione real-world eseguita** su una
codebase VB6 di produzione (2.163 file, 0 file non parsati).
Suite di conformità **69/69**, 0 gap, 0 falsi positivi, 0 regressioni.
Su codebase reale: tasso di risoluzione **72,3%**, copertura del call graph interno **86,6%**
(**96,8%** escludendo le chiamate che VB6 stesso non può raggiungere senza qualificatore).
Giudizio (§25): **READY FOR USE**.

Committato e pushato su `origin/main`:
- `2d59953` — `docs(vb6): architecture audit, phased plan and implementation decisions`
- `530b584` — `feat(vb6): register language and add custom extractors`

---

## 0. Stato oggettivo del repo (audit iniziale)

- Fork configurato: `origin` → `github.com/sbuzzrum/codegraph`, `upstream` → `colbymchenry/codegraph`
  (push upstream disabilitato con `git remote set-url --push upstream DISABLED_no_push_to_upstream`).
  Branch `main`, working tree pulito, allineato a `origin/main`.
  **Nota:** `upstream` risultava dichiarato qui ma non esisteva davvero nel repo; riconfigurato il 2026-08-14.
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

**Provenance VB6** — *questa proposta iniziale è stata superata in fase 4.* La colonna
`edges.provenance` ha dominio chiuso e la convenzione del repo è che solo i synthesizer la marchino:
gli edge statici VB6 la lasciano vuota e portano il dettaglio in `metadata` (`scope`, `qualifier`,
`vb6`), i binding eventi sono `heuristic`. Vedi D4/D5 in `IMPLEMENTATION_DECISIONS.md` e la tabella
autoritativa in `VB6_SEMANTIC_MODEL.md`.
Regola d'oro (§21): in caso di ambiguità **preferire `unresolved`** a un edge falso.

---

## 2. Decisione grammar VB6 — *esito*

L'ipotesi iniziale (riusare `tree-sitter-vbnet.wasm`) è stata **verificata e scartata**: la grammar
VB.NET va in errore su costrutti **core** di VB6 (`Property Get/Let/Set`, `Event`, `WithEvents`,
`Implements`, `Type…End Type`, `As New`, `Declare`), e la macchina non ha una toolchain per compilare
una grammar VBA→wasm. Decisione finale: **extractor custom line-based**, senza tree-sitter — vedi D1
in `IMPLEMENTATION_DECISIONS.md` per l'evidenza misurata e `VB6_ARCHITECTURE.md` per il risultato.
ProLeap ANTLR resta oracle concettuale, mai a runtime. Nessuna grammar di terzi vendorizzata, quindi
nessuna implicazione di licenza.

---

## 3. Piano a fasi

### Fase 1 — Audit & fondamenta *(completa)*
- [x] Audit architettura extractor/registry/resolver/provenance/MCP.
- [x] `docs/vb6/IMPLEMENTATION_DECISIONS.md` (mapping ontologia, scelta grammar D1, scope model). Committato.
- [x] `docs/vb6/VB6_SEMANTIC_MODEL.md` — nodi, relazioni, marker `vb6:*`, modello eventi e provenance.
      Ogni regola è pinnata da una fixture, quindi il documento non può divergere dal codice in silenzio.
- **Commit:** `2d59953` + `test(vb6): add conformance suite`.

### Fase 2 — Grammar & fixture di conformità *(completa)*
- [x] Registrare lingua `vb6` (types.ts + grammars.ts: estensioni `.bas/.cls/.frm/.ctl/.vbp/.vbg`, display
      name, liste supported/grammar-loaded; dispatch in `tree-sitter.ts`). **NB:** nessuna entry wasm (D1).
      Committato in `530b584`.
- [x] Suite conformità §7: **56 mini-progetti** in `__tests__/fixtures/vb6/NN_*/` con oracolo
      machine-readable (`expected.json`) e runner `__tests__/vb6-conformance.test.ts` che indicizza
      ogni fixture con la pipeline reale (estrazione → resolution → SQLite) e confronta il graph
      prodotto con quello atteso. Copre i 50 casi del prompt più 6 casi aggiuntivi (adversarial §21 e range del corpo dei simboli).
- [x] ~~Test grammar~~ → **N/A** (nessuna grammar tree-sitter, D1); sostituito dal runner di conformità,
      che esercita l'extractor end-to-end.
- **Commit:** `feat(vb6): register language and grammar` (`530b584`); `test(vb6): add conformance suite`.

#### Come funziona l'oracolo
`expected.json` descrive la **semantica VB6 corretta**, non il comportamento odierno. Ogni fixture ha
`"status"`:
- `pass` → il graph deve già combaciare; se non combacia è una **regressione** e il test fallisce;
- `known-gap` + `"gap"` (motivo) → la semantica non è ancora implementata: il runner verifica che il gap
  sia **ancora presente** e fallisce con "promote to pass" appena viene chiuso.

Il match è per sottoinsieme: `expect` deve trovare almeno un riscontro (o esattamente `count`), `forbid`
nessuno. `forbid` è ciò che rende la suite utile contro i falsi positivi (§21).

`VB6_CONFORMANCE_REPORT=1 npx vitest run __tests__/vb6-conformance.test.ts` rigenera
`docs/vb6/VB6_CONFORMANCE.md` con le metriche §20.

### Fase 3 — Extractor dichiarazioni & progetto *(implementato e validato)*
> Realizzato come extractor custom line-based in `src/extraction/vb6-extractor.ts` (non `languages/vb6.ts`), per D1.
- [x] `Vb6Extractor`: Sub/Function/Property (Get/Let/Set)/Event/Enum/Type(UDT)/Const/Variable/Declare,
      visibilità Public/Private/Friend/Static, return type; chiamate (`calls`), `references`, `RaiseEvent`,
      member/implicit call e `As New` (`instantiates`) con hint `metadata.vb6` (raises_event/member_call/…).
- [x] `Vb6ProjectExtractor` — `.vbp`/`.vbg`: nodo progetto + membership/reference.
- [x] `Vb6FormExtractor` — `.frm`/`.ctl`: Form/UserControl + controlli (tipo, nome istanza, nesting,
      control array `Index`, OCX ref) dalla designer section `Begin…End`; corpo codice delegato a `Vb6Extractor`.
- [x] **Validazione (2026-08-14):** i tre extractor sono ora esercitati dalle 55 fixture. Risultato:
      **29 fixture passano, 26 sono gap documentati, 0 parse error, 0 regressioni** sulle altre lingue.
      Funzionano dichiarazioni, visibilità, designer `.frm/.ctl` con nesting e control array, `.vbp`/`.vbg`,
      `New`/`As New`, `Implements`, line continuation, commenti.
- **Commit:** confluiti in `530b584` (`feat(vb6): register language and add custom extractors`).

#### Difetti dell'extractor — **tutti corretti** (2026-08-14)
Rilevati dalla suite, corretti, e ora protetti dalla fixture che li ha scoperti (promossa a `pass`).

| # | Difetto | Correzione | Fixture |
|---|---|---|---|
| E1 | chiamata qualificata senza parentesi (`ModA.Process`, `Form1.Reload`, `obj.Metodo arg`) risolta sul **qualificatore** | scanner per catene puntate: il target è il membro, il qualificatore va in `metadata.qualifier` | 26, 29, 30 |
| E2 | `With … End With` ignorato | stack dei `With`; `.Membro` è attribuito al target del blocco | 24 |
| E3 | argomenti nominati (`Nome:=valore`) impediscono il rilevamento della chiamata | l'assegnazione è riconosciuta da `=`, non da `:=` | 33 |
| E4 | le **label** di riga (`ErrHandler:`) lette come chiamate implicite | le label sono riconosciute e ignorate | 38 |
| E5 | **stringhe letterali** non neutralizzate prima della ricerca di chiamate | `maskStrings` azzera il contenuto conservando gli offset | 54 |
| E6 | indicizzazione di array indistinguibile da una chiamata | scan differito: a fine file si sa cosa è stato dichiarato come dato → `references` invece di `calls` | 55 |
| E7 | tipi intrinseci emessi come riferimenti mai risolvibili | i tipi intrinseci non generano riferimenti | 05 |
| E8 | dichiarazioni locali non estratte come nodi | `Dim`/`Static`/`Const` in procedura creano nodi `variable`/`constant` | 13 |
| E9 | `Declare` non cattura il tipo di ritorno | regex estesa a `(…) As Tipo` | 49 |
| E10 | control array: un nodo per elemento | i `Begin` omonimi nello stesso contenitore condividono un nodo, gli indici si accumulano | 18 |
| E11 | nessun nodo per COM/OCX reference | nodi `import` con `vb6:com-reference`/`vb6:ocx-reference`, CLSID e typelib nel `docstring`, edge dal progetto e dal form | 40, 47 |
| E12 | ProgID di `CreateObject` scartato | il ProgID è estratto prima del mascheramento stringhe e registrato con `vb6: late_binding` | 41 |

Effetto misurato sul progetto di prova: le chiamate qualificate ora puntano al membro giusto
(`cmdAction_Click → DoWork`, prima `→ Module1`), il `With` produce riferimenti, e i riferimenti
irrisolvibili scendono da 44 a 9 — tutte funzioni/tipi intrinseci VB6 o simboli realmente assenti.

Effetto sulla suite: **29 → 44 fixture pass**, gap 26 → 11, falsi positivi 10 → 4, falsi negativi 30 → 13.

### Fase 4 — Resolver, scope, eventi, COM *(completa)*

#### Difetti della resolution — **tutti corretti** (2026-08-14)

| # | Difetto | Correzione | Fixture |
|---|---|---|---|
| R1 | nessuno scope: si risolveva per uguaglianza di nome | `Vb6Resolver` applica le regole VB6; in caso di ambiguità **nessun edge** | 02, 31, 51 |
| R2 | membri di classe raggiungibili non qualificati dall'esterno | lo scope di progetto ammette solo i `Public` dei moduli standard (`.bas`) | 31 |
| R3 | chiamate late-bound legate a un membro omonimo qualsiasi | un qualificatore di tipo `Object`/`Variant` non ha target statico → unresolved | 43 |
| R4 | nessun binding di eventi | `vb6-event-synthesizer`: produttore → handler per controlli, Form, `WithEvents`, OCX | 16, 17, 22, 23, 39, 52 |
| R5 | `provenance` assente sugli edge VB6 | binding eventi marcati `heuristic`; gli edge risolti portano `metadata.scope` e il qualificatore | 21 |
| R6 | il metadata dell'extractor non sopravvive al giro in DB | il qualificatore viaggia in `candidates` (già persistito); il resto si deduce dal grafo | 21 |

**Nessuna migrazione di schema.** `unresolved_refs.candidates` significa già "nomi qualificati a cui il
riferimento potrebbe risolversi": l'extractor ci scrive `["c.Compute"]` e il resolver ne ricava il
qualificatore. Che un riferimento sia un `RaiseEvent` si deduce dal target (nodo `vb6:event`); che una
chiamata sia late-bound, dal `returnType` della variabile (`Object`/`Variant`).

**Bug del core trovato strada facendo:** `candidates` veniva scritto e riletto dal DB, ma **tre**
conversioni `UnresolvedReference → UnresolvedRef` in `src/resolution/index.ts` lo lasciavano cadere
prima di `resolveOne`. Nessun'altra lingua lo legge, quindi il campo era di fatto morto; ora è
propagato — modifica additiva, nessun effetto sulle altre lingue.

Il flusso end-to-end richiesto dal §13 è ora percorribile nel graph:

```
ReadBarcode --references(raises_event)--> ItemRead --calls(withevents)--> mReader_ItemRead
```

Componenti aggiunti:
- `src/resolution/vb6-resolver.ts` — scope VB6, chiamate qualificate, late binding, ambiguità → unresolved.
  Agganciato in `resolveOne` **senza fallthrough**: se le regole VB6 non determinano un target, la ref
  resta unresolved invece di finire al name-matcher generico.
- `src/resolution/vb6-event-synthesizer.ts` — binding eventi per convenzione di nome, con `provenance:
  'heuristic'` e `metadata.binding`. Un handler si lega solo se il produttore esiste davvero nello stesso
  file: è ciò che impedisce a `IWorker_Run` (metodo di `Implements`) di essere scambiato per un handler.

Lavoro residuo di fase (non bloccante, nessuna fixture lo copre): `Friend` è trattato come `Public`
(corretto entro il singolo progetto), i rami `#If` restano tutti nel graph senza valutazione, e le
type library COM esterne non sono lette — i tipi dei controlli VB standard restano quindi unresolved.

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

### Fase 5 — Integrazione, adversarial, MCP, regressioni *(completa)*
- [x] Casi adversarial §21: fixture 43, 50, 51, 52, 54, 55 — il resolver preferisce `unresolved`
      al falso positivo, e i `forbid` degli oracoli lo verificano.
- [x] MCP: `query`/`callers`/`callees`/`impact`/`node`/`explore` verificati sul graph VB6.
      `explore` mostra i binding eventi come hop dinamici con il wiring site
      (`VB6 WithEvents event `ItemRead` @Form1.frm:28`); `node` restituisce il corpo completo.
      VB6 nominato in `server-instructions.ts`.
      **Bug trovato e corretto qui:** i nodi procedura avevano `endLine == startLine`, quindi
      `codegraph_node` mostrava la sola firma e costringeva l'agente a rileggere il file — esattamente
      ciò che CodeGraph esiste per evitare. Fixture `56_symbol_body_range` a protezione.
- [x] Metriche §20 riproducibili: `VB6_CONFORMANCE.md` generato dal runner; `TEST_RESULTS.md`.
- [x] Regressioni: `npm test` verde (2958 test), nessuna lingua esistente toccata.

### Fase 6 — Documentazione, assessment, push *(completa)*
- [x] `VB6_SUPPORT.md`, `VB6_ARCHITECTURE.md`, `VB6_LIMITATIONS.md`, `VB6_CONFORMANCE.md`,
      `TEST_RESULTS.md`, `OPEN_QUESTIONS.md`, `VB6_SEMANTIC_MODEL.md`, `IMPLEMENTATION_DECISIONS.md`.
- [x] `VB6_FORK_ASSESSMENT.md` — **READY FOR USE WITH RESERVATIONS**.
- [x] Push su `origin/main` (mai upstream, mai PR upstream).

### Fase 7 — Validazione real-world *(completa, §24)*

Eseguita su una codebase VB6 di produzione multi-progetto, **copiata in sola lettura** fuori dal
repository (§28: sorgenti originali mai toccati; §23: nessun nome, path o snippet proprietario nel fork).

**Scala:** 2.163 file (111 `.vbp`, 49 `.vbg`, 903 `.frm`, 647 `.cls`, 299 `.bas`, 154 `.ctl`),
141.016 nodi, 321.442 edge, 9,5 s di indicizzazione, 1,8 GB RSS di picco, **0 file senza simboli**.

**Tre difetti generali trovati e corretti**, ognuno con la sua fixture *generica* scritta prima del fix:

| Difetto | Fixture |
|---|---|
| scope di progetto assente: un nome VB6 è scoped al suo `.vbp`, non all'intero indice — con 111 progetti insieme un nome aveva 21 definizioni e nessuna risolveva | `57_project_scope_isolation` |
| visibilità implicita: procedura senza modificatore = `Public`, `Dim`/`Const` a livello modulo = `Private` (3.280 metodi senza visibilità) | `58_implicit_visibility` |
| accesso ad array dichiarato in un altro modulo registrato come chiamata | `59_public_array_cross_module` |

La membership di progetto doveva essere leggibile **senza dipendere dall'ordine di resolution** (gli edge
`contains` che la descrivono sono prodotti dalla stessa pass che ne ha bisogno): l'extractor del `.vbp`
registra ora i file membri sul nodo progetto.

**Effetto round 1:** edge non-contains 141.648 → 177.036; tasso di risoluzione 28,5% → 35,7%.

**Round 2** — dalla misura del residuo (il 61,9% erano accessi a membri di tipi non indicizzati) sono
emersi altri due difetti, entrambi con fixture:

| Difetto | Fixture |
|---|---|
| l'accesso a un membro di un tipo esterno (`txtName.Text`) non produceva **alcun** edge: si perdeva anche il legame con l'oggetto, che invece è nel graph | `60_member_on_external_type` |
| il designer scrive il tipo come `Libreria.Controllo` e l'extractor scartava la libreria: con due UserControl omonimi in progetti diversi il tipo non si risolveva (collegamento client→OCX del §12) | `61_ocx_type_from_sibling_project` |

**Effetto round 2:** edge non-contains → **269.496**; tasso di risoluzione → **55,1%**.
91.268 riferimenti a membri ora legati all'oggetto su cui sono usati (con il nome del membro sull'edge),
1.198 controlli collegati al proprio UserControl nel progetto ActiveX fratello.

Un ripiego per prossimità di directory (per i file fuori da ogni `.vbp`) recupererebbe l'1,7% al prezzo
di un'euristica che può sbagliare: **scartato** (§21).

**Indicizzare per singolo `.vbp` invece che l'albero intero: misurato e scartato.** L'ipotesi era che
isolare un progetto eliminasse le ambiguità create dalle copie duplicate di UserControl condivisi.
Tre progetti (711, 45 e 13 file) indicizzati in entrambi i modi: tasso identico a meno di 0,1 punti —
il filtro di scope già scarta i candidati fuori progetto — e l'isolamento **perde** i tipi forniti dai
progetti ActiveX fratelli (su un progetto piccolo, `type_of` risolti da 31 a 5). Raccomandazione
corretta: **indicizzare l'albero intero**.

**Eventi su scala reale:** 12.229 binding sintetizzati; il 38% dei metodi con nome `X_Y` resta non
legato — è il gate di precisione che rifiuta di indovinare dove il produttore non esiste nel file.

## 4. Vincoli permanenti (dal prompt)
- Codice **general-purpose**: zero nomi, path o workaround del progetto proprietario usato per la validazione. §23.
- Lavoro solo su branch `main` del fork, commit piccoli e logici, push solo su `origin`. §4–5.
- Mai riscrivere storia upstream, mai PR upstream, mai credenziali/codice proprietario nel fork. §28.
- Nessun supporto VB6 è "completo" se rompe lingue esistenti. §19.
- Repo reali VB6 solo dopo che conformità + regressioni passano. §24.

## 5. Log decisioni & avanzamento
- 2026-08-13: audit architettura completato; piano scritto; mapping ontologia proposto. In attesa di via libera per Fase 1 (doc semantico) → Fase 2.
- 2026-08-14: lingua `vb6` registrata (types.ts/grammars.ts/tree-sitter.ts) ed extractor custom
  implementato (`vb6-extractor.ts`: code + designer + progetto). Typecheck `tsc --noEmit` verde.
  Committato in due unità logiche (`2d59953` docs, `530b584` feat) e **pushato su `origin/main`**.
  Nulla ancora testato: gli extractor erano da validare.
- 2026-08-14 (audit): ricostruito l'ambiente (`node_modules`/`dist` assenti), rieseguita la suite
  (**2901 test verdi, 0 regressioni**) e verificato il comportamento reale indicizzando mini-progetti VB6.
  L'estrazione si è rivelata più solida del previsto; la **resolution** no: risolve per nome e produce
  edge falsi (R1–R3). Riconfigurato il remote `upstream` che risultava dichiarato ma assente.
- 2026-08-14 (Fase 2): scritta la suite di conformità — 55 fixture + runner con oracolo.
  Stato misurato: **29 pass, 26 known-gap, 0 fail, 0 parse error**; 127 assertion attese, 97 soddisfatte,
  30 falsi negativi, 10 falsi positivi. Le fixture hanno scoperto tre difetti non previsti dall'audit:
  argomenti nominati (E3), label di riga (E4), identificatori dentro stringhe (E5).
  Scritto `VB6_SEMANTIC_MODEL.md` fissando il modello eventi e la provenance che il resolver dovrà produrre.
- 2026-08-14 (fix extractor): chiusi **E1–E12**. Il cambiamento strutturale è lo **scan differito**
  degli statement: le dichiarazioni possono seguire le procedure che le usano, quindi i riferimenti si
  analizzano quando l'intero file è noto — è ciò che permette di distinguere `Items(3)` da `Helper(3)`.
  Riscritto lo scanner degli identificatori su regole sintattiche VB6 (catene puntate, blocchi `With`,
  assegnazione vs chiamata implicita) invece che su euristiche di forma della riga.
  `provenance` non viene impostata dagli edge statici: il dominio è chiuso (`tree-sitter|scip|heuristic`)
  e la convenzione del repo è che solo i synthesizer la marchino — `VB6_SEMANTIC_MODEL.md` aggiornato.
  Suite: **44 pass / 11 gap / 0 fail**; nessuna regressione sulle altre lingue.
- 2026-08-14 (resolver): chiusi **R1–R6**. Aggiunti `vb6-resolver.ts` (scope) e
  `vb6-event-synthesizer.ts` (binding eventi). Il qualificatore viaggia in `candidates`, quindi
  **nessuna migrazione di schema**; scoperto e corretto un bug del core che scartava `candidates`
  in tre punti prima di `resolveOne`.
  Suite: **55/55 pass, 0 gap, 0 FN, 0 FP**; 2957+ test verdi, nessuna regressione.
  **Prossimo passo:** Fase 5–6 — MCP (verifica che il graph VB6 sia esposto correttamente),
  metriche §20 sul progetto reale, e i documenti finali `VB6_SUPPORT.md`, `VB6_ARCHITECTURE.md`,
  `VB6_LIMITATIONS.md`, `TEST_RESULTS.md`, `OPEN_QUESTIONS.md`, `VB6_FORK_ASSESSMENT.md`.
