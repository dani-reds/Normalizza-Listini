# Listini Normalization

Repository PowerShell per normalizzare listini marittimi `.xlsx` e `.pdf`
in un workbook Excel finale a schema fisso, importabile in Freightools.

## Obiettivo del progetto

Il flusso operativo del repository e' questo:

1. riconoscere la famiglia di layout del file sorgente;
2. convertire il contenuto in righe tariffarie standard;
3. scrivere un unico `.xlsx` finale con schema fisso;
4. congelare i comportamenti corretti tramite baseline approvati e regressione.

Il progetto privilegia stabilita', prudenza semantica e cambi localizzati.
Non e' un parser che deve inferire liberamente campi commerciali mancanti:
quando il file non esprime un valore in modo affidabile, il campo deve arrivare
da parametro esplicito oppure restare vuoto.

## Fonti di verita'

Usare queste fonti in quest'ordine:

1. `Normalize-Listino.ps1`
2. `docs/DecisionLog.md`
3. `samples/SampleManifest.psd1`
4. `samples/expected-output/*`
5. `scripts/Test-NormalizedWorkbook.ps1`
6. `scripts/Invoke-Phase1Validation.ps1`

`README.md` e' documentazione di orientamento. In caso di conflitto prevalgono
sempre codice, Decision Log, manifest e baseline approvati.

## Stato attuale

Stato verificato nel repository al 24 marzo 2026:

- parser unico operativo per `.xlsx` e `.pdf`;
- detector e adapter dedicati per piu' famiglie;
- fallback generico conservativo per workbook Excel;
- writer `.xlsx` finale a schema fisso;
- validator workbook-to-workbook;
- runner di regressione Phase 1;
- baseline formalmente approvati: `3`;
- ultimo stato regressione presente in repo: `PASS=3`, `FAIL=0`.

Baseline approvati e row count congelati:

- `altro-listino-1` -> `96`
- `baseline1` -> `104`
- `baseline2` -> `767`

## Struttura repository

```text
.
|-- Normalize-Listino.ps1
|-- NormalizationRules.psd1
|-- README.md
|-- docs/
|   `-- DecisionLog.md
|-- scripts/
|   |-- Invoke-Phase1Validation.ps1
|   `-- Test-NormalizedWorkbook.ps1
|-- samples/
|   |-- input/
|   |-- expected-output/
|   `-- SampleManifest.psd1
|-- output/
|-- logs/
`-- .codex/
```

Ruolo delle cartelle e dei file principali:

- `Normalize-Listino.ps1`: entrypoint unico del parser.
- `NormalizationRules.psd1`: regole statiche, alias UNLOCODE, surcharge attesi,
  pattern e apply targets.
- `docs/DecisionLog.md`: decisioni business e operative stabili.
- `scripts/`: validazione e regressione.
- `samples/input/`: input sorgente versionati dei baseline approvati.
- `samples/expected-output/`: golden file approvati.
- `samples/SampleManifest.psd1`: manifest dei baseline approvati.
- `output/`: artefatti runtime generati.
- `logs/`: log, diff report e file temporanei di debug.
- `.codex/`: materiale operativo di supporto per Codex.

## Famiglie supportate nel codice

Workbook `.xlsx` con detector dedicato:

- COSCO Far East
- COSCO Structured
- Evergreen RVS
- COSCO IET
- HMM CIT
- Baseline workbook family
- Baseline2 workbook family
- fallback generico workbook, da mantenere conservativo

PDF con detector dedicato:

- COSCO Canada
- COSCO South America
- COSCO Ipak
- Hapag-Lloyd

La presenza di un adapter in codice non significa automaticamente che la
famiglia sia gia' coperta da un baseline approvato.

## Principi guida

- Meglio un campo vuoto che un valore commerciale sbagliato.
- Meglio un falso negativo di detection che un falso positivo.
- Non inferire `Carrier` e `Reference` da segnali deboli.
- Le regole layout-specific non vanno globalizzate nel fallback.
- Nessun refactor largo senza necessita' reale.
- Ogni modifica al parser va verificata almeno con la Phase 1.

## Requisiti pratici

- Windows PowerShell 5.1 o PowerShell 7.
- Per i PDF: `pdftotext` disponibile nel `PATH` oppure in una delle posizioni
  controllate dallo script.
- Per il lookup UNLOCODE avanzato: `UNLOCODE.txt` disponibile via
  `-UnlocodePath`, variabile ambiente `UNLOCODE_LOOKUP_PATH` o percorso atteso
  dal parser.

La cache `UNLOCODE.lookup.clixml` viene usata come cache tecnica e non come
fonte di verita' business.

## Uso base

Esempio minimo su workbook Excel:

```powershell
powershell -ExecutionPolicy Bypass -File .\Normalize-Listino.ps1 `
  -InputPath ".\samples\input\Baseline2.xlsx" `
  -OutputPath ".\output\Baseline2_candidate.xlsx" `
  -Carrier "ONE" `
  -Direction "Export" `
  -Reference "GOAN00967A"
```

Esempio con `ValidityStartDate` esplicita:

```powershell
powershell -ExecutionPolicy Bypass -File .\Normalize-Listino.ps1 `
  -InputPath ".\samples\input\Baseline1.xlsx" `
  -OutputPath ".\output\Baseline1_candidate.xlsx" `
  -Carrier "ZIM" `
  -Direction "Export" `
  -Reference "BASELINE1-CANDIDATE" `
  -ValidityStartDate "01/08/2025"
```

Se `-OutputPath` non viene passato, lo script genera un file
`*_normalized.xlsx` accanto all'input.

## Validazione e regressione

Eseguire la regressione completa:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Invoke-Phase1Validation.ps1
```

Oppure forzando esplicitamente il manifest:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Invoke-Phase1Validation.ps1 `
  -ManifestPath ".\samples\SampleManifest.psd1"
```

La Phase 1:

- legge `samples/SampleManifest.psd1`;
- rigenera gli output approvati;
- salva i workbook generati in `output/validation/`;
- salva summary e log in `logs/validation/`;
- fallisce se ci sono differenze di schema, contenuto, row count o address.

## Baseline approvati

### `altro-listino-1`

- input: `samples/input/ALTRO_LISTINO_1.xlsx`
- expected: `samples/expected-output/ALTRO_LISTINO_1_normalized.xlsx`
- parametri: `Carrier=MSC`, `Direction=Export`, `Reference=ALTRO-REF-001`
- famiglia: generic classic pair matrix riconosciuta dal fallback con match stretto
- row count approvato: `96`

### `baseline1`

- input: `samples/input/Baseline1.xlsx`
- expected: `samples/expected-output/Baseline1_normalized.xlsx`
- parametri: `Carrier=ZIM`, `Direction=Export`,
  `Reference=BASELINE1-CANDIDATE`, `ValidityStartDate=01/08/2025`
- famiglia: workbook multi-sheet con fogli dry e reefer distinti
- row count approvato: `104`

### `baseline2`

- input: `samples/input/Baseline2.xlsx`
- expected: `samples/expected-output/Baseline2_normalized.xlsx`
- parametri: `Carrier=ONE`, `Direction=Export`, `Reference=GOAN00967A`
- famiglia: workbook dedicato con 2 fogli tariffari `F.A.K.*`
- row count approvato: `767`

## Workflow sicuro per nuovi listini

1. capire la famiglia di layout reale del file;
2. verificare se il file appartiene a un adapter esistente o a una nuova famiglia;
3. passare esplicitamente i parametri mancanti quando il layout non li esprime;
4. fare review manuale semantica prima di promuovere il risultato;
5. aggiungere baseline, manifest e Decision Log solo dopo approvazione;
6. rieseguire la Phase 1 prima di committare.

## Housekeeping del repository

Da tenere versionato:

- parser e regole (`Normalize-Listino.ps1`, `NormalizationRules.psd1`);
- documentazione contrattuale (`docs/DecisionLog.md`);
- script operativi in `scripts/`;
- input e output approvati in `samples/`;
- manifest dei baseline;
- documentazione operativa davvero utile.

Da trattare come artefatti runtime o temporanei:

- contenuto di `output/validation/`;
- log in `logs/validation/`;
- candidate output temporanei;
- diff report di debug;
- file temporanei `_tmp_*`.

Regola pratica: se un file e' riproducibile, non e' baseline approvato e non e'
fonte di verita', non dovrebbe restare come rumore nella repo di lavoro.

## Note finali

- `README.md` non sostituisce `docs/DecisionLog.md`.
- Un workbook formalmente valido puo' essere semanticamente sbagliato.
- Se una regola non e' supportata da codice e decisioni gia' approvate, va
  considerata pending e non inventata.
