# Listini Normalization

Repository PowerShell per normalizzare listini marittimi `.xlsx` e `.pdf`
in un workbook Excel finale a schema fisso, importabile in Freightools.

Il progetto va gestito in modo conservativo: niente refactor larghi, niente
inferenze deboli, niente estensioni semantiche non approvate.

## Project Overview

L'obiettivo del repository e' questo:

1. riconoscere la famiglia di layout del sorgente;
2. convertire il contenuto in righe tariffarie standard;
3. scrivere un unico workbook `.xlsx` finale con schema fisso;
4. congelare i comportamenti corretti tramite baseline approvati e regressione.

Il parser non deve inventare significati commerciali mancanti. Se il layout non
fornisce un valore in modo esplicito e affidabile, il campo deve arrivare da
parametro oppure restare vuoto.

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

## Current Status

Stato verificato nel repository al 25 marzo 2026:

- parser unico operativo per `.xlsx` e `.pdf`;
- detector e adapter dedicati piu' fallback conservativo;
- runner di regressione Phase 1 funzionante;
- baseline approvati totali: `6`;
- composizione baseline approvati: `3` workbook `.xlsx` storici + `3` baseline
  PDF Hapag approvati;
- ultimo stato di regressione verificato: `PASS=6 FAIL=0`.

I baseline PDF Hapag approvati congelano solo il comportamento rappresentato da
manifest, expected output e Decision Log. La loro promozione non deve essere
letta come supporto generale a tutte le varianti PDF Hapag oltre quanto
formalmente approvato.

## Repository Structure

```text
.
|-- Normalize-Listino.ps1
|-- NormalizationRules.psd1
|-- README.md
|-- docs/
|   `-- DecisionLog.md
|-- scripts/
|   |-- Invoke-Phase1Validation.ps1
|   |-- Test-NormalizedWorkbook.ps1
|   `-- Invoke-HapagDryStdFixtureValidation.ps1
|-- samples/
|   |-- input/
|   |-- expected-output/
|   `-- SampleManifest.psd1
|-- output/
|-- logs/
|-- tmp/
`-- .codex/
```

Ruolo pratico delle aree principali:

- `Normalize-Listino.ps1`: entrypoint unico del parser e della scrittura output.
- `NormalizationRules.psd1`: alias UNLOCODE, surcharge attesi, pattern e
  `ApplyTargets`.
- `docs/DecisionLog.md`: decisioni business e operative stabili.
- `scripts/Invoke-Phase1Validation.ps1`: regressione principale basata sul
  manifest.
- `scripts/Test-NormalizedWorkbook.ps1`: validator workbook-to-workbook.
- `scripts/Invoke-HapagDryStdFixtureValidation.ps1`: helper operativo per la
  fixture validation Hapag dry/std; utile ma non fonte di verita' primaria.
- `samples/input/`: input versionati dei baseline approvati.
- `samples/expected-output/`: golden file approvati.
- `samples/SampleManifest.psd1`: elenco contrattuale dei baseline approvati.
- `output/` e `logs/`: artefatti runtime rigenerabili.
- `tmp/`: scratch locale / materiale di lavoro, non fonte di verita'.
- `.codex/`: documentazione operativa non contrattuale.

## Main Parsing Architecture

### Workbook `.xlsx`

Per i workbook Excel il parser:

1. espande il pacchetto OpenXML;
2. legge workbook, fogli e shared strings;
3. applica detector in ordine stretto;
4. instrada il file al converter corretto;
5. converte le route in righe finali;
6. scrive un unico foglio `Sheet` a schema fisso.

Famiglie workbook riconosciute nel codice:

- COSCO Far East
- COSCO Structured
- Evergreen RVS
- COSCO IET
- HMM CIT
- Baseline workbook family
- Baseline2 workbook family
- fallback generico workbook

### PDF

Per i PDF il parser:

1. estrae il testo tramite `pdftotext`;
2. applica detector PDF in ordine stretto;
3. usa un adapter dedicato oppure fallisce in modo esplicito.

Famiglie PDF riconosciute nel codice:

- COSCO Canada
- COSCO South America
- COSCO Ipak
- Hapag dry/std PDF dedicato
- Hapag PDF legacy quotation adapter

Il codice contiene anche una guardia esplicita per varianti Hapag non ancora
supportate in modo sicuro.

### Convergenza finale

Tutte le famiglie convergono nello stesso schema output:

- 18 colonne meta iniziali;
- 25 blocchi di `Price Detail`;
- 8 colonne per ciascun blocco.

Le aree di rischio piu' alto quando si cambia codice sono:

- detection order;
- validity logic;
- location / UNLOCODE resolution;
- surcharge filtering;
- transshipment logic;
- output column mapping;
- writer OpenXML.

## Approved Business Rules

Le regole gia' approvate e da non cambiare silenziosamente sono:

- non inferire `Carrier` o `Reference` da testo debole, note, URL o remark;
- se il layout non esprime `Carrier`/`Reference` in modo affidabile, devono
  arrivare da parametri espliciti oppure restare vuoti;
- `Validity Start Date` e `Expiration Date` non devono coincidere salvo caso
  esplicito nel sorgente;
- se esiste una sola data di validita', trattarla come `Expiration Date`;
- `Validity Start Date` arriva solo da parametro esplicito oppure resta vuoto;
- nelle famiglie approvate che lo richiedono, le date in output vanno rese come
  `dd/MM/yyyy`;
- il label corretto e' `Ocean Freight - Containers`;
- `Price Detail n - Comment` deve restare vuoto salvo futura decisione
  layout-specific;
- la duplicazione `40 Box -> Cntrs 40' HC` e' ammessa solo nelle famiglie in
  cui e' stata approvata;
- il fallback generico non va allargato per inglobare nuove famiglie;
- meglio lasciare un campo vuoto che valorizzarlo in modo commerciale errato.

## Approved Baselines

Baseline approvati attuali nel manifest:

### Workbook `.xlsx`

- `altro-listino-1`
  - input: `samples/input/ALTRO_LISTINO_1.xlsx`
  - expected: `samples/expected-output/ALTRO_LISTINO_1_normalized.xlsx`
  - parametri: `Carrier=MSC`, `Direction=Export`, `Reference=ALTRO-REF-001`
  - row count approvato: `96`
  - famiglia: generic classic pair matrix nel fallback, tramite match stretto

- `baseline1`
  - input: `samples/input/Baseline1.xlsx`
  - expected: `samples/expected-output/Baseline1_normalized.xlsx`
  - parametri: `Carrier=ZIM`, `Direction=Export`,
    `Reference=BASELINE1-CANDIDATE`, `ValidityStartDate=01/08/2025`
  - row count approvato: `104`
  - famiglia: workbook multi-sheet con fogli dry e reefer distinti

- `baseline2`
  - input: `samples/input/Baseline2.xlsx`
  - expected: `samples/expected-output/Baseline2_normalized.xlsx`
  - parametri: `Carrier=ONE`, `Direction=Export`, `Reference=GOAN00967A`
  - row count approvato: `767`
  - famiglia: workbook dedicato con 2 fogli tariffari `F.A.K.*`

### PDF Hapag approvati

- `hapag-pdf-q2603goa03287-casasc-003`
  - input: `samples/input/Quotation_Q2603GOA03287_CASASC_003.pdf`
  - expected: `samples/expected-output/Quotation_Q2603GOA03287_CASASC_003_normalized.xlsx`
  - parametri: `Carrier=HAPAG-LLOYD`, `Direction=Export`,
    `Reference=Q2603GOA03287`, `ValidityStartDate=2026-04-01`
  - row count approvato: `8`
  - famiglia: baseline approvato del dedicated Hapag dry/std PDF adapter, nei
    limiti del comportamento port-to-port approvato

- `hapag-pdf-q2603goa02143-gdt-002-1`
  - input: `samples/input/Quotation_Q2603GOA02143_GDT_002 (1).pdf`
  - expected: `samples/expected-output/Quotation_Q2603GOA02143_GDT_002 (1)_normalized.xlsx`
  - parametri: `Carrier=HAPAG-LLOYD`, `Direction=Export`,
    `Reference=Q2603GOA02143`, `ValidityStartDate=2026-04-01`
  - row count approvato: `36`
  - famiglia: baseline approvato del dedicated Hapag dry/std PDF adapter dopo
    la regola approvata di inclusione basata su endpoint risolti in modo
    affidabile

- `hapag-pdf-q2603goa02149-gdt-002-1`
  - input: `samples/input/Quotation_Q2603GOA02149_GDT_002 (1).pdf`
  - expected: `samples/expected-output/Quotation_Q2603GOA02149_GDT_002 (1)_normalized.xlsx`
  - parametri: `Carrier=HAPAG-LLOYD`, `Direction=Export`,
    `Reference=Q2603GOA02149`, `ValidityStartDate=2026-04-01`
  - row count approvato: `35`
  - famiglia: baseline PDF Hapag reefer, congelato nel manifest e negli
    expected output approvati

I baseline Hapag approvati fissano il comportamento oggi validato. Non vanno
usati per dedurre supporto semantico a varianti Hapag ulteriori non presenti nel
manifest.

## Pending / Not Yet Approved

Elementi ancora pending o da trattare come debito tecnico:

- nessun baseline formalmente "in promozione" oltre quelli gia' nel manifest;
- supporto funzionale completo dei casi Hapag misti / parziali fuori dallo scope
  approvato;
- validazione dedicata del detector-routing;
- route assertions e check business piu' granulari;
- pulizia globale UNLOCODE;
- tema OpenXML / packaging `.xlsx`;
- eventuale pulizia di materiale scratch storico in `tmp/`.

La presenza di un adapter in codice non equivale a baseline approvato. Fino a
quando una famiglia non e' congelata tramite manifest + expected output, il suo
comportamento va considerato non ancora formalmente stabilizzato.

## Validation and Regression Workflow

### Manifest

`samples/SampleManifest.psd1` e' l'elenco dei baseline approvati. Un sample
entra nella regressione solo se e' registrato li'.

### Runner Phase 1

`scripts/Invoke-Phase1Validation.ps1`:

- legge il manifest;
- rigenera ogni output approvato;
- salva i workbook generati in `output/validation/`;
- salva summary e log in `logs/validation/`;
- fallisce se almeno un baseline non combacia.

### Validator workbook-to-workbook

`scripts/Test-NormalizedWorkbook.ps1` confronta:

- schema e intestazioni;
- row count;
- contenuto celle;
- address (`From`, `To`, `Transshipment`).

`PASS` significa nessuna differenza strutturale o contenutistica rilevata.
`FAIL` significa divergenza rispetto al baseline approvato oppure errore di
generazione/validazione.

### Fixture validation Hapag

`scripts/Invoke-HapagDryStdFixtureValidation.ps1` e' un helper operativo per la
validazione di fixture Hapag dry/std su pagine selezionate. E' utile per lavoro
mirato, ma non sostituisce la Phase 1 e non e' una fonte di verita' contrattuale.

## Safe Workflow for New Listini

1. identificare prima la famiglia di layout reale;
2. decidere se il file rientra in un adapter esistente o in una nuova famiglia;
3. passare esplicitamente `Carrier`, `Direction`, `Reference` e
   `ValidityStartDate` quando il layout non li fornisce in modo deterministico;
4. generare un candidate output;
5. fare review manuale semantica su date, location, evaluation, surcharge,
   commenti e copertura rotte;
6. promuovere a baseline solo dopo approvazione esplicita;
7. aggiornare manifest e Decision Log solo quando il comportamento e' da
   congelare;
8. rieseguire la Phase 1 prima di committare.

Per Hapag vale una regola aggiuntiva: non generalizzare il supporto oltre le
famiglie e i PDF gia' formalmente approvati nel manifest.

## Housekeeping / Repository Hygiene

Classificazione pratica dei file:

- file produttivi: parser, regole, script operativi;
- baseline approvati: `samples/input/*` e `samples/expected-output/*` registrati
  nel manifest;
- documentazione contrattuale: `docs/DecisionLog.md`, manifest, expected output;
- documentazione operativa/non contrattuale: `README.md`, `.codex/`, helper
  operativi;
- artefatti runtime/generati: `output/`, `logs/`;
- scratch locale: `tmp/`.

Da mantenere centrali e visibili su GitHub:

- `README.md`
- `docs/DecisionLog.md`
- `samples/SampleManifest.psd1`
- `scripts/Invoke-Phase1Validation.ps1`

Da non far percepire come fonte di verita':

- `output/`
- `logs/`
- `tmp/`
- documentazione operativa in `.codex/`

Regola pratica: se un file e' rigenerabile, non e' baseline approvato e non e'
referenziato dalle fonti di verita', non deve restare come rumore nella working
area principale della repository.
