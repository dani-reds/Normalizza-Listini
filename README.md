# Listini Normalization

Repository PowerShell per normalizzare listini marittimi `.xlsx` e `.pdf`
in un workbook Excel finale a schema fisso, importabile in Freightools.

Il progetto va gestito in modo conservativo: niente refactor larghi, niente
inferenze deboli, niente estensioni semantiche non approvate.

## Project Overview

Il repository serve a:

1. riconoscere la famiglia di layout del sorgente;
2. convertire il contenuto in righe tariffarie standard;
3. scrivere un unico workbook `.xlsx` finale con schema fisso;
4. mantenere stabile il comportamento tramite baseline approvati e regressione.

Principio guida: il parser non deve inventare significati commerciali mancanti.
Se il layout non fornisce un valore in modo esplicito e affidabile, il campo
deve arrivare da parametro oppure restare vuoto.

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

Distinzione importante:

- supportato in codice != approvato formalmente;
- approvato formalmente = comportamento congelato in manifest + expected output;
- i baseline PDF Hapag approvati congelano solo i casi presenti nel manifest e
  non vanno generalizzati oltre quanto gia' validato.

## Quick Start / Come usare il progetto

Nota: passare esplicitamente `Carrier`, `Direction`, `Reference` e
`ValidityStartDate` quando il layout non li fornisce in modo deterministico o
quando si vuole congelare un run in modo stabile.

### 1. Normalizzare un `.xlsx`

```powershell
powershell -ExecutionPolicy Bypass -File .\Normalize-Listino.ps1 `
  -InputPath ".\samples\input\Baseline2.xlsx" `
  -OutputPath ".\output\Baseline2_candidate.xlsx" `
  -Carrier "ONE" `
  -Direction "Export" `
  -Reference "GOAN00967A"
```

### 2. Normalizzare un `.pdf`

```powershell
powershell -ExecutionPolicy Bypass -File .\Normalize-Listino.ps1 `
  -InputPath ".\samples\input\Quotation_Q2603GOA02143_GDT_002 (1).pdf" `
  -OutputPath ".\output\Quotation_Q2603GOA02143_GDT_002_candidate.xlsx" `
  -Carrier "HAPAG-LLOYD" `
  -Direction "Export" `
  -Reference "Q2603GOA02143" `
  -ValidityStartDate "2026-04-01"
```

### 3. Lanciare la regressione Phase 1

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Invoke-Phase1Validation.ps1
```

### 4. Promuovere un nuovo baseline

1. generare un candidate output con parametri espliciti se servono;
2. fare review manuale semantica;
3. salvare input sorgente e expected output nelle cartelle `samples/`;
4. aggiornare `samples/SampleManifest.psd1`;
5. aggiornare `docs/DecisionLog.md` solo se c'e' una decisione stabile nuova;
6. rieseguire la Phase 1 prima di considerare il baseline approvato.

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

Ruolo delle aree principali:

- `Normalize-Listino.ps1`: entrypoint unico del parser.
- `NormalizationRules.psd1`: regole statiche, alias UNLOCODE, surcharge attesi.
- `docs/DecisionLog.md`: decisioni business e operative stabili.
- `scripts/`: validazione e helper operativi.
- `samples/input/`: input dei baseline approvati.
- `samples/expected-output/`: golden file approvati.
- `samples/SampleManifest.psd1`: elenco contrattuale dei baseline approvati.
- `output/`, `logs/`: artefatti runtime rigenerabili.
- `tmp/`: scratch locale, non fonte di verita'.
- `.codex/`: documentazione operativa non contrattuale.

## Main Parsing Architecture

Per i workbook `.xlsx`, il parser espande il pacchetto OpenXML, legge fogli e
shared strings, applica detector in ordine stretto e instrada il file al
converter corretto. Le famiglie workbook riconosciute nel codice sono:

- COSCO Far East
- COSCO Structured
- Evergreen RVS
- COSCO IET
- HMM CIT
- Baseline workbook family
- Baseline2 workbook family
- fallback generico workbook

Per i `.pdf`, il parser estrae il testo tramite `pdftotext`, applica detector
PDF in ordine stretto e usa un adapter dedicato oppure fallisce in modo
esplicito. Le famiglie PDF riconosciute nel codice sono:

- COSCO Canada
- COSCO South America
- COSCO Ipak
- flussi Hapag PDF specifici

Per Hapag, il README resta volutamente prudente: i casi formalmente approvati
sono quelli congelati nel manifest e negli expected output, senza generalizzare
il supporto oltre quanto gia' validato.

Tutte le famiglie convergono nello stesso schema output:

- 18 colonne meta iniziali;
- 25 blocchi di `Price Detail`;
- 8 colonne per ciascun blocco.

Zone ad alto rischio quando si cambia codice: detection order, validity logic,
UNLOCODE resolution, surcharge filtering, transshipment, output mapping,
writer OpenXML.

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

Per i dettagli e le decisioni layout-specific complete, fare sempre riferimento
a `docs/DecisionLog.md`.

## Approved Baselines

Baseline approvati attuali nel manifest:

- `altro-listino-1` - `xlsx` - `96` righe
- `baseline1` - `xlsx` - `104` righe
- `baseline2` - `xlsx` - `767` righe
- `hapag-pdf-q2603goa03287-casasc-003` - `pdf` - `8` righe
- `hapag-pdf-q2603goa02143-gdt-002-1` - `pdf` - `36` righe
- `hapag-pdf-q2603goa02149-gdt-002-1` - `pdf` - `35` righe

Per percorsi, parametri e note complete, usare `samples/SampleManifest.psd1`.
I baseline Hapag approvati fissano solo i casi presenti nel manifest e non
vanno usati per dedurre supporto semantico a varianti ulteriori.

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

La presenza di un adapter in codice non equivale a baseline approvato.

## Validation and Regression Workflow

- `samples/SampleManifest.psd1` e' l'elenco dei baseline approvati.
- `scripts/Invoke-Phase1Validation.ps1` rigenera tutti i baseline del manifest,
  salva i generated workbook in `output/validation/` e i log in
  `logs/validation/`.
- `scripts/Test-NormalizedWorkbook.ps1` confronta schema, row count, contenuto
  e address (`From`, `To`, `Transshipment`).
- `PASS` = nessuna differenza rilevata rispetto ai baseline approvati.
- `FAIL` = differenza rispetto al baseline o errore di generazione/validazione.
- `scripts/Invoke-HapagDryStdFixtureValidation.ps1` resta un helper operativo,
  non una fonte di verita' contrattuale.

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

Per Hapag vale una regola aggiuntiva: non generalizzare il supporto oltre i PDF
gia' formalmente approvati nel manifest.

## Housekeeping / Repository Hygiene

Classificazione pratica:

- file produttivi: parser, regole, script operativi;
- baseline approvati: `samples/input/*` e `samples/expected-output/*` presenti
  nel manifest;
- documentazione contrattuale: Decision Log, manifest, expected output;
- documentazione operativa: `README.md`, `.codex/`, helper script;
- artefatti runtime: `output/`, `logs/`;
- scratch locale: `tmp/`.

Su GitHub devono restare centrali e visibili:

- `README.md`
- `docs/DecisionLog.md`
- `samples/SampleManifest.psd1`
- `scripts/Invoke-Phase1Validation.ps1`

Non devono sembrare fonte di verita':

- `output/`
- `logs/`
- `tmp/`
- documentazione operativa in `.codex/`
