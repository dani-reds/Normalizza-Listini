# normalize-listini

Motore locale per normalizzare listini marittimi in formato `xlsx` e `pdf` e produrre un file Excel finale compatibile con l'import nel TMS.

## Struttura del progetto

```text
normalize-listini/
|
|- Normalize-Listino.ps1
|- NormalizationRules.psd1
|- README.md
|- .gitignore
|
|- samples/
|  |- input/
|  `- expected-output/
|
|- output/
|- logs/
|
`- .codex/
   `- config.toml
```

## Cartelle

- `samples/input/`
  Qui puoi mettere file esempio da usare per test e sviluppo.

- `samples/expected-output/`
  Contiene output normalizzati di riferimento, utili per confronti e regressioni.

- `output/`
  Output generati durante l'uso quotidiano. Non sono pensati per essere versionati.

- `logs/`
  File temporanei, estrazioni testuali e artefatti di debug. Non sono pensati per essere versionati.

- `.codex/`
  Documentazione operativa e file di supporto per lavorare bene con Codex su questo progetto.

## File principali

- `Normalize-Listino.ps1`
  Script principale di parsing, normalizzazione e generazione dell'Excel finale.

- `NormalizationRules.psd1`
  Regole business: surcharge per carrier/direction, mapping UNLOCODE, matching specifici, ecc.

## Uso Base

Esempio di esecuzione:

```powershell
powershell -ExecutionPolicy Bypass -File .\Normalize-Listino.ps1 `
  -InputPath "C:\percorso\al\file.xlsx" `
  -OutputPath ".\output\file_normalized.xlsx"
```

Se `-OutputPath` non viene passato, lo script genera un file `_normalized.xlsx` accanto all'input.

## Note Pratiche

- La cartella esterna del progetto al momento puo` anche chiamarsi diversamente da `normalize-listini`.
- Se vuoi, in un secondo momento possiamo anche rinominare la cartella fisica del repository.
- La cache `UNLOCODE.lookup.clixml` viene rigenerata automaticamente quando serve.

## Git

Le cartelle `output/` e `logs/` sono ignorate da Git, cosi` la repo resta pulita.

Gli output di riferimento che vuoi conservare nel repository vanno invece in:

- `samples/expected-output/`
