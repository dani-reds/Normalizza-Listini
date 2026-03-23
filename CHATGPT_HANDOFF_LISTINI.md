# Handoff Per Revisione Esterna

## Obiettivo

Sto costruendo un motore locale che prende listini marittimi in formato `xlsx` o `pdf`, estrae le informazioni utili e genera un file Excel normalizzato da importare in un TMS.

Il motore vive qui:

- `C:\Users\DanieleRossi\Documents\New project\Normalize-Listino.ps1`
- `C:\Users\DanieleRossi\Documents\New project\NormalizationRules.psd1`

Lookup UNLOCODE:

- sorgente principale: `C:\Users\DanieleRossi\Desktop\n8n\n8n BA Extractor\UNLOCODE.txt`
- cache locale: `C:\Users\DanieleRossi\Documents\New project\UNLOCODE.lookup.clixml`

## Cosa Fa Oggi

Il parser:

- legge `xlsx` e `pdf`
- riconosce layout diversi con adapter dedicati
- estrae tratte, validità, riferimento, transit time e nolo base
- converte le località in UNLOCODE
- filtra le surcharge in base a `carrier + import/export`
- scrive un Excel finale con struttura fissa per il TMS

Ogni riga del file finale rappresenta una tratta `From Address -> To Address` con:

- date di validità
- carrier
- reference
- comment
- transshipment address se presente
- fino a 25 `Price Detail`

## Matrice Surcharge Attuale

Le surcharge attese sono già modellate in `NormalizationRules.psd1` per:

- COSCO
- EVERGREEN
- MSC
- OOCL
- YANG MING
- ARKAS
- TARROS
- HMM
- HAPAG-LLOYD

Per HMM:

- Export: `EES`, `ECC`, `PEK`, `KLS`
- Import: `ECC`, `STF`, `BUC`, `ETS`

## Adapter Attualmente Supportati

- listino Excel “semplice” storico
- COSCO Far East workbook
- COSCO structured workbook (`Tariffe`, `Sovrapprezzi`, `POL_AddOn`)
- COSCO IPAK PDF
- COSCO Canada PDF
- COSCO South America PDF
- COSCO IET workbook
- EVERGREEN FAK RVS workbook
- HAPAG-LLOYD quotation PDF
- HMM CIT workbook `CIT2200102-amd135.xlsx`

## Output Già Generati

- `C:\Users\DanieleRossi\Documents\New project\COSCO_Rate_Export_FAR_EAST_2026Q2_normalized.xlsx`
- `C:\Users\DanieleRossi\Documents\New project\COSCO_Rate_Export_IPAK_2026Q2_normalized.xlsx`
- `C:\Users\DanieleRossi\Documents\New project\COSCO_Rate_Export_IPAK_2026Q2_from_pdf_normalized.xlsx`
- `C:\Users\DanieleRossi\Documents\New project\IT_Guideline_20260315_20260331_FAK_RVS_normalized.xlsx`
- `C:\Users\DanieleRossi\Documents\New project\iet_tariff_March_2026_normalized.xlsx`
- `C:\Users\DanieleRossi\Documents\New project\Quotation_Q2603GOA02151_GDT_002_normalized.xlsx`
- `C:\Users\DanieleRossi\Documents\New project\CIT2200102-amd135_normalized.xlsx`

## Focus: Adapter HMM CIT

File sorgente:

- `C:\Users\DanieleRossi\Downloads\CIT2200102-amd135.xlsx`

Struttura workbook:

- `Head`
- `Freight`
- `Arb Addon`
- `Subject to`
- `DEMDET`

### Dati letti da `Head`

- Contract Number: `CIT2200102`
- Amend No.: `135`
- Effective: `2026-03-22`
- Contract Duration From: `2022-06-01`
- Contract Duration To: `2026-03-31`
- Service Coverage: `MEDW,NCPW`

### Decisioni implementate per CIT

1. Carrier forzato a `HMM`
2. Direction inferita come `IMPORT`
3. Reference scritto come `CIT2200102-amd135`
4. Validità usata: `Effective` -> `Contract Duration To`, cioè `2026-03-22` -> `2026-03-31`
5. I codici origine/destinazione del foglio `Freight` sono già UNLOCODE e vengono usati direttamente
6. Gli `Origin Arb` e `Dest Arb` non vengono sommati al nolo base: vengono scritti come `INLAND FREIGHT` separato
7. Per `Origin Arb`, il `Transshipment Address` viene valorizzato con `Via` se presente, altrimenti con `Base Rate Code`
8. Per `Dest Arb`, stesso criterio del punto sopra
9. Le surcharge HMM considerate nel CIT sono quelle coerenti con `HMM Import`

### Regole surcharge implementate per CIT

Dal foglio `Subject to`:

- `ECC`
- `STF`
- `ETS`

Esclusioni deliberate:

- `BUC` non viene scritto perché nel file è indicato come `Inclusive`
- `PEK` non viene scritto perché è nella matrice HMM Export, non Import
- `KLS` non viene scritto perché è nella matrice HMM Export, non Import
- `HGR`, `OW`, `DG Premium`, `45' add-on` non vengono scritti perché fuori scope rispetto alla matrice surcharge richiesta

### Logica ETS implementata per CIT

Per `MEDW`:

- origini `CN*` -> `ETS USD 87/TEU`
- origini non `CN*` -> `ETS EUR 76/TEU`

Per `NCPW`:

- origini `CN*` -> `ETS USD 49/TEU`
- origini non `CN*` -> `ETS EUR 42/TEU`

Il parser scrive solo il valore `TEUS` della taglia 20, coerente con gli altri output già costruiti.

### Risultato CIT attuale

File generato:

- `C:\Users\DanieleRossi\Documents\New project\CIT2200102-amd135_normalized.xlsx`

Stato verificato:

- `2.657` righe dati
- surcharge presenti: `ECC`, `STF`, `ETS`
- surcharge assenti: `PEK`, `KLS`, `BUC`
- `INLAND FREIGHT` presente per `Origin Arb` e `Dest Arb`

Esempi verificati:

- `CNNGB -> FRFOS` con `ECC`, `STF`, `ETS USD 87`
- `KRPUS -> FRFOS` con `ETS EUR 76`
- `CNNKG -> FRFOS` come `Origin Arb`, con `Transshipment Address = CNSHA`
- `CNNGB -> ITTRS` come `Dest Arb`, con `Transshipment Address = GRPIR`

## Domande Su Cui Vorrei Un Parere Critico

1. Trattare il workbook `CIT2200102-amd135.xlsx` come `HMM Import` è corretto secondo te, dato che le tratte sono Far East -> Europe?
2. La validità giusta è davvero `Effective -> Contract Duration To`, oppure andrebbe usato un altro criterio?
3. Gli `Arb Addon` è meglio tenerli come `INLAND FREIGHT` separato oppure andrebbero sommati al nolo base?
4. Per il TMS è sensato scrivere solo le surcharge coerenti con la matrice `carrier + direction`, anche se il contratto contiene altre note tariffarie?
5. Le surcharge `PEK` e `KLS` in questo file vanno davvero escluse perché le sto trattando come Export-only, oppure c’è un caso in cui dovrebbero comparire anche qui?
6. Il fatto di usare solo la componente `TEUS` per `ECC/STF/ETS` è corretto per il TMS oppure dovrei invece generare anche le versioni `40'` come dettagli separati?
7. Ci sono rischi nel valorizzare `Transshipment Address` con il `Base Rate Code` quando il campo `Via` è vuoto?
8. Ci sono miglioramenti architetturali evidenti nel modo in cui sto organizzando adapter, regole surcharge e output?

## Richiesta

Vorrei una review critica, non solo conferme. Se noti errori di interpretazione, regressioni potenziali o un modo migliore di modellare:

- validity
- direction
- arbitrary origin/destination
- surcharge filtering
- struttura output TMS

segnalamelo in modo esplicito.
