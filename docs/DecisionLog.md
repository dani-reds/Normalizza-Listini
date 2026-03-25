# Decision Log

Questo file registra le decisioni stabili del progetto di normalizzazione listini.
Obiettivo: evitare che regole business e scelte di implementazione restino solo nella chat.

## 1. Regole generali del progetto

### Carrier e Reference
- Non inferire `Carrier` o `Reference` da testo debole, URL, note, remark o segnali indiretti.
- Quando il layout non fornisce un valore esplicito e affidabile, `Carrier` e `Reference` devono arrivare da parametri di invocazione espliciti.
- Se non vengono passati e il layout non li esprime chiaramente, lasciare il campo vuoto.

### Validity
- `Validity Start Date` e `Expiration Date` non devono mai essere uguali, a meno che il sorgente fornisca esplicitamente entrambe e siano davvero uguali.
- Se il sorgente fornisce una sola data di validità:
  - trattarla come `Expiration Date`
  - `Validity Start Date` deve arrivare da parametro esplicito, oppure restare vuoto
- Non inventare start date da euristiche non dichiarate.

### Formato date
- In output, quando una famiglia di layout supportata lo richiede, le date devono essere rese come `dd/MM/yyyy`.

### Price detail principale
- Il label corretto è `Ocean Freight - Containers`.
- Non usare uppercase completo tipo `OCEAN FREIGHT - CONTAINERS` nei layout che hanno già una decisione esplicita sul label.

### Commenti price detail
- Per ora i campi `Price Detail n - Comment` non devono essere popolati, salvo decisione futura diversa per layout specifici.
- I commenti di riga (`Comment`) possono restare se fanno parte di una scelta esplicita per una famiglia di layout.

### 40HC
- Se il listino non fornisce un valore esplicito per `40HC` ma fornisce `40 Box`, il valore `40 Box` va duplicato su `Cntrs 40' HC` solo nelle famiglie di layout in cui questa regola è stata approvata.
- Se il listino fornisce un valore esplicito distinto per `40HC`, usare quello e non duplicare.

### Fallback vs adapter dedicati
- Non forzare nel fallback generico workbook che appartengono chiaramente a una nuova famiglia di layout.
- Quando un workbook ha multi-sheet reali, validity per foglio, blocchi addizionali separati, strutture header specifiche o gestione distinta di dry/reefer, preferire detector stretto e converter dedicato.

### Prudenza semantica
- Meglio lasciare un campo vuoto che valorizzarlo in modo sbagliato.
- Meglio un falso negativo di detection che un falso positivo commerciale.
- Un file Excel formalmente valido puo` comunque essere semanticamente sbagliato.

## 2. Baseline approvati

### ALTRO_LISTINO_1
**Stato:** approvato  
**Input:** `samples/input/ALTRO_LISTINO_1.xlsx`  
**Expected output:** `samples/expected-output/ALTRO_LISTINO_1_normalized.xlsx`

**Parametri approvati**
- `Carrier = MSC`
- `Direction = Export`
- `Reference = ALTRO-REF-001`

**Decisioni specifiche**
- `Carrier` e `Reference` devono essere passati esplicitamente.
- Date in `dd/MM/yyyy`.
- Label principale: `Ocean Freight - Containers`.
- `Price Detail n - Comment` vuoti.
- Se il listino non distingue `40HC`, duplicare `40 Box` su `Cntrs 40' HC`.
- Queste regole non sono generiche per tutto il fallback `xlsx`: sono applicate alla sua famiglia di layout tramite condizione stretta.

### Baseline1
**Stato:** approvato  
**Input:** `samples/input/Baseline1.xlsx`  
**Expected output:** `samples/expected-output/Baseline1_normalized.xlsx`

**Parametri approvati**
- `Carrier = ZIM`
- `Direction = Export`
- `Reference = BASELINE1-CANDIDATE`
- `ValidityStartDate = 01/08/2025`

**Decisioni specifiche**
- Workbook family multi-sheet con fogli dry e reefer distinti.
- Validita` locale per foglio.
- Se il foglio fornisce una sola `VALIDITY`, quella vale come `Expiration Date`.
- `Validity Start Date` arriva solo da parametro esplicito.
- Date in `dd/MM/yyyy`.
- Label principale: `Ocean Freight - Containers`.
- `Transshipment Address` lasciato vuoto sui casi `via ...`.
- Addizionali per `ZIM` lasciati vuoti per ora, perche' non ancora supportati dalla matrice surcharge.

### Baseline2
**Stato:** approvato  
**Input:** `samples/input/Baseline2.xlsx`  
**Expected output:** `samples/expected-output/Baseline2_normalized.xlsx`

**Parametri approvati**
- `Carrier = ONE`
- `Direction = Export`
- `Reference = GOAN00967A`

**Decisioni specifiche**
- Processare solo i 2 fogli tariffari.
- Escludere `ADDITIONAL & LOCAL CHARGES` dalla generazione delle righe tariffarie.
- `Validity Start Date = 28/07/2024`.
- `Expiration Date = 31/08/2024`.
- Date in `dd/MM/yyyy`.
- Label principale: `Ocean Freight - Containers`.
- `Price Detail 1 = Cntr 20' Box`.
- `Price Detail 2 = Cntr 40' Box`.
- `Price Detail 3 = Cntrs 40' HC`.
- `NO SERVICE OPTION` da escludere per singola cella.
- `REMARKS` da mantenere solo come commento di riga.
- Nessuna estrazione surcharge aggressiva dai remark.

## 3. Baseline in corso / non ancora approvati

Al momento nessun baseline e` formalmente in promozione.

## 4. Regole di promozione baseline

### Quando promuovere un candidate a baseline
Promuovere un candidate a baseline solo se:
- il file e` semanticamente corretto
- i parametri usati per generarlo sono noti e salvati
- il row count e` plausibile
- le date sono corrette
- i price detail e le evaluation sono corretti
- non ci sono localita` chiaramente errate
- il comportamento e` abbastanza stabile da volerlo congelare

### Quando NON promuovere
Non promuovere se:
- il parser va in crash e viene “riparato” solo con una toppa tecnica
- ci sono dubbi business su validity, carrier, reference, 40HC, reefer o remarks
- l’output e` formalmente valido ma semanticamente incerto

## 5. Regole operative su Codex

### Modalita` di lavoro
- Usare cambi piccoli, localizzati e verificabili.
- Separare parsing, regole business, mapping output e validazione.
- Non cambiare il comportamento esistente senza confrontarlo con i casi gia` funzionanti.
- Fermarsi e segnalare le ambiguita` invece di inventare significati commerciali.

### Aree ad alto rischio
- Detection order degli adapter
- Validity logic
- Surcharge filtering
- UNLOCODE resolution
- Arbitrary handling
- Transshipment logic
- Output column mapping

### Evidenza minima da lasciare
- File toccati
- Logica aggiunta o modificata
- Test eseguiti
- Rischi residui

### Regressione
- Ogni baseline approvato va mantenuto in `samples/SampleManifest.psd1`.
- Ogni modifica al parser va verificata almeno con `scripts/Invoke-Phase1Validation.ps1`.

## 6. Debito tecnico noto

### Workbook packaging / OpenXML
- Esiste un tema separato di compatibilita`/compliance del packaging `.xlsx`.
- Non va mischiato con fix semantici dei layout salvo necessita` forte.
- **Stato:** pending, non risolto in questo log.

### Validazione di routing
- La validazione Phase 1 oggi rigenera e confronta i baseline approvati, ma non verifica ancora il detector-routing in modo dedicato.
- **Stato:** pending.

### Hapag PDF dry/std con blocchi door/inland
- Il file `Quotation_Q2603GOA03287_CASASC_003.pdf` non deve essere instradato sull'adapter Hapag PDF generico attuale.
- La famiglia Hapag dry/std PDF e` ora supportata operativamente in modo parziale tramite adapter dedicato, limitatamente alle pagine port-to-port che rientrano nello scope v1.
- Lo scope v1 include sia le pagine a tre colonne `20'STD / 40'STD / 40'HC`, sia le pagine layout-specific con `20'STD / 40'HC` senza `40'STD`.
- Per le pagine `20'STD / 40'HC` senza `40'STD`, la mappatura approvata e` layout-specific: `Cntr 20' Box` = `20'STD`, `Cntr 40' Box` lasciato vuoto, `Cntrs 40' HC` = `40'HC`, senza duplicazioni automatiche.
- Le route spezzate su due righe dopo `to` sono supportate dal solo adapter dedicato Hapag dry/std PDF.
- Le pagine door/inland/hazardous restano fuori scope e vengono ignorate intenzionalmente dal nuovo adapter dedicato, con `Write-Warning` esplicito che riporta numeri pagina e motivo dello skip.
- Il workbook risultante puo` quindi essere parziale rispetto al PDF originale completo; questo comportamento e` intenzionale e non modifica l'adapter Hapag legacy.
- La v1 port-to-port e` stata validata positivamente tramite harness separato che isola solo le pagine supportate; la regressione minima dei baseline approvati resta verde (`PASS=3 FAIL=0`).
- Il supporto funzionale completo del file misto resta **pending**. Un eventuale supporto futuro delle pagine door/inland dovra` restare separato, prudente e senza allargare implicitamente la semantica della v1 port-to-port.

### Route assertions piu` ricche
- La validazione attuale e` centrata sul confronto workbook-to-workbook.
- Route assertions e check business piu` granulari sono ancora un’estensione futura.
- **Stato:** pending.

### Pulizia globale UNLOCODE
- Alcune famiglie di layout richiedono alias locali per casi ambigui o per voci errate nel lookup sorgente.
- Una pulizia/normalizzazione piu` sistematica del lookup non e` ancora stata fatta.
- **Stato:** pending.
