# DIGIL Monitoring Dashboard

**Tool di monitoraggio degli apparati DIGIL IoT installati sulla rete di trasmissione elettrica.**

Sviluppato per Terna S.p.A. — Team IoT

---

## Cosa fa

Questo tool sostituisce il workflow manuale basato su Excel per il monitoraggio quotidiano dei ~986 dispositivi DIGIL installati sui tralicci della rete elettrica. Importa il file Excel di monitoraggio, analizza lo stato di ogni dispositivo, genera alert automatici per le situazioni critiche, e presenta tutto in una dashboard desktop con filtri avanzati.

---

## Installazione

```bash
cd digil-monitoring-pyqt
pip install -r requirements.txt
python main.py
```

Al primo avvio, clicca **Importa Excel** e seleziona il file `Monitoraggio_APPARATI_DIGIL_INSTALLATI_*.xlsx`.

⚠️ **Dopo aggiornamenti che modificano lo schema DB**, cancellare `data/digil_monitoring.db` prima del primo import.

---

## Sorgente Dati

Il tool legge due sheet dal file Excel di monitoraggio:

**Sheet "Stato"** (fonte primaria, ~987 righe): contiene l'anagrafica completa di ogni dispositivo (linea, sostegno, fornitore, DT, IP, tipo installazione), i check diagnostici (MongoDB, batteria, porta), le informazioni su malfunzionamenti e ticket, e le colonne storiche di availability (formato "AVAILABILITY DD mmm", es. "AVAILABILITY 17 dic").

**Sheet "Av Status"** (~1029 righe): contiene l'availability giornaliera recente con codici numerici per gli ultimi 5 giorni circa.

**Regola importante**: Solo i dispositivi presenti nello sheet Stato vengono importati. Lo sheet Av Status contiene circa 45 dispositivi in più che corrispondono ad apparati non collaudati; questi vengono automaticamente ignorati.

---

## I 4 Stati di Availability

Ogni dispositivo ha **esattamente 4 possibili stati** di availability:

| Codice | Stato | Significato | Colore Timeline | Norm |
|--------|-------|------------|----------------|------|
| 1 | **COMPLETE** | Tutte le metriche presenti | Verde scuro (#2E7D32) | OK |
| 2 | **AVAILABLE** | Metriche meteo + almeno una di tiro, ma non tutte | Verde chiaro (#66BB6A) | OK |
| 3 | **NOT AVAILABLE** | Manca almeno una metrica meteo e un sensore di tiro | Giallo (#F9A825) | KO |
| 4 | **NO DATA** | Il dispositivo non comunica misure (ma potrebbe arrivare allarmi e diagnostiche) | Rosso (#C62828) | KO |

I codici numerici vengono dallo sheet Av Status. Nello sheet Stato, i valori testuali vengono normalizzati:
- `ON` → AVAILABLE (OK)
- `OFF` → NOT AVAILABLE (KO)
- `COMPLETE` → COMPLETE (OK)

Nella timeline del dettaglio device, ogni quadrato è colorato secondo il suo stato specifico, con tooltip che mostra data e stato.

---

## Concetti Chiave

### Health (stato di salute)

- **OK**: availability OK e nessuna diagnostica critica
- **KO**: availability KO
- **DEGRADED**: availability OK ma diagnostiche in allarme (porta aperta)
- **UNKNOWN**: nessun dato

### Trend 7 giorni

`■■■□□□■` = OK 3gg, KO 3gg, OK 1gg. Visualizza la stabilità del dispositivo.

### Giorni

Giorni consecutivi nello stato attuale. KO da 14 giorni → "14".

### Sotto Corona

Installazione senza sensori di tiro. Badge **SC**. Non genera alert su sensori assenti.

### Fornitori

- **INDRA** (Lotto 1 — IndraOlivetti): ~491 dispositivi
- **MII** (Lotto 2 — TelebitMarini): ~319 dispositivi
- **SIRTI** (Lotto 3): ~176 dispositivi

### Diagnostiche

| Check | Significato |
|-------|------------|
| **Mongo** | Invio dati telemetria a MongoDB |
| **Batteria** | Stato batteria |
| **Porta** | Sensore porta quadro |

---

## Sistema di Alert (8 regole)

| Severity | Colore | Significato |
|----------|--------|------------|
| **CRITICAL** | Rosso | Intervento immediato |
| **HIGH** | Arancione | Problema serio |
| **MEDIUM** | Giallo | Da monitorare |
| **LOW** | Verde | Informativo |

### Le 8 regole

1. **KO_NO_TICKET** — KO da 3+ giorni senza ticket attivo (Aperto/Interno). CRITICAL ≥7gg, HIGH 3-6gg. Sotto corona con Mongo OK → LOW.
2. **NEW_KO** — Passaggio da OK (2+ gg) a KO. HIGH senza ticket, MEDIUM con ticket.
3. **DOOR_ALARM** — Porta KO senza ticket specifico. MEDIUM/LOW.
4. **BATTERY_ALARM** — Batteria KO. HIGH/MEDIUM.
5. **INTERMITTENT** — 3+ cambi stato in 7 giorni. MEDIUM.
6. **RECOVERED** — Tornato OK dopo 2+ giorni KO. LOW.
7. **OPEN_TICKET_OK** — OK da 5+ giorni con ticket ancora aperto. LOW.
8. **NO_DATA** — Ultimo giorno di availability = NO DATA (il dispositivo non comunica misure). MEDIUM se 1-4 giorni, HIGH se 5+ giorni. Solo senza ticket attivo.

### Logica Ticket Attivi

Un ticket è considerato **attivo** solo se:
- `ticket_id` è presente (non vuoto) **E**
- `ticket_stato` è `Aperto` o `Interno`

Ticket con stato `Chiuso`, `Scartato`, `Risolto` = **non attivi** → gli alert vengono generati.

---

## Integrazione Jira

### Creazione Ticket

Dalla tab Alert o Dispositivi:
1. Seleziona righe (Ctrl+Click, Shift+Click per multi-selezione)
2. Clicca **Jira Selezionati**
3. Nel dialog, configura:
   - **Assignee** e **Reporter**: dropdown con rubrica utenti preconfigurata
   - **Watchers**: checkbox per aggiungere osservatori
   - **Labels**: 9 etichette flaggabili (Misure_fuori_range, Porta_aperta, Misure_assenti, Disconnesso, Comandi, Misure_parziali, Allarme_batteria, DigilApp, PiattaformaIoT)
   - **Priority**, **Due Date**, **Data Reply**
4. Modifica Summary/Description nella tabella se necessario
5. **Esporta CSV Jira**

### Rubrica Utenti

| Nome | ID Jira |
|------|---------|
| Festa Rosa | 60705508126db9006f3be9e8 |
| Massimiliano Tavernese | 5e8ae84a84dec20b8159e37a |
| Matteo Massarotto | 712020:166aab81-574f-4eff-8eb0-c2ba527b79bf |
| Vittorio Mitri | 5e86f312b39dbf0c114bdefa |
| Team AMS | 622f434533fb840069656a1a |

### Labels Predefinite

- Misure_fuori_range
- Porta_aperta
- Misure_assenti
- Disconnesso
- Comandi
- Misure_parziali
- Allarme_batteria
- DigilApp
- PiattaformaIoT

### Formato CSV Jira

- Separatore: `;` (punto e virgola)
- Encoding: `utf-8-sig` (BOM per Excel)
- Colonne: Summary, Issue Type, Priority, Assignee, [Reporter], Due date, Labels..., [Watchers...], Description
- Import in Jira: Projects → Import Issues → CSV (separatore `;`)

---

## Interfaccia

### Tab Alert

Alert attivi filtrabili per severity, tipo, fornitore. Filtri inline sotto ogni colonna.

**Toggle "Solo senza ticket"**: bottone arancione per mostrare solo alert su dispositivi senza ticket. Utile per identificare le priorità non ancora gestite.

### Tab Dispositivi

Tutti i ~986 dispositivi. Filtri: fornitore, health, tipo (M/S), installazione, stato ticket (Tutti/Vuoto/Aperto/Chiuso/Scartato/Interno/Risolto).

**Toggle "Solo senza ticket"**: come nel tab Alert.

### Tab Overview

Aggregazioni per fornitore, DT, matrice correlazione. Export Excel.

### Dettaglio Device (doppio click)

- Anagrafica + diagnostiche con semaforo
- **Ultimo availability**: mostra lo stato specifico (COMPLETE/AVAILABLE/NOT AVAILABLE/NO DATA) con barra colorata
- Malfunzionamento, ticket corrente, storico ticket
- **Timeline availability**: griglia colorata con i 4 colori ufficiali + legenda
- Alert recenti
- Bottoni: **Copia Info** (clipboard), **Ticket Jira** (singolo)

### Storico Ticket

Ad ogni import, i ticket vengono tracciati nel DB. Anche dopo che un ticket viene chiuso e sostituito da uno nuovo, lo storico rimane visibile nel dettaglio device con date di prima/ultima osservazione.

---

## Workflow Quotidiano

1. Scarica Excel aggiornato
2. **Importa Excel** → carica dati + genera alert + aggiorna storico ticket
3. Tab **Alert** → filtra CRITICAL/HIGH
4. Attiva **"Solo senza ticket"** per vedere le priorità non gestite
5. Seleziona righe → **Jira Selezionati** → configura rubrica + labels → export CSV
6. Import CSV in Jira (sep `;`)
7. Tab **Overview** per la visione d'insieme
8. Tab **Ticket** per il tracciamento ticket Jira

---

## Integrazione Jira — Tracciamento Ticket

### Sorgente Dati

I ticket vengono scaricati dal progetto **IA20** su Jira Cloud (`terna-it.atlassian.net`). Solo i ticket di tipo **Bug in esercizio** vengono importati.

### Download Automatico

All'avvio del tool, se le credenziali Jira sono configurate come variabili d'ambiente, i ticket vengono scaricati automaticamente:

```bash
# Windows
set JIRA_EMAIL=nome.cognome@reply.it
set JIRA_API_TOKEN=your_api_token

# Linux/Mac
export JIRA_EMAIL=nome.cognome@reply.it
export JIRA_API_TOKEN=your_api_token
```

Il download si ripete automaticamente ogni ora. In alternativa:
- **Aggiorna da Jira**: bottone per forzare il refresh manuale
- **Importa Excel**: carica un file Excel generato da `scaricaTicketJira.py`

### Estrazione DeviceID

Il DeviceID viene estratto automaticamente dal campo Summary del ticket Jira:
- `[Issue_487]: Device 1:1:2:16:25:DIGIL_SR2_0201 Misure parziali` → `1:1:2:16:25:DIGIL_SR2_0201`
- `[Issue_422]: 1:1:2:16:22:DIGIL_MRN_0332 Disconnesso` → `1:1:2:16:22:DIGIL_MRN_0332`

Il fornitore è derivato automaticamente: `DIGIL_IND` → INDRA, `DIGIL_MRN` → MII, `DIGIL_SR2` → SIRTI.

### Correlazione con Excel Monitoring

Quando sia i ticket Jira che il foglio Excel di monitoring sono importati, il tool correla automaticamente:
- **Risoluzione attuata**: dalla colonna del foglio Excel (se assente, cella gialla)
- **Macro-area Causa Problema**: dalla nuova colonna del foglio Excel (se assente, cella gialla)

### Timing (SLA)

Per ogni ticket viene calcolato il tempo trascorso dall'ultimo evento (apertura o aggiornamento), escludendo i weekend:
- 🟢 **Verde**: meno di 24h lavorative
- 🟠 **Arancione**: tra 24h e 48h lavorative
- 🔴 **Rosso**: più di 48h lavorative

### Mappatura Stati Jira → Overview

| Stato Jira | Stato Overview |
|------------|---------------|
| Aperto | Aperto |
| Work In Progress | Aperto |
| Selected For Evaluation | Aperto |
| Chiusa | Chiuso |
| Discarded | Chiuso |
| Suspended | Sospeso |

### Tabella Overview Ticket

Nella tab Overview, la tabella "Ticket Jira per Fornitore e Stato" mostra i ticket aggregati per:
- **Lotto1-IndraOlivetti** (INDRA)
- **Lotto2-TelebitMarini** (MII)
- **Lotto3-Sirti** (SIRTI)
- **Totale complessivo**

Con le 4 colonne: Aperto, Chiuso, Interno, Sospeso.

### Dettaglio Ticket (doppio click)

- Summary, DeviceID, Data Apertura, Reporter, Assignee
- Priority, Resolution, Labels, Fornitore
- **Macro-area Causa Problema** (dal foglio Excel monitoring)
- Descrizione (espandibile se lunga)
- Storico Ticket (Issue Links)
- Commenti
- Bottone **Vedi su Jira** (apre il browser)

---

## Struttura File

```
digil-monitoring-pyqt/
├── main.py           # GUI PyQt5
├── database.py       # Schema SQLAlchemy + TicketHistory
├── importer.py       # ETL: Excel → DB (4 stati availability)
├── detection.py      # 9 regole alert (incluso NO_DATA)
├── jira_client.py    # Client Jira: download API + import Excel + correlazione
├── requirements.txt
├── data/
│   └── digil_monitoring.db
├── assets/
│   └── logo_terna.png (opzionale)
└── README.md
```

---

## Note Tecniche

- SQLite con WAL mode, `data/digil_monitoring.db`
- 4 stati availability: COMPLETE, AVAILABLE, NOT AVAILABLE, NO DATA
- Valori legacy (ON→AVAILABLE, OFF→NOT AVAILABLE) mappati automaticamente
- Alert rigenerati ad ogni import; acknowledged preservati
- Trend su ultimi 7 giorni; giorni = streak consecutiva nello stato attuale
- Storico ticket persistente con first_seen/last_seen
- Ticket Jira scaricati automaticamente all'avvio (se credenziali presenti) e ogni ora
- Colonna "Macro-area Causa Problema" correlata dal foglio Excel monitoring al DeviceID
