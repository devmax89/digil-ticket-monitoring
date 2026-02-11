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

---

## Novità

### Integrazione Jira (Apertura Ticket Singola e Massiva)

Il tool genera CSV pronti per l'import massivo in Jira. Il formato è identico allo script standalone utilizzato precedentemente:

- **Separatore**: `;` (punto e virgola)
- **Issue Type**: `Bug in esercizio`
- **Assignee**: ID Jira preconfigurato (modificabile)
- **Labels**: prese automaticamente da `Tipo Malf Jira` e `Cluster convertito Jira`
- **Description**: formato `Reply DD-MM-YY: Valori recuperati: Check LTE:..., check SSH:..., Batteria:..., Porta aperta:..., Check Mongo:...`

**Come usare:**

1. **Singolo**: Apri il dettaglio di un device (doppio click) → clicca "🎫 Ticket Jira"
2. **Massivo**: Nella tabella Alert o Dispositivi, seleziona più righe (Ctrl+Click o Shift+Click) → clicca "🎫 Jira Selezionati"
3. Si apre un dialog dove **puoi modificare ogni campo** (Summary, Labels, Description) prima dell'export
4. Clicca "Esporta CSV Jira" → il file viene salvato
5. In Jira: Projects → Import Issues → CSV (separatore `;`)

### Cluster NO DATA

Nuova regola di detection che identifica dispositivi con status `NO DATA` / `NOT AVAILABLE` per 5+ giorni consecutivi senza ticket attivo. Diverso da OFF: il device potrebbe trasmettere ma senza dati validi.

- **Severity**: MEDIUM se 5-9 giorni, HIGH se 10+ giorni
- **Colore nella timeline**: indaco (#5C6BC0) per distinguere NO DATA da OFF

### Filtro Ticket

Nel tab Dispositivi è disponibile un filtro dropdown per stato ticket:

- **Tutti**: nessun filtro
- **Vuoto**: dispositivi senza ticket
- **Aperto**: ticket aperto
- **Chiuso**: ticket chiuso
- **Scartato**: ticket scartato
- **Interno**: ticket interno
- **Risolto**: ticket risolto

### Fix KO_NO_TICKET

Corretto il bug per cui dispositivi con ticket già aperti (stato `Aperto` o `Interno`) venivano erroneamente segnalati come "KO senza ticket". Ora la regola verifica:
1. Che `ticket_id` non sia vuoto/nullo
2. Che `ticket_stato` sia `Aperto` o `Interno`

Se il device ha un ticket con stato diverso (es. `Chiuso`, `Scartato`), viene segnalato con nota "(ticket XXX stato: Chiuso)".

### Copia Info Dispositivo

Dal dettaglio dispositivo (doppio click), il bottone "📋 Copia Info" copia negli appunti tutte le informazioni:
- Anagrafica completa (DeviceID, IP, Linea, Fornitore, DT, etc.)
- Diagnostiche (LTE, SSH, Mongo, Batteria, Porta)
- Ticket corrente
- Malfunzionamento
- Storico ticket completo

Tutti i campi di testo nel dettaglio sono selezionabili con il mouse.

### Storico Ticket

Il tool mantiene uno storico di tutti i ticket visti per ogni dispositivo. Ad ogni import:
- Se il ticket è nuovo → viene creato un record nello storico
- Se il ticket esiste già → viene aggiornato (last_seen, stato, risoluzione)

Quando un dispositivo cambia ticket (es. il vecchio viene chiuso e ne viene aperto uno nuovo), entrambi restano nello storico. Nella finestra di dettaglio dispositivo, la sezione "Storico Ticket" mostra tutti i ticket passati e presenti con date e informazioni associate.

---

## Sorgente Dati

Il tool legge due sheet dal file Excel di monitoraggio:

**Sheet "Stato"** (fonte primaria, ~987 righe): contiene l'anagrafica completa di ogni dispositivo, i check diagnostici, le informazioni su malfunzionamenti e ticket, e le colonne storiche di availability.

**Sheet "Av Status"** (~1029 righe): contiene l'availability giornaliera recente con codici numerici (1=COMPLETE, 2=AVAILABLE, 3=NOT AVAILABLE, 4=NO DATA).

**Regola**: Solo i dispositivi presenti nello sheet Stato vengono importati.

---

## Concetti Chiave

### Health (stato di salute)

- **OK**: availability OK e nessuna diagnostica critica
- **KO**: availability KO (dispositivo non operativo)
- **DEGRADED**: availability OK ma almeno una diagnostica in allarme
- **UNKNOWN**: nessun dato di availability

### Trend 7 giorni

Sequenza di ■ (OK) e □ (KO) per gli ultimi 7 giorni.

### Sotto Corona

Installazione senza sensori di tiro. Badge **SC** nella dashboard.

### Fornitori

- **INDRA** (Lotto 1): ~491 dispositivi
- **MII** (Lotto 2): ~319 dispositivi
- **SIRTI** (Lotto 3): ~176 dispositivi

### Diagnostiche

| Check | Significato |
|-------|------------|
| **LTE** | Connettività cellulare |
| **SSH** | Accesso remoto |
| **Mongo** | Invio dati telemetria |
| **Batteria** | Stato batteria |
| **Porta** | Sensore porta quadro |

---

## Sistema di Alert (9 regole)

| # | Regola | Condizione | Severity |
|---|--------|-----------|----------|
| 1 | **KO_NO_TICKET** | KO 3+ gg senza ticket attivo | CRITICAL (7+gg) / HIGH (3-6gg) |
| 2 | **NEW_KO** | Passaggio OK→KO dopo 2+ gg OK | HIGH / MEDIUM |
| 3 | **CONNECTIVITY_LOST** | LTE=KO + SSH=KO senza ticket | HIGH |
| 4 | **DOOR_ALARM** | Porta=KO senza ticket specifico | MEDIUM / LOW |
| 5 | **BATTERY_ALARM** | Batteria=KO | HIGH / MEDIUM |
| 6 | **INTERMITTENT** | 3+ cambi stato in 7 giorni | MEDIUM |
| 7 | **RECOVERED** | Tornato OK dopo 2+ gg KO | LOW |
| 8 | **OPEN_TICKET_OK** | OK 5+ gg ma ticket ancora aperto | LOW |
| 9 | **NO_DATA** | NO DATA/NOT AVAILABLE 5+ gg | HIGH (10+gg) / MEDIUM |

---

## Interfaccia

### Tab Alert

Alert attivi con filtri per severity, tipo, fornitore. Multi-selezione per generazione ticket Jira massiva.

### Tab Dispositivi

Tutti i dispositivi con filtri per fornitore, health, tipo, installazione, **stato ticket**.

### Tab Overview

Vista aggregata per fornitore e DT con export Excel.

### Dettaglio Device (doppio click)

Anagrafica, diagnostiche, malfunzionamento, ticket corrente, **storico ticket**, timeline availability (con colore NO DATA indaco), alert recenti. Bottoni per copiare info e creare ticket Jira.

### Timeline Availability - Legenda Colori

| Colore | Significato |
|--------|------------|
| 🟢 Verde | OK (ON / COMPLETE / AVAILABLE) |
| 🔴 Rosso | OFF (KO) |
| 🟡 Giallo | NOT AVAILABLE |
| 🟣 Indaco | NO DATA |

---

## Workflow Quotidiano

1. Scarica il file Excel aggiornato
2. Apri il tool e clicca **Importa Excel**
3. Vai nel tab **Alert** per le situazioni critiche
4. Filtra per CRITICAL/HIGH
5. Seleziona alert senza ticket → "🎫 Jira Selezionati" → modifica se necessario → esporta CSV
6. Importa CSV in Jira
7. Tab **Dispositivi** con filtro Ticket=Vuoto per vedere chi non ha ticket
8. Doppio click per dettaglio con storico ticket

---

## Struttura File

```
digil-monitoring-pyqt/
├── main.py           # GUI PyQt5 (dashboard + dialog Jira)
├── database.py       # Modelli SQLAlchemy + TicketHistory
├── importer.py       # ETL: Excel → Database + storico ticket
├── detection.py      # 9 regole di detection (+ NO_DATA)
├── requirements.txt  # Dipendenze Python
├── data/
│   └── digil_monitoring.db  # Database SQLite
└── README.md
```

---

## Note Tecniche

- SQLite con WAL mode per performance
- Storico ticket persistente: i ticket chiusi non vengono mai cancellati
- CSV Jira: encoding UTF-8 BOM, separatore `;`, compatibile con import bulk Jira
- Multi-selezione nelle tabelle (Ctrl+Click, Shift+Click) per operazioni massive
- Al primo import dopo l'aggiornamento, cancellare `data/digil_monitoring.db` per ricreare lo schema con la nuova tabella TicketHistory
