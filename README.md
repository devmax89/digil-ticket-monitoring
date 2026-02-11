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

## Sorgente Dati

Il tool legge due sheet dal file Excel di monitoraggio:

**Sheet "Stato"** (fonte primaria, ~987 righe): contiene l'anagrafica completa di ogni dispositivo (linea, sostegno, fornitore, DT, IP, tipo installazione), i check diagnostici (LTE, SSH, MongoDB, batteria, porta), le informazioni su malfunzionamenti e ticket, e le colonne storiche di availability (formato "AVAILABILITY DD mmm", es. "AVAILABILITY 17 dic").

**Sheet "Av Status"** (~1029 righe): contiene l'availability giornaliera recente con codici numerici (1=COMPLETE, 2=AVAILABLE, 3=NOT AVAILABLE, 4=NO DATA) per gli ultimi 5 giorni circa.

**Regola importante**: Solo i dispositivi presenti nello sheet Stato vengono importati. Lo sheet Av Status contiene circa 45 dispositivi in più che corrispondono ad apparati non collaudati; questi vengono automaticamente ignorati.

---

## Concetti Chiave

### Health (stato di salute)

Ogni dispositivo ha uno stato di salute calcolato:

- **OK**: l'ultimo dato di availability è OK e non ci sono diagnostiche critiche
- **KO**: l'ultimo dato di availability è KO (dispositivo non operativo)
- **DEGRADED**: l'availability è OK ma almeno una diagnostica è in allarme (es. LTE KO o porta aperta). Il dispositivo funziona ma con problemi
- **UNKNOWN**: nessun dato di availability disponibile

### Trend 7 giorni

Il trend mostra lo storico availability degli ultimi 7 giorni come sequenza di quadrati pieni (OK) e vuoti (KO). Esempio: `■■■□□□■` significa OK per 3 giorni, poi KO per 3 giorni, poi di nuovo OK.

Serve per capire a colpo d'occhio se un dispositivo è stabilmente OK, stabilmente KO, o intermittente (cambia spesso stato).

### Giorni

Indica da quanti giorni consecutivi il dispositivo è nello stato attuale. Se un dispositivo è KO da 14 giorni, vedrai "14" nella colonna Giorni. Questo numero è critico per capire la gravità: un KO da 1 giorno è diverso da un KO da 14 giorni senza ticket.

### Sotto Corona

Un'installazione "sotto corona" (flag "SC" nella dashboard) indica un traliccio dove non sono installati i sensori di tiro sui conduttori e funi di guardia. Questi dispositivi non riceveranno mai metriche di tiro, quindi non devono generare alert relativi a quei sensori.

L'informazione viene letta dalla colonna "Tipo Installazione AM" del file Excel:
- "Inst. Completa" → installazione completa con tutti i sensori
- "Inst. Sotto corona" o "Sotto Corona" → senza sensori di tiro

Nella dashboard, i dispositivi sotto corona sono identificati dal badge **SC** visibile nelle tabelle Alert e Dispositivi, e nel dettaglio device.

### Fornitori

I dispositivi sono gestiti da 3 fornitori (lotti):
- **INDRA** (Lotto 1 — IndraOlivetti): ~491 dispositivi
- **MII** (Lotto 2 — TelebitMarini): ~319 dispositivi
- **SIRTI** (Lotto 3): ~176 dispositivi

### Diagnostiche

Per ogni dispositivo vengono monitorati 5 check diagnostici:

| Check | Significato |
|-------|------------|
| **LTE** | Connettività cellulare. KO = il dispositivo non è raggiungibile via LTE |
| **SSH** | Accesso remoto via SSH. KO = impossibile connettersi al dispositivo |
| **Mongo** | Invio dati a MongoDB. KO = il dispositivo non sta trasmettendo telemetria |
| **Batteria** | Stato batteria. KO = batteria in allarme, rischio disconnessione |
| **Porta** | Sensore porta. KO = porta del quadro aperta (possibile intrusione o guasto) |

Quando sia LTE che SSH sono KO, il dispositivo è considerato **disconnesso** — non è raggiungibile in nessun modo.

---

## Sistema di Alert

Il tool genera automaticamente alert analizzando la combinazione di availability, diagnostiche e ticket. Ogni alert ha una **severity** (gravità):

| Severity | Significato | Colore |
|----------|------------|--------|
| **CRITICAL** | Richiede intervento immediato | Rosso |
| **HIGH** | Problema serio, va gestito presto | Arancione |
| **MEDIUM** | Situazione da monitorare | Giallo |
| **LOW** | Informativo, azione suggerita | Verde |

### Le 8 regole di detection

**1. KO_NO_TICKET** — Device KO senza ticket
- Condizione: KO da 3+ giorni consecutivi E nessun ticket aperto
- Severity: CRITICAL se KO da 7+ giorni, HIGH se 3-6 giorni
- Azione suggerita: aprire un ticket al fornitore

**2. NEW_KO** — Nuovo passaggio a KO
- Condizione: era OK da 2+ giorni E oggi è KO
- Severity: HIGH se senza ticket, MEDIUM se ticket già aperto
- Significato: qualcosa si è rotto di recente

**3. CONNECTIVITY_LOST** — Disconnessione completa
- Condizione: LTE=KO E SSH=KO E nessun ticket aperto
- Severity: HIGH
- Significato: il dispositivo è irraggiungibile, potrebbe essere un problema di alimentazione o SIM

**4. DOOR_ALARM** — Porta aperta
- Condizione: Porta=KO E nessun ticket specifico per porta aperta
- Severity: MEDIUM se nessun ticket, LOW se c'è un ticket per altro motivo
- Significato: il quadro sul traliccio è aperto

**5. BATTERY_ALARM** — Batteria in allarme
- Condizione: Batteria=KO
- Severity: HIGH se senza ticket, MEDIUM se ticket aperto
- Significato: rischio imminente di disconnessione

**6. INTERMITTENT** — Comportamento intermittente
- Condizione: 3+ cambi di stato (OK↔KO) negli ultimi 7 giorni
- Severity: MEDIUM
- Significato: il dispositivo è instabile, potrebbe indicare problemi hardware o di connettività

**7. RECOVERED** — Ripresa dopo guasto
- Condizione: era KO da 2+ giorni E oggi è OK
- Severity: LOW
- Azione suggerita: se c'è un ticket aperto, verificare se chiudibile

**8. OPEN_TICKET_OK** — Ticket aperto ma device OK
- Condizione: OK da 5+ giorni consecutivi E ticket ancora aperto
- Severity: LOW
- Azione suggerita: verificare se il ticket può essere chiuso

---

## Interfaccia

### Tab Alert

Mostra tutti gli alert attivi (non ancora confermati), ordinabili e filtrabili per severity, tipo, fornitore, DT. Ogni riga mostra il dispositivo con tutti i check diagnostici, il trend, e la descrizione del problema.

Filtri disponibili:
- **Dropdown in alto**: filtrano per severity, tipo alert, fornitore, DT
- **Filtri inline sotto ogni colonna**: digitare per filtrare (es. digitare "SIRTI" nella colonna Fornitore)
- **Bottone "Pulisci Filtri"**: resetta tutti i filtri
- **Doppio click** su una riga: apre il dettaglio completo del dispositivo

### Tab Dispositivi

Mostra tutti i 986 dispositivi con il loro stato corrente. Filtri per fornitore, health, tipo (master/slave), tipo installazione (completa/sotto corona).

### Tab Overview

Vista aggregata con:
- Stato per Fornitore (totale, OK, KO, degraded, % OK, ticket, sotto corona)
- Stato per Direzione Territoriale
- Matrice di correlazione diagnostica (quanti LTE KO, SSH KO, etc. per fornitore)

### Dettaglio Device (doppio click)

Finestra modale con:
- Anagrafica completa (linea, sostegno, IP, tipo installazione, fornitore, DT)
- 5 indicatori diagnostici con semaforo colorato
- Trend e giorni nello stato attuale
- Informazioni malfunzionamento (tipo, cluster, analisi, strategia)
- Ticket associato (ID, stato, date)
- Timeline availability (griglia colorata con tutti i giorni disponibili)
- Alert recenti sul dispositivo

---

## Workflow Quotidiano

1. Scarica il file Excel aggiornato di monitoraggio
2. Apri il tool e clicca **Importa Excel**
3. Seleziona il file → il tool importa i dati e genera gli alert
4. Vai nel tab **Alert** per vedere le situazioni critiche
5. Filtra per CRITICAL/HIGH per le priorità
6. Doppio click per il dettaglio di ogni dispositivo problematico
7. Usa il tab **Overview** per la visione d'insieme

Il database SQLite viene aggiornato ad ogni import. Non è necessario cancellare nulla tra un import e l'altro — i dati vengono sovrascritti/aggiornati.

---

## Struttura File

```
digil-monitoring-pyqt/
├── main.py           # GUI PyQt5 (dashboard principale)
├── database.py       # Modelli SQLAlchemy + schema DB
├── importer.py       # ETL: Excel → Database
├── detection.py      # Motore di generazione alert (8 regole)
├── requirements.txt  # Dipendenze Python
├── data/
│   └── digil_monitoring.db  # Database SQLite (creato all'import)
└── README.md         # Questa documentazione
```

---

## Note Tecniche

- Il tool usa SQLite come database locale (file `data/digil_monitoring.db`)
- Le date di availability dallo sheet Stato sono nel formato "AVAILABILITY DD mmm" (dicembre=2025, gen-feb=2026)
- Le date dallo sheet Av Status sono oggetti datetime Python nelle colonne
- I codici numerici di Av Status: 1=COMPLETE(OK), 2=AVAILABLE(OK), 3=NOT AVAILABLE(KO), 4=NO DATA(KO)
- Gli alert vengono rigenerati ad ogni import; quelli confermati (acknowledged) vengono preservati
- Il trend è calcolato sugli ultimi 7 giorni di availability disponibili
