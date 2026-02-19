"""
DIGIL Monitoring - Jira Client
Scarica ticket dal progetto IA20 e li salva nel DB locale.
"""
import re, os, json
from datetime import datetime, date, timedelta
from pathlib import Path
from database import get_session, init_db, Device

try:
    from jira import JIRA
    HAS_JIRA = True
except ImportError:
    HAS_JIRA = False

from sqlalchemy import create_engine, Column, String, Integer, Boolean, Date, DateTime, Text, Index
from sqlalchemy.orm import declarative_base, sessionmaker
from database import Base, engine, SessionLocal

BASE_DIR = Path(__file__).parent
ENV_FILE = BASE_DIR / ".env"


def _load_credentials():
    """Carica credenziali Jira da .env file o variabili d'ambiente.
    Il file .env deve essere nella stessa cartella del tool, formato:
        JIRA_EMAIL=nome.cognome@reply.it
        JIRA_API_TOKEN=your_token
    """
    email = os.environ.get("JIRA_EMAIL", "")
    token = os.environ.get("JIRA_API_TOKEN", "")
    # Prova a leggere da .env file
    if ENV_FILE.exists():
        try:
            for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    k, v = line.split("=", 1)
                    k = k.strip(); v = v.strip()
                    if k == "JIRA_EMAIL" and v:
                        email = v
                    elif k == "JIRA_API_TOKEN" and v:
                        token = v
        except Exception:
            pass
    return email, token

# ============================================================
# MODELLO DB PER I TICKET JIRA
# ============================================================
class JiraTicket(Base):
    __tablename__ = "jira_tickets"
    key = Column(String, primary_key=True)          # IA20-123
    summary = Column(Text)
    description = Column(Text)
    issue_type = Column(String)
    status = Column(String)
    resolution = Column(String)
    priority = Column(String)
    assignee = Column(String)
    reporter = Column(String)
    created = Column(DateTime)
    updated = Column(DateTime)
    due_date = Column(Date)
    labels = Column(String)
    comments = Column(Text)
    num_comments = Column(Integer, default=0)
    issue_links = Column(Text)
    url = Column(String)
    # Nuovi campi custom Jira
    assignee_level = Column(String)                  # L3, L4, etc.
    vendor = Column(String)
    info_l1 = Column(Text)
    info_l2 = Column(Text)
    info_l3 = Column(Text)
    info_l4 = Column(Text)
    status_l1 = Column(String)
    status_l2 = Column(String)
    status_l3 = Column(String)
    status_l4 = Column(String)
    # Campi derivati
    device_id = Column(String)                       # Estratto da Summary
    fornitore = Column(String)                       # INDRA/MII/SIRTI
    # Campi da Excel monitoring (correlazione)
    risoluzione_attuata = Column(Text)
    macro_area = Column(String)
    last_synced = Column(DateTime, default=datetime.utcnow)
    __table_args__ = (
        Index("idx_jt_status", "status"),
        Index("idx_jt_device", "device_id"),
        Index("idx_jt_level", "assignee_level"),
    )


# Crea tabella se non esiste
def init_jira_db():
    Base.metadata.create_all(engine)


def extract_device_id(summary: str) -> str:
    """Estrae il DeviceID dal Summary Jira.
    Patterns:
      'Device 1:1:2:16:25:DIGIL_SR2_0201 ...'
      '[Issue_422]: 1:1:2:16:22:DIGIL_MRN_0332 Disconnesso'
      'Problema Invio dati in Piattaforma 1:1:2:15:25:DIGIL_SR2_0109'
    """
    if not summary:
        return ""
    # Pattern 1: "Device XXXXX"
    m = re.search(r'Device\s+([\d:]+:DIGIL_\w+)', str(summary))
    if m:
        return m.group(1)
    # Pattern 2: DeviceID diretto nel testo
    m = re.search(r'(\d+:\d+:\d+:\d+:\d+:DIGIL_\w+)', str(summary))
    if m:
        return m.group(1)
    return ""


def extract_fornitore(device_id: str) -> str:
    """Ricava il fornitore dal DeviceID."""
    if not device_id:
        return ""
    did_upper = device_id.upper()
    if "DIGIL_IND" in did_upper:
        return "INDRA"
    elif "DIGIL_MRN" in did_upper:
        return "MII"
    elif "DIGIL_SR2" in did_upper or "DIGIL_SRT" in did_upper:
        return "SIRTI"
    return ""

# Mappatura stati Jira -> 4 stati overview
JIRA_STATUS_MAP = {
    "Aperto": "Aperto",
    "Work In Progress": "Aperto",
    "Selected For Evaluation": "Aperto",
    "Chiusa": "Chiuso",
    "Discarded": "Chiuso",
    "Suspended": "Sospeso",
}

FORNITORE_DISPLAY = {
    "INDRA": "Lotto1-IndraOlivetti",
    "MII": "Lotto2-TelebitMarini",
    "SIRTI": "Lotto3-Sirti",
}

def map_jira_status(raw_status: str) -> str:
    return JIRA_STATUS_MAP.get(raw_status, raw_status)


def compute_timing_hours(created_str, updated_str):
    """Calcola ore lavorative (escl weekend) dall'ultimo evento.
    Ritorna (ore, colore): verde <24h, arancione <48h, rosso >=48h
    """
    try:
        now = datetime.now()
        # Usa updated se presente e diverso da created, altrimenti created
        if updated_str and str(updated_str) != str(created_str):
            ref = updated_str if isinstance(updated_str, datetime) else datetime.fromisoformat(str(updated_str)[:19])
        else:
            ref = created_str if isinstance(created_str, datetime) else datetime.fromisoformat(str(created_str)[:19])

        # Conta ore lavorative (lun-ven, 8h/giorno)
        biz_hours = 0
        current = ref
        while current < now:
            if current.weekday() < 5:  # Lun-Ven
                biz_hours += 1
            current += timedelta(hours=1)

        if biz_hours < 24:
            return biz_hours, "GREEN"
        elif biz_hours < 48:
            return biz_hours, "ORANGE"
        else:
            return biz_hours, "RED"
    except Exception:
        return 0, "GREEN"


# ============================================================
# CUSTOM FIELDS HELPER
# ============================================================
# Nomi dei campi custom da cercare su Jira
CUSTOM_FIELD_NAMES = [
    "Assignee Level", "Vendor",
    "Info L1", "Info L2", "Info L3", "Info L4",
    "Status L1", "Status L2", "Status L3", "Status L4",
]

_custom_field_cache = {}

def _discover_custom_fields(jira) -> dict:
    """Scopre automaticamente gli ID dei custom fields da Jira.
    Ritorna un dict {nome_campo: customfield_XXXXX}.
    Risultati cachati per la sessione."""
    global _custom_field_cache
    if _custom_field_cache:
        return _custom_field_cache
    try:
        all_fields = jira.fields()
        name_to_id = {}
        for field in all_fields:
            name = field.get("name", "")
            fid = field.get("id", "")
            if name in CUSTOM_FIELD_NAMES:
                name_to_id[name] = fid
        _custom_field_cache = name_to_id
        if name_to_id:
            print(f"[Jira] Custom fields trovati: {', '.join(f'{k}={v}' for k,v in name_to_id.items())}")
        return name_to_id
    except Exception as e:
        print(f"[Jira] Errore discovery custom fields: {e}")
        return {}


def _get_custom_field(fields, field_map: dict, field_name: str) -> str:
    """Estrae il valore di un custom field dall'issue.
    Gestisce sia valori stringa che oggetti con .value o .name."""
    fid = field_map.get(field_name)
    if not fid:
        return ""
    try:
        val = getattr(fields, fid, None)
        if val is None:
            return ""
        # Jira custom fields possono essere: stringa, dict con value, oggetto con .value/.name
        if isinstance(val, str):
            return val.strip()
        if isinstance(val, dict):
            return str(val.get("value", val.get("name", ""))).strip()
        if hasattr(val, 'value'):
            return str(val.value).strip()
        if hasattr(val, 'name'):
            return str(val.name).strip()
        return str(val).strip()
    except Exception:
        return ""


# ============================================================
# DOWNLOAD DA JIRA API
# ============================================================
def download_from_jira(email=None, token=None, jira_url="https://terna-it.atlassian.net", project="IA20"):
    """Scarica tutti i ticket Bug in esercizio da Jira e li salva nel DB."""
    if not HAS_JIRA:
        return False, "Libreria 'jira' non installata. Esegui: pip install jira"

    if not email or not token:
        email, token = _load_credentials()
    if not email or not token:
        return False, "Credenziali Jira mancanti — crea file .env nella cartella del tool"

    try:
        jira = JIRA(server=jira_url, basic_auth=(email, token))
    except Exception as e:
        return False, f"Connessione fallita: {e}"

    # Auto-discover custom field IDs
    custom_field_map = _discover_custom_fields(jira)

    jql = f'project = {project} AND type = "Bug in esercizio" ORDER BY created DESC'
    try:
        issues = list(jira.search_issues(jql, maxResults=False))
    except Exception as e:
        return False, f"Query fallita: {e}"

    init_jira_db()
    session = SessionLocal()
    count = 0
    try:
        for issue in issues:
            f = issue.fields
            key = issue.key
            summary = f.summary or ""
            device_id = extract_device_id(summary)
            fornitore = extract_fornitore(device_id)

            # Comments
            comments_str = ""
            try:
                comments = jira.comments(key)
                if comments:
                    parts = []
                    for c in comments:
                        parts.append(f"[{c.created[:19]}] {c.author.displayName}:\n{c.body}")
                    comments_str = "\n---\n".join(parts)
            except Exception:
                pass

            # Issue Links
            links_str = ""
            try:
                if hasattr(f, 'issuelinks') and f.issuelinks:
                    links = []
                    for link in f.issuelinks:
                        if hasattr(link, 'outwardIssue'):
                            links.append(f"{link.type.outward}: {link.outwardIssue.key}")
                        elif hasattr(link, 'inwardIssue'):
                            links.append(f"{link.type.inward}: {link.inwardIssue.key}")
                    links_str = ", ".join(links)
            except Exception:
                pass

            ticket = session.get(JiraTicket, key)
            if ticket is None:
                ticket = JiraTicket(key=key)
                session.add(ticket)

            ticket.summary = summary
            ticket.description = f.description or ""
            ticket.issue_type = f.issuetype.name if f.issuetype else ""
            ticket.status = f.status.name if f.status else ""
            ticket.resolution = f.resolution.name if f.resolution else "Unresolved"
            ticket.priority = f.priority.name if f.priority else ""
            ticket.assignee = f.assignee.displayName if f.assignee else "Unassigned"
            ticket.reporter = f.reporter.displayName if f.reporter else ""
            ticket.labels = ", ".join(f.labels) if f.labels else ""
            ticket.url = f"{jira_url}/browse/{key}"
            ticket.num_comments = len(jira.comments(key)) if comments_str else 0
            ticket.comments = comments_str or "Nessun commento"
            ticket.issue_links = links_str or "Nessun link"
            ticket.device_id = device_id
            ticket.fornitore = fornitore

            # Custom fields
            ticket.assignee_level = _get_custom_field(f, custom_field_map, "Assignee Level")
            ticket.vendor = _get_custom_field(f, custom_field_map, "Vendor")
            ticket.info_l1 = _get_custom_field(f, custom_field_map, "Info L1")
            ticket.info_l2 = _get_custom_field(f, custom_field_map, "Info L2")
            ticket.info_l3 = _get_custom_field(f, custom_field_map, "Info L3")
            ticket.info_l4 = _get_custom_field(f, custom_field_map, "Info L4")
            ticket.status_l1 = _get_custom_field(f, custom_field_map, "Status L1")
            ticket.status_l2 = _get_custom_field(f, custom_field_map, "Status L2")
            ticket.status_l3 = _get_custom_field(f, custom_field_map, "Status L3")
            ticket.status_l4 = _get_custom_field(f, custom_field_map, "Status L4")

            try:
                ticket.created = datetime.fromisoformat(f.created[:19]) if f.created else None
            except Exception:
                pass
            try:
                ticket.updated = datetime.fromisoformat(f.updated[:19]) if f.updated else None
            except Exception:
                pass
            try:
                ticket.due_date = date.fromisoformat(str(f.duedate)) if f.duedate else None
            except Exception:
                pass

            ticket.last_synced = datetime.utcnow()
            count += 1

        # Correlazione con dati Excel (Device table)
        _correlate_with_devices(session)

        session.commit()
        return True, f"{count} ticket scaricati"
    except Exception as e:
        session.rollback()
        return False, f"Errore salvataggio: {e}"
    finally:
        session.close()


def import_from_excel(file_path: str):
    """Importa ticket da un file Excel o CSV esportato da Jira."""
    import pandas as pd
    fp = str(file_path)
    if fp.lower().endswith('.csv'):
        # CSV export diretto da Jira (colonne in italiano)
        df = pd.read_csv(fp, encoding='utf-8')
        # Mappa colonne italiane → inglesi
        col_map = {
            "Riepilogo": "Summary", "Chiave di ticket": "Key", "Tipo ticket": "Type",
            "Stato": "Status", "Priorità": "Priority", "PrioritÃ\xa0": "Priority",
            "Risoluzione": "Resolution",
            "Assegnatario": "Assignee", "Richiedente": "Reporter",
            "Creati": "Created", "Aggiornato": "Updated", "Data di scadenza": "Due Date",
            "Descrizione": "Description", "Etichette": "Labels",
        }
        df.rename(columns={k: v for k, v in col_map.items() if k in df.columns}, inplace=True)
        # Tipo ticket: filtra solo Bug in esercizio
        if "Type" in df.columns:
            df = df[df["Type"] == "Bug in esercizio"].copy()
    else:
        try:
            df = pd.read_excel(fp, sheet_name="All Tickets")
        except Exception:
            df = pd.read_excel(fp, sheet_name=0)
        if "Type" in df.columns:
            df = df[df["Type"] == "Bug in esercizio"].copy()

    init_jira_db()
    session = SessionLocal()
    count = 0
    try:
        for _, row in df.iterrows():
            key = str(row.get("Key", ""))
            if not key:
                continue
            summary = str(row.get("Summary", ""))
            device_id = extract_device_id(summary)
            fornitore = extract_fornitore(device_id)

            ticket = session.get(JiraTicket, key)
            if ticket is None:
                ticket = JiraTicket(key=key)
                session.add(ticket)

            ticket.summary = summary
            ticket.description = str(row.get("Description", "")) if not pd.isna(row.get("Description")) else ""
            ticket.issue_type = str(row.get("Type", ""))
            ticket.status = str(row.get("Status", ""))
            ticket.resolution = str(row.get("Resolution", "Unresolved"))
            ticket.priority = str(row.get("Priority", ""))
            ticket.assignee = str(row.get("Assignee", "")) if not pd.isna(row.get("Assignee")) else "Unassigned"
            ticket.reporter = str(row.get("Reporter", "")) if not pd.isna(row.get("Reporter")) else ""
            ticket.labels = str(row.get("Labels", "")) if not pd.isna(row.get("Labels")) else ""
            ticket.url = str(row.get("URL", "")) if not pd.isna(row.get("URL")) else ""
            ticket.comments = str(row.get("Comments", "")) if not pd.isna(row.get("Comments")) else ""
            ticket.num_comments = int(row.get("Num Comments", 0)) if not pd.isna(row.get("Num Comments")) else 0
            ticket.issue_links = str(row.get("Issue Links", "")) if not pd.isna(row.get("Issue Links")) else ""
            ticket.device_id = device_id
            ticket.fornitore = fornitore

            # Custom fields (se presenti nel file Excel)
            # Supporta sia nomi diretti che formato CSV Jira "Campo personalizzato (Nome)"
            def _safe_excel(col, alt_col=None):
                v = row.get(col)
                if (v is None or (isinstance(v, float) and pd.isna(v))) and alt_col:
                    v = row.get(alt_col)
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    return ""
                return str(v).strip()
            ticket.assignee_level = _safe_excel("Assignee Level", "Campo personalizzato (Assignee Level)")
            ticket.vendor = _safe_excel("Vendor", "Campo personalizzato (Vendor)")
            ticket.info_l1 = _safe_excel("Info L1", "Campo personalizzato (Info L1)")
            ticket.info_l2 = _safe_excel("Info L2", "Campo personalizzato (Info L2)")
            ticket.info_l3 = _safe_excel("Info L3", "Campo personalizzato (Info L3)")
            ticket.info_l4 = _safe_excel("Info L4", "Campo personalizzato (Info L4)")
            ticket.status_l1 = _safe_excel("Status L1", "Campo personalizzato (Status L1)")
            ticket.status_l2 = _safe_excel("Status L2", "Campo personalizzato (Status L2)")
            ticket.status_l3 = _safe_excel("Status L3", "Campo personalizzato (Status L3)")
            ticket.status_l4 = _safe_excel("Status L4", "Campo personalizzato (Status L4)")

            try:
                c = str(row.get("Created", ""))
                ticket.created = datetime.fromisoformat(c[:19]) if c and c != "nan" else None
            except Exception:
                pass
            try:
                u = str(row.get("Updated", ""))
                ticket.updated = datetime.fromisoformat(u[:19]) if u and u != "nan" else None
            except Exception:
                pass
            try:
                d = str(row.get("Due Date", ""))
                ticket.due_date = date.fromisoformat(d[:10]) if d and d != "nan" and d != "NaT" else None
            except Exception:
                pass

            ticket.last_synced = datetime.utcnow()
            count += 1

        _correlate_with_devices(session)
        session.commit()
        return True, f"{count} ticket importati da Excel"
    except Exception as e:
        session.rollback()
        return False, f"Errore import: {e}"
    finally:
        session.close()


def _correlate_with_devices(session):
    """Correla i ticket Jira con i dati dei Device (Risoluzione attuata, Macro-area)."""
    tickets = session.query(JiraTicket).filter(JiraTicket.device_id != "", JiraTicket.device_id != None).all()
    for ticket in tickets:
        device = session.get(Device, ticket.device_id)
        if device:
            ticket.risoluzione_attuata = device.risoluzione_attuata or ""
            ticket.macro_area = getattr(device, 'macro_area_causa', None) or ""


def get_ticket_data(filters=None):
    """Ritorna i ticket filtrati per la visualizzazione."""
    session = SessionLocal()
    try:
        q = session.query(JiraTicket)
        if filters:
            if filters.get("status"):
                q = q.filter(JiraTicket.status == filters["status"])
            if filters.get("reporter"):
                q = q.filter(JiraTicket.reporter == filters["reporter"])
            if filters.get("assignee"):
                q = q.filter(JiraTicket.assignee == filters["assignee"])
            if filters.get("resolution"):
                q = q.filter(JiraTicket.resolution == filters["resolution"])
            if filters.get("priority"):
                q = q.filter(JiraTicket.priority == filters["priority"])
            if filters.get("created_from"):
                q = q.filter(JiraTicket.created >= filters["created_from"])
            if filters.get("created_to"):
                q = q.filter(JiraTicket.created <= filters["created_to"])

        tickets = q.order_by(JiraTicket.created.desc()).all()
        result = []
        for t in tickets:
            hours, color = compute_timing_hours(t.created, t.updated)
            result.append({
                "key": t.key,
                "device_id": t.device_id or "",
                "created": t.created,
                "status": t.status or "",
                "labels": t.labels or "",
                "risoluzione": t.risoluzione_attuata or "",
                "updated": t.updated,
                "due_date": t.due_date,
                "timing_hours": hours,
                "timing_color": color,
                "summary": t.summary or "",
                "description": t.description or "",
                "reporter": t.reporter or "",
                "assignee": t.assignee or "",
                "assignee_level": t.assignee_level or "",
                "resolution": t.resolution or "",
                "priority": t.priority or "",
                "macro_area": t.macro_area or "",
                "comments": t.comments or "",
                "issue_links": t.issue_links or "",
                "url": t.url or "",
                "fornitore": t.fornitore or "",
                "num_comments": t.num_comments or 0,
            })
        return result
    finally:
        session.close()


def get_ticket_overview_by_fornitore():
    """Ritorna dati per tabella overview: ticket per fornitore con Aperto L3/L4."""
    session = SessionLocal()
    try:
        tickets = session.query(JiraTicket).all()
        target_stati = ["Aperto L3", "Aperto L4", "Chiuso", "Sospeso"]
        fornitore_order = ["INDRA", "MII", "SIRTI"]
        # Stati Jira considerati "aperti" (non chiusi/scartati)
        stati_aperti = {"Aperto", "Work In Progress", "Selected For Evaluation"}

        data = {f: {s: 0 for s in target_stati} for f in fornitore_order}
        for t in tickets:
            forn = t.fornitore if t.fornitore in fornitore_order else None
            if not forn:
                continue
            mapped = map_jira_status(t.status or "")
            if mapped == "Aperto":
                # Split per Assignee Level
                level = (t.assignee_level or "").strip().upper()
                if level == "L4":
                    data[forn]["Aperto L4"] += 1
                else:
                    # Default L3 se livello non specificato o L3
                    data[forn]["Aperto L3"] += 1
            elif mapped == "Chiuso":
                data[forn]["Chiuso"] += 1
            elif mapped == "Sospeso":
                data[forn]["Sospeso"] += 1

        return data, target_stati
    finally:
        session.close()


def get_filter_options():
    """Ritorna le opzioni uniche per i filtri."""
    session = SessionLocal()
    try:
        reporters = sorted(set(r[0] for r in session.query(JiraTicket.reporter).distinct().all() if r[0]))
        statuses = sorted(set(r[0] for r in session.query(JiraTicket.status).distinct().all() if r[0]))
        assignees = sorted(set(r[0] for r in session.query(JiraTicket.assignee).distinct().all() if r[0]))
        resolutions = sorted(set(r[0] for r in session.query(JiraTicket.resolution).distinct().all() if r[0]))
        priorities = sorted(set(r[0] for r in session.query(JiraTicket.priority).distinct().all() if r[0]))
        return {"reporters": reporters, "statuses": statuses, "assignees": assignees,
                "resolutions": resolutions, "priorities": priorities}
    finally:
        session.close()


def get_jira_stats():
    """Ritorna statistiche Jira per le cards.
    Card 1: Aperto L3, Aperto L4, Chiuso, Sospeso (totali)
    Card 2: Settimanale aperti, chiusi, scartati
    """
    session = SessionLocal()
    try:
        total = session.query(JiraTicket).count()
        empty = {"total": 0, "aperto_l3": 0, "aperto_l4": 0, "chiuso": 0, "sospeso": 0,
                 "week_aperti": 0, "week_chiusi": 0, "week_scartati": 0}
        if total == 0:
            return empty

        now = datetime.now()
        stati_chiusi = ["Chiusa", "Discarded"]
        stati_sospesi = ["Suspended"]

        # Totali: Aperto L3, Aperto L4, Chiuso, Sospeso
        all_tickets = session.query(JiraTicket).all()
        aperto_l3 = 0; aperto_l4 = 0; chiuso = 0; sospeso = 0
        for t in all_tickets:
            mapped = map_jira_status(t.status or "")
            if mapped == "Aperto":
                level = (t.assignee_level or "").strip().upper()
                if level == "L4":
                    aperto_l4 += 1
                else:
                    aperto_l3 += 1
            elif mapped == "Chiuso":
                chiuso += 1
            elif mapped == "Sospeso":
                sospeso += 1

        # Settimanale (ultimi 7 giorni)
        d7 = now - timedelta(days=7)
        week_aperti = session.query(JiraTicket).filter(
            ~JiraTicket.status.in_(stati_chiusi),
            JiraTicket.created >= d7).count()
        week_chiusi = session.query(JiraTicket).filter(
            JiraTicket.status == "Chiusa",
            JiraTicket.updated >= d7).count()
        week_scartati = session.query(JiraTicket).filter(
            JiraTicket.status == "Discarded",
            JiraTicket.updated >= d7).count()

        # Ultimi 30 giorni
        d30 = now - timedelta(days=30)
        month_aperti = session.query(JiraTicket).filter(
            ~JiraTicket.status.in_(stati_chiusi),
            JiraTicket.created >= d30).count()
        month_chiusi = session.query(JiraTicket).filter(
            JiraTicket.status == "Chiusa",
            JiraTicket.updated >= d30).count()
        month_scartati = session.query(JiraTicket).filter(
            JiraTicket.status == "Discarded",
            JiraTicket.updated >= d30).count()

        return {"total": total, "aperto_l3": aperto_l3, "aperto_l4": aperto_l4,
                "chiuso": chiuso, "sospeso": sospeso,
                "week_aperti": week_aperti, "week_chiusi": week_chiusi, "week_scartati": week_scartati,
                "month_aperti": month_aperti, "month_chiusi": month_chiusi, "month_scartati": month_scartati}
    finally:
        session.close()
