"""
DIGIL Monitoring - Excel Importer
Aggiunto: storico ticket (TicketHistory) aggiornato ad ogni import.
"""
import pandas as pd
import numpy as np
import json, re
from datetime import datetime, date
from typing import Dict, Tuple, Optional
from pathlib import Path
from database import get_session, init_db, Device, AvailabilityDaily, AnomalyEvent, ImportLog, TicketHistory

AV_STATUS_NUMERIC = {1: "COMPLETE", 2: "AVAILABLE", 3: "NOT AVAILABLE", 4: "NO DATA"}

FORNITORE_MAP = {
    "Lotto1-IndraOlivetti": "INDRA", "Lotto2-TelebitMarini": "MII", "Lotto3-Sirti": "SIRTI",
    "Indra": "INDRA", "MII": "MII", "Sirtiv2": "SIRTI",
}

SOTTO_CORONA_TYPES = {"inst. sotto corona", "sotto corona"}


def normalize_fornitore(raw) -> str:
    if pd.isna(raw) or not raw: return "UNKNOWN"
    return FORNITORE_MAP.get(str(raw).strip(), str(raw).strip().upper())


def normalize_availability(raw_value) -> Tuple[str, str]:
    if pd.isna(raw_value) or raw_value == "" or raw_value is None:
        return ("UNKNOWN", "UNKNOWN")
    if isinstance(raw_value, (int, float)):
        val = int(raw_value)
        raw = AV_STATUS_NUMERIC.get(val, f"CODE_{val}")
        norm = "OK" if val in [1, 2] else "KO"
        return (raw, norm)
    val = str(raw_value).strip().upper()
    if val in ["ON", "AVAILABLE", "COMPLETE"]: return (val, "OK")
    elif val in ["OFF", "NO DATA", "NOT AVAILABLE", "KO"]: return (val, "KO")
    return (val, "UNKNOWN")


def parse_availability_date(col_name) -> Optional[date]:
    if isinstance(col_name, datetime): return col_name.date()
    match = re.match(r'AVAILABILITY\s+(\d{1,2})\s+(\w+)', str(col_name), re.IGNORECASE)
    if not match: return None
    day = int(match.group(1))
    month_map = {"gen":1,"feb":2,"mar":3,"apr":4,"mag":5,"giu":6,"lug":7,"ago":8,"set":9,"ott":10,"nov":11,"dic":12}
    month = month_map.get(match.group(2).lower()[:3])
    if month is None: return None
    year = 2025 if month == 12 else 2026
    try: return date(year, month, day)
    except: return None


def safe_str(val) -> Optional[str]:
    if pd.isna(val) or val is None: return None
    s = str(val).strip()
    return s if s and s.lower() != 'nan' else None


def safe_date(val) -> Optional[date]:
    if pd.isna(val) or val is None: return None
    if isinstance(val, datetime): return val.date()
    if isinstance(val, date): return val
    try: return pd.to_datetime(str(val)).date()
    except: return None


def is_sotto_corona(tipo_install: Optional[str]) -> bool:
    if not tipo_install: return False
    return tipo_install.strip().lower() in SOTTO_CORONA_TYPES


class ExcelImporter:
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.stats = {"devices_imported": 0, "availability_records": 0,
                      "tickets_new": 0, "tickets_updated": 0, "new_dates": [], "errors": []}

    def run(self) -> Dict:
        init_db()
        session = get_session()
        try:
            print("[1/5] Lettura sheet 'Stato'...")
            df_stato = pd.read_excel(self.file_path, sheet_name='Stato', engine='openpyxl', header=1)
            print(f"       {len(df_stato)} righe")

            print("[2/5] Lettura sheet 'Av Status'...")
            df_av = pd.read_excel(self.file_path, sheet_name='Av Status', engine='openpyxl')
            print(f"       {len(df_av)} righe")

            print("[3/5] Import dispositivi e availability...")
            self._import_devices(session, df_stato)
            self._import_availability_stato(session, df_stato)
            self._import_availability_av_status(session, df_av)

            print("[4/5] Aggiornamento storico ticket...")
            self._update_ticket_history(session)

            print("[5/5] Calcolo trend e stati...")
            self._compute_derived_states(session)

            log = ImportLog(filename=self.file_path.name, devices_total=self.stats["devices_imported"], status="OK")
            session.add(log)
            session.commit()
            print(f"\nImport completato: {self.stats['devices_imported']} dispositivi, "
                  f"{self.stats['availability_records']} record avail, "
                  f"{self.stats['tickets_new']} nuovi ticket, {self.stats['tickets_updated']} aggiornati")
            return self.stats
        except Exception as e:
            session.rollback()
            raise
        finally:
            session.close()

    def _import_devices(self, session, df_stato):
        for _, row in df_stato.iterrows():
            device_id = safe_str(row.get("DeviceID"))
            if not device_id: continue

            device = session.get(Device, device_id)
            if device is None:
                device = Device(device_id=device_id)
                session.add(device)

            device.tipo_install = safe_str(row.get("Tipo Installazione AM"))
            device.is_sotto_corona = is_sotto_corona(safe_str(row.get("Tipo Installazione AM")))
            device.linea = safe_str(row.get("Linea"))
            device.st_sostegno = safe_str(row.get("ST Sostegno"))
            sd = safe_str(row.get("Sistema DigiL"))
            device.sistema_digil = sd.lower() if sd else None
            device.dt = safe_str(row.get("DT"))
            device.denominazione = safe_str(row.get("Denominazione Linea"))
            device.ui = safe_str(row.get("UI"))
            device.regione = safe_str(row.get("Regione"))
            device.provincia = safe_str(row.get("Provincia"))
            device.ip_address = safe_str(row.get("IP address SIM"))
            device.rischio_neve = safe_str(row.get("Rischio neve"))
            device.fornitore_raw = safe_str(row.get("Fornitore"))
            device.fornitore = normalize_fornitore(row.get("Fornitore"))
            device.data_install = safe_date(row.get("Data Installazione Digil"))
            device.da_file_master = safe_str(row.get("Da file master"))
            device.check_lte = safe_str(row.get("Check LTE"))
            device.check_ssh = safe_str(row.get("check SSH"))
            device.batteria = safe_str(row.get("Batteria"))
            device.porta_aperta = safe_str(row.get("Porta aperta"))
            device.check_mongo = safe_str(row.get("Check Mongo"))
            device.tipo_malfunzionamento = safe_str(row.get("Tipo Malfunzionamento"))
            device.dettagli_malfunzionamento = safe_str(row.get("Eventuali Dettagli Malfunzionamento"))
            device.cluster_analisi = safe_str(row.get("Cluster Analisi"))
            device.analisi_malfunzionamento = safe_str(row.get("Analisi malfunzionamento"))
            device.tipologia_intervento = safe_str(row.get("Tipologia intervento"))
            device.strategia_risolutiva = safe_str(row.get("Strategia risolutiva"))
            device.risoluzione_attuata = safe_str(row.get("Risoluzione attuata"))
            device.note = safe_str(row.get("Note"))
            device.cause_anomalie_grezzo = safe_str(row.get("Cause di Anomalie GREZZO"))
            device.cause_anomalie = safe_str(row.get("Cause di Anomalie"))
            device.cluster_convertito = safe_str(row.get("Cluster convertito"))
            device.cluster_jira = safe_str(row.get("Cluster convertito Jira"))
            device.tipo_malf_jira = safe_str(row.get("Tipo Malf Jira"))
            device.ticket_id = safe_str(row.get("Ticket"))
            device.ticket_stato = safe_str(row.get("Stato Ticket"))
            device.ticket_data_apertura = safe_date(row.get("Data apertura ticket"))
            device.ticket_data_risoluzione = safe_date(row.get("Data risoluzione"))
            for col in df_stato.columns:
                if 'Unnamed: 68' in str(col) or col == df_stato.columns[-1]:
                    device.misure_mancanti = safe_str(row.get(col))
                    break
            self.stats["devices_imported"] += 1
        session.flush()

    def _update_ticket_history(self, session):
        """Aggiorna storico ticket per ogni device che ha un ticket_id.
        Se il ticket esiste già in history → aggiorna last_seen e stato.
        Se è un ticket nuovo → crea record storico."""
        now = datetime.utcnow()
        for device in session.query(Device).filter(Device.ticket_id.isnot(None)).all():
            if not device.ticket_id:
                continue

            existing = (session.query(TicketHistory)
                        .filter(TicketHistory.device_id == device.device_id,
                                TicketHistory.ticket_id == device.ticket_id)
                        .first())

            if existing:
                existing.last_seen = now
                existing.ticket_stato = device.ticket_stato
                existing.ticket_data_risoluzione = device.ticket_data_risoluzione
                if device.tipo_malfunzionamento: existing.tipo_malfunzionamento = device.tipo_malfunzionamento
                if device.cluster_analisi: existing.cluster_analisi = device.cluster_analisi
                if device.analisi_malfunzionamento: existing.analisi_malfunzionamento = device.analisi_malfunzionamento
                if device.tipologia_intervento: existing.tipologia_intervento = device.tipologia_intervento
                if device.strategia_risolutiva: existing.strategia_risolutiva = device.strategia_risolutiva
                if device.risoluzione_attuata: existing.risoluzione_attuata = device.risoluzione_attuata
                if device.cause_anomalie: existing.cause_anomalie = device.cause_anomalie
                if device.note: existing.note = device.note
                if device.cluster_jira: existing.cluster_jira = device.cluster_jira
                if device.tipo_malf_jira: existing.tipo_malf_jira = device.tipo_malf_jira
                self.stats["tickets_updated"] += 1
            else:
                session.add(TicketHistory(
                    device_id=device.device_id,
                    ticket_id=device.ticket_id,
                    ticket_stato=device.ticket_stato,
                    ticket_data_apertura=device.ticket_data_apertura,
                    ticket_data_risoluzione=device.ticket_data_risoluzione,
                    tipo_malfunzionamento=device.tipo_malfunzionamento,
                    cluster_analisi=device.cluster_analisi,
                    analisi_malfunzionamento=device.analisi_malfunzionamento,
                    tipologia_intervento=device.tipologia_intervento,
                    strategia_risolutiva=device.strategia_risolutiva,
                    risoluzione_attuata=device.risoluzione_attuata,
                    cause_anomalie=device.cause_anomalie,
                    note=device.note,
                    cluster_jira=device.cluster_jira,
                    tipo_malf_jira=device.tipo_malf_jira,
                    first_seen=now, last_seen=now,
                ))
                self.stats["tickets_new"] += 1
        session.flush()

    def _import_availability_stato(self, session, df_stato):
        avail_cols = [c for c in df_stato.columns if 'AVAILABILITY' in str(c).upper()]
        known_ids = {d.device_id for d in session.query(Device.device_id).all()}
        for col in avail_cols:
            check_date = parse_availability_date(col)
            if check_date is None: continue
            for _, row in df_stato.iterrows():
                device_id = safe_str(row.get("DeviceID"))
                if not device_id or device_id not in known_ids: continue
                raw_status, norm_status = normalize_availability(row.get(col))
                if norm_status == "UNKNOWN": continue
                existing = session.get(AvailabilityDaily, (device_id, check_date))
                if existing is None:
                    session.add(AvailabilityDaily(device_id=device_id, check_date=check_date,
                                                  raw_status=raw_status, norm_status=norm_status))
                    self.stats["availability_records"] += 1
                else:
                    existing.raw_status = raw_status; existing.norm_status = norm_status
        session.flush()

    def _import_availability_av_status(self, session, df_av):
        known_ids = {d.device_id for d in session.query(Device.device_id).all()}
        date_cols = [c for c in df_av.columns if isinstance(c, datetime)]
        skipped = 0
        for col in date_cols:
            check_date = col.date()
            self.stats["new_dates"].append(check_date)
            for _, row in df_av.iterrows():
                device_id = safe_str(row.get("DeviceID"))
                if not device_id: continue
                if device_id not in known_ids:
                    skipped += 1; continue
                raw_status, norm_status = normalize_availability(row.get(col))
                if norm_status == "UNKNOWN": continue
                existing = session.get(AvailabilityDaily, (device_id, check_date))
                if existing is None:
                    session.add(AvailabilityDaily(device_id=device_id, check_date=check_date,
                                                  raw_status=raw_status, norm_status=norm_status))
                    self.stats["availability_records"] += 1
                else:
                    existing.raw_status = raw_status; existing.norm_status = norm_status
        if skipped: print(f"       {skipped} record Av Status ignorati (device non in Stato)")
        session.flush()

    def _compute_derived_states(self, session):
        for device in session.query(Device).all():
            avail = (session.query(AvailabilityDaily)
                     .filter(AvailabilityDaily.device_id == device.device_id)
                     .order_by(AvailabilityDaily.check_date.desc()).limit(30).all())
            if not avail:
                device.current_health = "UNKNOWN"; continue
            latest = avail[0]
            device.last_avail_status = latest.raw_status
            device.last_avail_norm = latest.norm_status
            device.last_avail_date = latest.check_date
            recent_7 = sorted(avail[:7], key=lambda r: r.check_date)
            device.trend_7d = "".join("O" if r.norm_status == "OK" else "K" for r in recent_7)
            current_norm = latest.norm_status
            days = 0
            for r in avail:
                if r.norm_status == current_norm: days += 1
                else: break
            device.days_in_current = days
            if latest.norm_status == "OK":
                device.current_health = "DEGRADED" if (device.check_lte == "KO" or device.porta_aperta == "KO") else "OK"
            else:
                device.current_health = "KO"
        session.commit()


def run_import(file_path: str) -> Dict:
    return ExcelImporter(file_path).run()
