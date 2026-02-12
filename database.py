"""
DIGIL Monitoring - Database Models
"""
from datetime import datetime, date
from sqlalchemy import (
    create_engine, Column, String, Integer, Float, Boolean, Date, DateTime,
    Text, ForeignKey, Index, event
)
from sqlalchemy.orm import declarative_base, sessionmaker, relationship
from pathlib import Path

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "data" / "digil_monitoring.db"

engine = create_engine(f"sqlite:///{DB_PATH}", echo=False)
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()


@event.listens_for(engine, "connect")
def set_sqlite_pragma(dbapi_conn, connection_record):
    c = dbapi_conn.cursor()
    c.execute("PRAGMA journal_mode=WAL")
    c.execute("PRAGMA synchronous=NORMAL")
    c.execute("PRAGMA foreign_keys=ON")
    c.close()


class Device(Base):
    __tablename__ = "devices"
    device_id = Column(String, primary_key=True)
    tipo_install = Column(String)
    is_sotto_corona = Column(Boolean, default=False)
    linea = Column(String)
    st_sostegno = Column(String)
    sistema_digil = Column(String)
    note_piano_lora = Column(String)
    dt = Column(String)
    denominazione = Column(String)
    ui = Column(String)
    regione = Column(String)
    provincia = Column(String)
    ip_address = Column(String)
    rischio_neve = Column(String)
    fornitore = Column(String)
    fornitore_raw = Column(String)
    data_install = Column(Date)
    da_file_master = Column(String)
    misure_mancanti = Column(Text)
    check_lte = Column(String)
    check_ssh = Column(String)
    batteria = Column(String)
    porta_aperta = Column(String)
    check_mongo = Column(String)
    last_avail_status = Column(String)
    last_avail_norm = Column(String)
    last_avail_date = Column(Date)
    days_in_current = Column(Integer, default=0)
    trend_7d = Column(String, default="")
    current_health = Column(String, default="UNKNOWN")
    tipo_malfunzionamento = Column(String)
    dettagli_malfunzionamento = Column(Text)
    cluster_analisi = Column(String)
    analisi_malfunzionamento = Column(Text)
    tipologia_intervento = Column(String)
    strategia_risolutiva = Column(Text)
    risoluzione_attuata = Column(Text)
    note = Column(Text)
    cause_anomalie_grezzo = Column(String)
    cause_anomalie = Column(String)
    cluster_risoluzioni = Column(String)
    cluster_convertito = Column(String)
    cluster_jira = Column(String)
    tipo_malf_jira = Column(String)
    ticket_id = Column(String)
    ticket_stato = Column(String)
    ticket_data_apertura = Column(Date)
    ticket_data_risoluzione = Column(Date)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    availability = relationship("AvailabilityDaily", back_populates="device", cascade="all, delete-orphan")
    events = relationship("AnomalyEvent", back_populates="device", cascade="all, delete-orphan")
    ticket_history = relationship("TicketHistory", back_populates="device", cascade="all, delete-orphan",
                                  order_by="TicketHistory.first_seen.desc()")


class AvailabilityDaily(Base):
    __tablename__ = "availability_daily"
    device_id = Column(String, ForeignKey("devices.device_id"), primary_key=True)
    check_date = Column(Date, primary_key=True)
    raw_status = Column(String)   # COMPLETE, AVAILABLE, NOT AVAILABLE, NO DATA
    norm_status = Column(String)  # OK or KO
    device = relationship("Device", back_populates="availability")
    __table_args__ = (Index("idx_avail_date", "check_date"),)


class AnomalyEvent(Base):
    __tablename__ = "anomaly_events"
    id = Column(Integer, primary_key=True, autoincrement=True)
    device_id = Column(String, ForeignKey("devices.device_id"), nullable=False)
    event_date = Column(Date, nullable=False)
    event_type = Column(String, nullable=False)
    severity = Column(String, nullable=False)
    description = Column(Text)
    context_json = Column(Text)
    acknowledged = Column(Boolean, default=False)
    acknowledged_at = Column(DateTime)
    action_taken = Column(Text)
    related_ticket = Column(String)
    created_at = Column(DateTime, default=datetime.utcnow)
    device = relationship("Device", back_populates="events")
    __table_args__ = (Index("idx_ev_sev", "severity"), Index("idx_ev_ack", "acknowledged"),)


class TicketHistory(Base):
    __tablename__ = "ticket_history"
    id = Column(Integer, primary_key=True, autoincrement=True)
    device_id = Column(String, ForeignKey("devices.device_id"), nullable=False)
    ticket_id = Column(String, nullable=False)
    ticket_stato = Column(String)
    ticket_data_apertura = Column(Date)
    ticket_data_risoluzione = Column(Date)
    tipo_malfunzionamento = Column(String)
    cluster_analisi = Column(String)
    analisi_malfunzionamento = Column(Text)
    tipologia_intervento = Column(String)
    strategia_risolutiva = Column(Text)
    risoluzione_attuata = Column(Text)
    cause_anomalie = Column(String)
    note = Column(Text)
    cluster_jira = Column(String)
    tipo_malf_jira = Column(String)
    first_seen = Column(DateTime, default=datetime.utcnow)
    last_seen = Column(DateTime, default=datetime.utcnow)
    device = relationship("Device", back_populates="ticket_history")
    __table_args__ = (Index("idx_th_device", "device_id"), Index("idx_th_ticket", "ticket_id"),)


class ImportLog(Base):
    __tablename__ = "import_log"
    id = Column(Integer, primary_key=True, autoincrement=True)
    filename = Column(String)
    import_date = Column(DateTime, default=datetime.utcnow)
    devices_total = Column(Integer)
    alerts_generated = Column(Integer, default=0)
    status = Column(String, default="OK")
    error_message = Column(Text)


def init_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    Base.metadata.create_all(engine)

def get_session():
    return SessionLocal()
