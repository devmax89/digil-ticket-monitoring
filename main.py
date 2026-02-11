"""
DIGIL Monitoring Dashboard - PyQt5
Stile Terna: bianco, blu #0066CC, grigio.
"""
import sys, json
from pathlib import Path
from datetime import datetime, date
from typing import Optional, List, Dict
from collections import Counter

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QProgressBar,
    QGroupBox, QFileDialog, QMessageBox, QTabWidget, QHeaderView,
    QAbstractItemView, QStatusBar, QFrame, QLineEdit, QComboBox,
    QDialog, QTextEdit, QScrollArea, QSplitter, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSortFilterProxyModel
from PyQt5.QtGui import QColor, QFont, QBrush

from database import get_session, init_db, Device, AvailabilityDaily, AnomalyEvent, ImportLog
from importer import run_import
from detection import run_detection

# ============================================================
# STYLE
# ============================================================
STYLE = """
QMainWindow { background-color: #FFFFFF; }
QWidget { font-family: 'Segoe UI', Arial; font-size: 12px; }
QLabel#headerTitle { font-size: 20px; font-weight: bold; color: #0066CC; }
QLabel#headerSubtitle { font-size: 11px; color: #666666; }
QGroupBox { font-weight: bold; border: 1px solid #CCCCCC; border-radius: 5px; margin-top: 10px; padding-top: 8px; background: #FAFAFA; }
QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 4px; color: #0066CC; }
QPushButton { background: #0066CC; color: white; border: none; padding: 6px 14px; border-radius: 3px; font-weight: bold; }
QPushButton:hover { background: #004C99; }
QPushButton:disabled { background: #CCCCCC; color: #666666; }
QPushButton#secondary { background: white; color: #0066CC; border: 1px solid #0066CC; }
QPushButton#secondary:hover { background: #E8F0FE; }
QPushButton#danger { background: #CC3300; }
QPushButton#success { background: #2E7D32; }
QTableWidget { border: 1px solid #CCCCCC; gridline-color: #E0E0E0; background: white; alternate-background-color: #F8FBFF; font-size: 11px; }
QTableWidget::item { padding: 3px 5px; }
QTableWidget::item:selected { background: #CCE5FF; color: black; }
QHeaderView::section { background: #0066CC; color: white; padding: 6px; border: none; font-weight: bold; font-size: 11px; }
QTabWidget::pane { border: 1px solid #CCCCCC; background: white; }
QTabBar::tab { background: #F0F0F0; border: 1px solid #CCCCCC; padding: 8px 18px; font-weight: bold; }
QTabBar::tab:selected { background: #0066CC; color: white; }
QTabBar::tab:hover:!selected { background: #E8F0FE; }
QLineEdit { border: 1px solid #CCCCCC; border-radius: 3px; padding: 4px 8px; }
QLineEdit:focus { border-color: #0066CC; }
QComboBox { border: 1px solid #CCCCCC; border-radius: 3px; padding: 4px 8px; }
QProgressBar { border: 1px solid #CCCCCC; border-radius: 3px; text-align: center; background: #F0F0F0; height: 22px; }
QProgressBar::chunk { background: #0066CC; border-radius: 2px; }
QStatusBar { background: #F5F5F5; border-top: 1px solid #CCCCCC; }
QToolTip { background: #FFFFDD; color: black; border: 1px solid #333; padding: 4px; font-size: 12px; }
"""

SEV_COLORS = {"CRITICAL": "#C62828", "HIGH": "#E65100", "MEDIUM": "#F9A825", "LOW": "#2E7D32"}
SEV_BG = {"CRITICAL": "#FFEBEE", "HIGH": "#FFF3E0", "MEDIUM": "#FFFDE7", "LOW": "#E8F5E9"}
HEALTH_BG = {"OK": "#E8F5E9", "KO": "#FFEBEE", "DEGRADED": "#FFF3E0", "UNKNOWN": "#F5F5F5"}
CHECK_BG = {"OK": "#E8F5E9", "KO": "#FFEBEE"}


def colored_item(text, bg=None, fg=None, bold=False):
    item = QTableWidgetItem(str(text) if text else "-")
    item.setTextAlignment(Qt.AlignCenter)
    if bg: item.setBackground(QColor(bg))
    if fg: item.setForeground(QColor(fg))
    if bold: item.setFont(QFont("Segoe UI", 11, QFont.Bold))
    return item


def check_item(val):
    if val == "OK": return colored_item("OK", "#E8F5E9", "#2E7D32", True)
    elif val == "KO": return colored_item("KO", "#FFEBEE", "#C62828", True)
    return colored_item(val or "-", "#F5F5F5", "#757575")


def trend_str(t):
    """Returns trend as unicode blocks for display"""
    if not t: return "-"
    return "".join("■" if c == "O" else "□" for c in t)


# ============================================================
# IMPORT THREAD
# ============================================================
class ImportThread(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(dict, int)
    error = pyqtSignal(str)

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def run(self):
        try:
            self.progress.emit("Importazione in corso...")
            stats = run_import(self.file_path)
            self.progress.emit("Generazione alert...")
            alert_count = run_detection()
            self.finished.emit(stats, alert_count)
        except Exception as e:
            self.error.emit(str(e))


# ============================================================
# FILTERABLE TABLE
# ============================================================
class FilterableTable(QWidget):
    """Tabella con filtri testo inline sotto l'header"""

    def __init__(self, columns: List[str], parent=None):
        super().__init__(parent)
        self.columns = columns
        self._all_data = []
        self._filtered_data = []
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)

        # Filter row
        filter_widget = QWidget()
        filter_layout = QHBoxLayout(filter_widget)
        filter_layout.setContentsMargins(2, 2, 2, 2)
        filter_layout.setSpacing(2)

        self.filters = {}
        for col in columns:
            le = QLineEdit()
            le.setPlaceholderText(f"Filtra {col}")
            le.setMaximumHeight(22)
            le.setStyleSheet("font-size: 10px; padding: 1px 4px;")
            le.textChanged.connect(self._apply_filters)
            self.filters[col] = le
            filter_layout.addWidget(le)

        layout.addWidget(filter_widget)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(len(columns))
        self.table.setHorizontalHeaderLabels(columns)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSortingEnabled(True)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setDefaultSectionSize(24)
        header = self.table.horizontalHeader()
        header.setStretchLastSection(True)
        layout.addWidget(self.table)

    def clear_filters(self):
        for le in self.filters.values():
            le.blockSignals(True)
            le.clear()
            le.blockSignals(False)
        self._apply_filters()

    def set_data(self, data: List[Dict], render_fn=None):
        self._all_data = data
        self._render_fn = render_fn
        self._apply_filters()

    def _apply_filters(self):
        active_filters = {}
        for col, le in self.filters.items():
            txt = le.text().strip().lower()
            if txt:
                active_filters[col] = txt

        if active_filters:
            filtered = []
            for row_data in self._all_data:
                match = True
                for col, txt in active_filters.items():
                    val = str(row_data.get(col, "")).lower()
                    if txt not in val:
                        match = False
                        break
                if match:
                    filtered.append(row_data)
            self._filtered_data = filtered
        else:
            self._filtered_data = list(self._all_data)

        self._populate_table()

    def _populate_table(self):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(self._filtered_data))

        for row_idx, row_data in enumerate(self._filtered_data):
            for col_idx, col_name in enumerate(self.columns):
                if self._render_fn:
                    item = self._render_fn(row_data, col_name)
                else:
                    item = QTableWidgetItem(str(row_data.get(col_name, "")))
                if item:
                    self.table.setItem(row_idx, col_idx, item)

        self.table.setSortingEnabled(True)

    def get_row_count(self):
        return len(self._filtered_data)


# ============================================================
# DEVICE DETAIL DIALOG
# ============================================================
class DeviceDetailDialog(QDialog):
    def __init__(self, device_id, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Dettaglio: {device_id}")
        self.setMinimumSize(700, 600)
        self.setStyleSheet(STYLE)

        session = get_session()
        try:
            device = session.get(Device, device_id)
            if not device:
                QMessageBox.warning(self, "Errore", f"Device {device_id} non trovato")
                return

            layout = QVBoxLayout(self)

            # Top info
            top = QHBoxLayout()
            title = QLabel(device_id)
            title.setStyleSheet("font-size: 16px; font-weight: bold; color: #0066CC;")
            top.addWidget(title)
            health_lbl = QLabel(device.current_health or "?")
            health_bg = HEALTH_BG.get(device.current_health, "#F5F5F5")
            health_lbl.setStyleSheet(f"background: {health_bg}; padding: 4px 12px; border-radius: 3px; font-weight: bold;")
            top.addWidget(health_lbl)
            if device.is_sotto_corona:
                sc = QLabel("SOTTO CORONA")
                sc.setStyleSheet("background: #E3F2FD; color: #1565C0; padding: 4px 8px; border-radius: 3px; font-weight: bold;")
                top.addWidget(sc)
            top.addStretch()
            layout.addLayout(top)

            # Scroll area
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            content = QWidget()
            cl = QVBoxLayout(content)

            # Grid: Anagrafica + Diagnostica
            grid = QGridLayout()
            grid.setSpacing(4)
            info = [
                ("Linea", device.linea), ("Sostegno", device.st_sostegno),
                ("Fornitore", device.fornitore), ("Tipo", device.sistema_digil),
                ("DT", device.dt), ("Denominazione", device.denominazione),
                ("Regione", f"{device.regione or ''} / {device.provincia or ''}"),
                ("IP", device.ip_address), ("Installazione", str(device.data_install) if device.data_install else "-"),
                ("Tipo Install.", device.tipo_install), ("Da file master", device.da_file_master),
            ]
            r = 0
            for lbl, val in info:
                if val:
                    grid.addWidget(QLabel(f"<b style='color:#666'>{lbl}:</b>"), r, 0)
                    grid.addWidget(QLabel(str(val)), r, 1)
                    r += 1

            grp_ana = QGroupBox("Anagrafica")
            grp_ana.setLayout(grid)
            cl.addWidget(grp_ana)

            # Diagnostica
            diag_layout = QHBoxLayout()
            for name, val in [("LTE", device.check_lte), ("SSH", device.check_ssh), ("Mongo", device.check_mongo),
                              ("Batteria", device.batteria), ("Porta", device.porta_aperta)]:
                box = QLabel(f"<center><small>{name}</small><br><b>{val or '-'}</b></center>")
                bg = "#E8F5E9" if val == "OK" else "#FFEBEE" if val == "KO" else "#F5F5F5"
                box.setStyleSheet(f"background: {bg}; border: 1px solid #CCC; border-radius: 4px; padding: 6px; min-width: 60px;")
                diag_layout.addWidget(box)
            diag_layout.addWidget(QLabel(f"<b>Trend 7d:</b> {trend_str(device.trend_7d)}"))
            diag_layout.addWidget(QLabel(f"<b>Giorni:</b> {device.days_in_current}"))
            diag_layout.addStretch()
            grp_diag = QGroupBox("Diagnostica")
            grp_diag.setLayout(diag_layout)
            cl.addWidget(grp_diag)

            # Malfunzionamento
            malf_info = [
                ("Tipo", device.tipo_malfunzionamento), ("Cluster", device.cluster_analisi),
                ("Analisi", device.analisi_malfunzionamento), ("Intervento", device.tipologia_intervento),
                ("Strategia", device.strategia_risolutiva), ("Cause", device.cause_anomalie),
                ("Risoluzione", device.risoluzione_attuata), ("Note", device.note),
            ]
            if any(v for _, v in malf_info):
                malf_grid = QGridLayout()
                malf_grid.setSpacing(4)
                mr = 0
                for lbl, val in malf_info:
                    if val:
                        malf_grid.addWidget(QLabel(f"<b style='color:#666'>{lbl}:</b>"), mr, 0, Qt.AlignTop)
                        vl = QLabel(str(val))
                        vl.setWordWrap(True)
                        malf_grid.addWidget(vl, mr, 1)
                        mr += 1
                grp_malf = QGroupBox("Malfunzionamento")
                grp_malf.setLayout(malf_grid)
                cl.addWidget(grp_malf)

            # Ticket
            if device.ticket_id:
                t_layout = QGridLayout()
                t_layout.addWidget(QLabel("<b style='color:#666'>ID:</b>"), 0, 0)
                t_layout.addWidget(QLabel(f"<b>{device.ticket_id}</b>"), 0, 1)
                t_layout.addWidget(QLabel("<b style='color:#666'>Stato:</b>"), 1, 0)
                stato_bg = "#FFEBEE" if device.ticket_stato == "Aperto" else "#E8F5E9"
                sl = QLabel(device.ticket_stato or "-")
                sl.setStyleSheet(f"background: {stato_bg}; padding: 2px 8px; border-radius: 3px;")
                t_layout.addWidget(sl, 1, 1)
                if device.ticket_data_apertura:
                    t_layout.addWidget(QLabel("<b style='color:#666'>Apertura:</b>"), 2, 0)
                    t_layout.addWidget(QLabel(str(device.ticket_data_apertura)), 2, 1)
                grp_ticket = QGroupBox("Ticket")
                grp_ticket.setLayout(t_layout)
                cl.addWidget(grp_ticket)

            # Timeline
            avail = (session.query(AvailabilityDaily)
                     .filter(AvailabilityDaily.device_id == device_id)
                     .order_by(AvailabilityDaily.check_date).all())
            if avail:
                tl_layout = QHBoxLayout()
                for a in avail:
                    raw_up = (a.raw_status or "").upper()
                    if a.norm_status == "OK":
                        bg = "#2E7D32"
                    elif raw_up in ("NOT AVAILABLE", "NO DATA", "CODE_3", "CODE_4"):
                        bg = "#F9A825"  # giallo per NOT AVAILABLE / NO DATA
                    elif a.norm_status == "KO":
                        bg = "#C62828"  # rosso per OFF
                    else:
                        bg = "#BDBDBD"
                    box = QLabel("")
                    box.setFixedSize(14, 14)
                    box.setStyleSheet(f"background: {bg}; border-radius: 2px;")
                    box.setToolTip(f"{a.check_date}: {a.raw_status}")
                    tl_layout.addWidget(box)
                tl_layout.addStretch()
                # Legend
                legend = QHBoxLayout()
                for lbl, clr in [("OK", "#2E7D32"), ("OFF", "#C62828"), ("NOT AVAILABLE", "#F9A825")]:
                    sq = QLabel("")
                    sq.setFixedSize(10, 10)
                    sq.setStyleSheet(f"background: {clr}; border-radius: 1px;")
                    legend.addWidget(sq)
                    legend.addWidget(QLabel(f"<small>{lbl}</small>"))
                    legend.addSpacing(8)
                legend.addStretch()
                tl_outer = QVBoxLayout()
                tl_outer.addLayout(tl_layout)
                tl_outer.addLayout(legend)
                grp_tl = QGroupBox(f"Timeline Availability ({len(avail)} giorni)")
                grp_tl.setLayout(tl_outer)
                cl.addWidget(grp_tl)

            # Alerts
            events = (session.query(AnomalyEvent)
                      .filter(AnomalyEvent.device_id == device_id)
                      .order_by(AnomalyEvent.created_at.desc()).limit(10).all())
            if events:
                evt_table = QTableWidget()
                evt_table.setColumnCount(4)
                evt_table.setHorizontalHeaderLabels(["Data", "Tipo", "Severity", "Descrizione"])
                evt_table.setRowCount(len(events))
                evt_table.horizontalHeader().setStretchLastSection(True)
                for i, e in enumerate(events):
                    evt_table.setItem(i, 0, QTableWidgetItem(str(e.event_date)))
                    evt_table.setItem(i, 1, QTableWidgetItem(e.event_type))
                    evt_table.setItem(i, 2, colored_item(e.severity, SEV_BG.get(e.severity), SEV_COLORS.get(e.severity), True))
                    evt_table.setItem(i, 3, QTableWidgetItem(e.description or ""))
                grp_evt = QGroupBox("Alert Recenti")
                el = QVBoxLayout()
                el.addWidget(evt_table)
                grp_evt.setLayout(el)
                cl.addWidget(grp_evt)

            cl.addStretch()
            scroll.setWidget(content)
            layout.addWidget(scroll)

            close_btn = QPushButton("Chiudi")
            close_btn.clicked.connect(self.close)
            layout.addWidget(close_btn, alignment=Qt.AlignRight)
        finally:
            session.close()


# ============================================================
# MAIN WINDOW
# ============================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.import_thread = None
        self.init_ui()
        self.setStyleSheet(STYLE)
        init_db()
        self.refresh_data()

    def init_ui(self):
        self.setWindowTitle("DIGIL Monitoring Dashboard - Terna IoT Team")
        self.setMinimumSize(1400, 900)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(12, 12, 12, 12)

        # Header
        hdr = QHBoxLayout()
        self.logo_label = QLabel("TERNA")
        self.logo_label.setStyleSheet("background: #0066CC; color: white; font-size: 18px; font-weight: bold; padding: 8px 16px; border-radius: 5px;")
        self.logo_label.setFixedHeight(40)
        self.logo_label.setFixedWidth(120)
        self.logo_label.setAlignment(Qt.AlignCenter)
        # Try loading logo from assets/logo_terna.png
        from PyQt5.QtGui import QPixmap
        script_dir = Path(__file__).parent
        logo_path = script_dir / "assets" / "logo_terna.png"
        if logo_path.exists():
            pix = QPixmap(str(logo_path))
            if not pix.isNull():
                self.logo_label.setPixmap(pix.scaled(120, 40, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                self.logo_label.setText("")
                self.logo_label.setStyleSheet("background: transparent;")
        hdr.addWidget(self.logo_label)
        title_w = QWidget()
        tl = QVBoxLayout(title_w)
        tl.setContentsMargins(10, 0, 0, 0)
        tl.setSpacing(0)
        t = QLabel("DIGIL Monitoring Dashboard")
        t.setObjectName("headerTitle")
        tl.addWidget(t)
        s = QLabel("Monitoraggio IoT — Terna S.p.A.")
        s.setObjectName("headerSubtitle")
        tl.addWidget(s)
        hdr.addWidget(title_w)
        hdr.addStretch()
        self.import_btn = QPushButton("📂 Importa Excel")
        self.import_btn.setObjectName("secondary")
        self.import_btn.clicked.connect(self.do_import)
        hdr.addWidget(self.import_btn)
        main_layout.addLayout(hdr)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("background: #CCCCCC; max-height: 1px;")
        main_layout.addWidget(sep)

        # KPI Cards
        self.cards_layout = QHBoxLayout()
        self.card_total = self._make_card("0", "Dispositivi")
        self.card_ok = self._make_card("0", "OK", "#E8F5E9")
        self.card_ko = self._make_card("0", "KO", "#FFEBEE")
        self.card_deg = self._make_card("0", "Degraded", "#FFF3E0")
        self.card_crit = self._make_card("0", "Alert Critical", "#FFEBEE")
        self.card_high = self._make_card("0", "Alert High", "#FFF3E0")
        self.card_tickets = self._make_card("0", "Ticket Aperti")
        for c in [self.card_total, self.card_ok, self.card_ko, self.card_deg, self.card_crit, self.card_high, self.card_tickets]:
            self.cards_layout.addWidget(c)
        main_layout.addLayout(self.cards_layout)

        # Tabs
        self.tabs = QTabWidget()
        self.tabs.addTab(self._create_alerts_tab(), "⚠ Alert")
        self.tabs.addTab(self._create_devices_tab(), "📋 Dispositivi")
        self.tabs.addTab(self._create_overview_tab(), "📊 Overview")
        self.tabs.currentChanged.connect(self._on_tab_changed)
        main_layout.addWidget(self.tabs)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_label = QLabel("Pronto")
        self.status_bar.addWidget(self.status_label, stretch=1)

    def _make_card(self, value, label, bg="#FFFFFF"):
        card = QWidget()
        card.setStyleSheet(f"background: {bg}; border: 1px solid #CCC; border-radius: 5px;")
        card.setFixedHeight(78)
        card.setMinimumWidth(100)
        cl = QVBoxLayout(card)
        cl.setSpacing(2)
        cl.setContentsMargins(12, 8, 12, 8)
        v = QLabel(value)
        v.setObjectName("cardValue")
        v_font = QFont("Segoe UI", 30, QFont.Bold)
        v.setFont(v_font)
        v.setStyleSheet("background: transparent; border: none; color: #333;")
        cl.addWidget(v)
        l = QLabel(label)
        l.setStyleSheet("font-size: 10px; color: #666; background: transparent; border: none;")
        cl.addWidget(l)
        card._value_label = v
        return card

    def _update_card(self, card, value):
        card._value_label.setText(str(value))

    # ========== ALERT TAB ==========
    def _create_alerts_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Top filters
        flt = QHBoxLayout()
        flt.addWidget(QLabel("Severity:"))
        self.alert_sev = QComboBox()
        self.alert_sev.addItems(["Tutti", "CRITICAL", "HIGH", "MEDIUM", "LOW"])
        self.alert_sev.currentTextChanged.connect(self.refresh_alerts)
        flt.addWidget(self.alert_sev)
        flt.addWidget(QLabel("Tipo:"))
        self.alert_type = QComboBox()
        self.alert_type.addItem("Tutti")
        self.alert_type.currentTextChanged.connect(self.refresh_alerts)
        flt.addWidget(self.alert_type)
        flt.addWidget(QLabel("Fornitore:"))
        self.alert_forn = QComboBox()
        self.alert_forn.addItems(["Tutti", "INDRA", "MII", "SIRTI"])
        self.alert_forn.currentTextChanged.connect(self.refresh_alerts)
        flt.addWidget(self.alert_forn)
        flt.addStretch()
        self.alert_count_label = QLabel("")
        self.alert_count_label.setStyleSheet("color: #666;")
        flt.addWidget(self.alert_count_label)
        clear_btn = QPushButton("Pulisci Filtri")
        clear_btn.setObjectName("secondary")
        clear_btn.clicked.connect(self._clear_alert_filters)
        flt.addWidget(clear_btn)
        layout.addLayout(flt)

        # Table
        cols = ["Severity", "Tipo", "DeviceID", "Fornitore", "DT", "Trend", "LTE", "SSH", "Mongo", "Batt", "Porta", "Sotto C.", "Descrizione", "Ticket"]
        self.alert_table = FilterableTable(cols)
        self.alert_table.table.setColumnWidth(0, 75)
        self.alert_table.table.setColumnWidth(1, 180)
        self.alert_table.table.setColumnWidth(2, 100)
        self.alert_table.table.setColumnWidth(3, 60)
        self.alert_table.table.setColumnWidth(4, 55)
        self.alert_table.table.setColumnWidth(5, 60)
        for i in range(6, 11):
            self.alert_table.table.setColumnWidth(i, 45)
        self.alert_table.table.setColumnWidth(11, 50)
        self.alert_table.table.setColumnWidth(12, 300)
        self.alert_table.table.doubleClicked.connect(self._on_alert_dblclick)
        layout.addWidget(self.alert_table)
        return tab

    def _clear_alert_filters(self):
        self.alert_sev.setCurrentIndex(0)
        self.alert_type.setCurrentIndex(0)
        self.alert_forn.setCurrentIndex(0)
        self.alert_table.clear_filters()

    def _on_alert_dblclick(self, index):
        row = index.row()
        did_item = self.alert_table.table.item(row, 2)
        if did_item:
            did = did_item.data(Qt.UserRole) or did_item.text()
            dlg = DeviceDetailDialog(did, self)
            dlg.exec_()

    # ========== DEVICES TAB ==========
    def _create_devices_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        flt = QHBoxLayout()
        flt.addWidget(QLabel("Fornitore:"))
        self.dev_forn = QComboBox()
        self.dev_forn.addItems(["Tutti", "INDRA", "MII", "SIRTI"])
        self.dev_forn.currentTextChanged.connect(self.refresh_devices)
        flt.addWidget(self.dev_forn)
        flt.addWidget(QLabel("Health:"))
        self.dev_health = QComboBox()
        self.dev_health.addItems(["Tutti", "OK", "KO", "DEGRADED"])
        self.dev_health.currentTextChanged.connect(self.refresh_devices)
        flt.addWidget(self.dev_health)
        flt.addWidget(QLabel("Tipo:"))
        self.dev_tipo = QComboBox()
        self.dev_tipo.addItems(["Tutti", "master", "slave"])
        self.dev_tipo.currentTextChanged.connect(self.refresh_devices)
        flt.addWidget(self.dev_tipo)
        flt.addWidget(QLabel("Install:"))
        self.dev_install = QComboBox()
        self.dev_install.addItems(["Tutti", "Completa", "Sotto corona"])
        self.dev_install.currentTextChanged.connect(self.refresh_devices)
        flt.addWidget(self.dev_install)
        flt.addStretch()
        self.dev_count_label = QLabel("")
        self.dev_count_label.setStyleSheet("color: #666;")
        flt.addWidget(self.dev_count_label)
        clear_btn2 = QPushButton("Pulisci Filtri")
        clear_btn2.setObjectName("secondary")
        clear_btn2.clicked.connect(self._clear_dev_filters)
        flt.addWidget(clear_btn2)
        layout.addLayout(flt)

        cols = ["DeviceID", "Linea", "Fornitore", "Tipo", "DT", "Health", "LTE", "SSH", "Mongo", "Batt", "Porta", "Sotto C.", "Trend", "Giorni", "Malf.", "Ticket"]
        self.dev_table = FilterableTable(cols)
        self.dev_table.table.setColumnWidth(0, 100)
        self.dev_table.table.setColumnWidth(1, 70)
        self.dev_table.table.setColumnWidth(2, 60)
        self.dev_table.table.setColumnWidth(3, 48)
        self.dev_table.table.setColumnWidth(4, 50)
        for i in range(5, 12):
            self.dev_table.table.setColumnWidth(i, 48)
        self.dev_table.table.setColumnWidth(12, 60)
        self.dev_table.table.setColumnWidth(13, 42)
        self.dev_table.table.doubleClicked.connect(self._on_dev_dblclick)
        layout.addWidget(self.dev_table)
        return tab

    def _clear_dev_filters(self):
        self.dev_forn.setCurrentIndex(0)
        self.dev_health.setCurrentIndex(0)
        self.dev_tipo.setCurrentIndex(0)
        self.dev_install.setCurrentIndex(0)
        self.dev_table.clear_filters()

    def _on_dev_dblclick(self, index):
        row = index.row()
        did_item = self.dev_table.table.item(row, 0)
        if did_item:
            did = did_item.data(Qt.UserRole) or did_item.text()
            dlg = DeviceDetailDialog(did, self)
            dlg.exec_()

    # ========== OVERVIEW TAB ==========
    def _create_overview_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Export button at top
        export_row = QHBoxLayout()
        export_row.addStretch()
        self.ov_export_btn = QPushButton("📥 Esporta Overview Excel")
        self.ov_export_btn.setObjectName("success")
        self.ov_export_btn.setStyleSheet("background: #2E7D32; color: white; padding: 8px 16px; font-weight: bold;")
        self.ov_export_btn.clicked.connect(self.export_overview)
        export_row.addWidget(self.ov_export_btn)
        layout.addLayout(export_row)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        cl = QVBoxLayout(content)

        # Fornitore table
        self.ov_forn_table = QTableWidget()
        self.ov_forn_table.setColumnCount(8)
        self.ov_forn_table.setHorizontalHeaderLabels(["Fornitore", "Totale", "OK", "KO", "Degraded", "% OK", "Ticket Aperti", "Sotto Corona"])
        self.ov_forn_table.horizontalHeader().setStretchLastSection(True)
        grp_f = QGroupBox("Stato per Fornitore")
        fl = QVBoxLayout()
        fl.addWidget(self.ov_forn_table)
        grp_f.setLayout(fl)
        cl.addWidget(grp_f)

        # DT table
        self.ov_dt_table = QTableWidget()
        self.ov_dt_table.setColumnCount(5)
        self.ov_dt_table.setHorizontalHeaderLabels(["DT", "Totale", "OK", "KO", "% OK"])
        self.ov_dt_table.horizontalHeader().setStretchLastSection(True)
        grp_d = QGroupBox("Stato per Direzione Territoriale")
        dl = QVBoxLayout()
        dl.addWidget(self.ov_dt_table)
        grp_d.setLayout(dl)
        cl.addWidget(grp_d)

        # Correlation matrix
        self.ov_corr_table = QTableWidget()
        self.ov_corr_table.setColumnCount(8)
        self.ov_corr_table.setHorizontalHeaderLabels(["Fornitore", "Totale", "LTE KO", "SSH KO", "Mongo KO", "Porta KO", "Batt KO", "Disconnessi"])
        self.ov_corr_table.horizontalHeader().setStretchLastSection(True)
        grp_c = QGroupBox("Matrice Correlazione Diagnostica")
        cll = QVBoxLayout()
        cll.addWidget(self.ov_corr_table)
        grp_c.setLayout(cll)
        cl.addWidget(grp_c)

        cl.addStretch()
        scroll.setWidget(content)
        layout.addWidget(scroll)
        return tab

    # ========== DATA LOADING ==========
    def refresh_data(self):
        session = get_session()
        try:
            total = session.query(Device).count()
            if total == 0:
                self.status_label.setText("Nessun dato. Importa un file Excel.")
                return

            health = dict(session.query(Device.current_health, func_count()).group_by(Device.current_health).all())
            tickets_open = session.query(Device).filter(Device.ticket_stato == "Aperto").count()
            crit = session.query(AnomalyEvent).filter(AnomalyEvent.severity == "CRITICAL", AnomalyEvent.acknowledged == False).count()
            high = session.query(AnomalyEvent).filter(AnomalyEvent.severity == "HIGH", AnomalyEvent.acknowledged == False).count()

            self._update_card(self.card_total, total)
            self._update_card(self.card_ok, health.get("OK", 0))
            self._update_card(self.card_ko, health.get("KO", 0))
            self._update_card(self.card_deg, health.get("DEGRADED", 0))
            self._update_card(self.card_crit, crit)
            self._update_card(self.card_high, high)
            self._update_card(self.card_tickets, tickets_open)

            # Populate alert type filter
            types = [r[0] for r in session.query(AnomalyEvent.event_type).distinct().all() if r[0]]
            self.alert_type.blockSignals(True)
            self.alert_type.clear()
            self.alert_type.addItem("Tutti")
            for t in sorted(types):
                self.alert_type.addItem(t)
            self.alert_type.blockSignals(False)

            self.refresh_alerts()
            self.refresh_devices()
            self.refresh_overview()
            self.status_label.setText(f"Dati caricati: {total} dispositivi")
        except Exception as e:
            self.status_label.setText(f"Errore: {e}")
        finally:
            session.close()

    def refresh_alerts(self):
        session = get_session()
        try:
            from sqlalchemy import func, and_
            q = (session.query(AnomalyEvent, Device)
                 .join(Device, AnomalyEvent.device_id == Device.device_id)
                 .filter(AnomalyEvent.acknowledged == False))

            sev = self.alert_sev.currentText()
            if sev != "Tutti": q = q.filter(AnomalyEvent.severity == sev)
            typ = self.alert_type.currentText()
            if typ != "Tutti": q = q.filter(AnomalyEvent.event_type == typ)
            forn = self.alert_forn.currentText()
            if forn != "Tutti": q = q.filter(Device.fornitore == forn)

            results = q.all()

            data = []
            for event, device in results:
                short_did = ":".join(device.device_id.split(":")[-2:]) if ":" in device.device_id else device.device_id
                data.append({
                    "Severity": event.severity,
                    "Tipo": (event.event_type or "").replace("_", " "),
                    "DeviceID": short_did,
                    "_full_did": device.device_id,
                    "Fornitore": device.fornitore or "-",
                    "DT": device.dt or "-",
                    "Trend": trend_str(device.trend_7d),
                    "LTE": device.check_lte or "-",
                    "SSH": device.check_ssh or "-",
                    "Mongo": device.check_mongo or "-",
                    "Batt": device.batteria or "-",
                    "Porta": device.porta_aperta or "-",
                    "Sotto C.": "SC" if device.is_sotto_corona else "",
                    "Descrizione": event.description or "",
                    "Ticket": f"{device.ticket_id} ({device.ticket_stato})" if device.ticket_id else "-",
                })

            def render_alert(row, col):
                val = row.get(col, "")
                if col == "Severity":
                    return colored_item(val, SEV_BG.get(val, ""), SEV_COLORS.get(val, ""), True)
                elif col == "DeviceID":
                    item = QTableWidgetItem(val)
                    item.setData(Qt.UserRole, row.get("_full_did"))
                    item.setForeground(QColor("#0066CC"))
                    font = item.font(); font.setBold(True); item.setFont(font)
                    return item
                elif col in ("LTE", "SSH", "Mongo", "Batt", "Porta"):
                    return check_item(val)
                elif col == "Sotto C." and val == "SC":
                    return colored_item("SC", "#E3F2FD", "#1565C0", True)
                elif col == "Trend":
                    item = QTableWidgetItem(val)
                    item.setFont(QFont("Consolas", 10))
                    return item
                elif col == "Ticket" and val != "-":
                    return colored_item(val, "#FFEBEE", "#C62828")
                return QTableWidgetItem(str(val))

            self.alert_table.set_data(data, render_alert)
            self.alert_count_label.setText(f"{len(data)} alert")
        finally:
            session.close()

    def refresh_devices(self):
        session = get_session()
        try:
            q = session.query(Device)
            forn = self.dev_forn.currentText()
            if forn != "Tutti": q = q.filter(Device.fornitore == forn)
            hlth = self.dev_health.currentText()
            if hlth != "Tutti": q = q.filter(Device.current_health == hlth)
            tipo = self.dev_tipo.currentText()
            if tipo != "Tutti":
                q = q.filter(Device.sistema_digil.contains(tipo))
            inst = self.dev_install.currentText()
            if inst == "Completa": q = q.filter(Device.is_sotto_corona == False)
            elif inst == "Sotto corona": q = q.filter(Device.is_sotto_corona == True)

            devices = q.all()

            data = []
            for d in devices:
                short_did = ":".join(d.device_id.split(":")[-2:]) if ":" in d.device_id else d.device_id
                sd = d.sistema_digil or ""
                tipo_short = "M" if "master" in sd else "S" if "slave" in sd else "?"
                data.append({
                    "DeviceID": short_did, "_full_did": d.device_id,
                    "Linea": d.linea or "-", "Fornitore": d.fornitore or "-",
                    "Tipo": tipo_short, "DT": d.dt or "-",
                    "Health": d.current_health or "-",
                    "LTE": d.check_lte or "-", "SSH": d.check_ssh or "-",
                    "Mongo": d.check_mongo or "-", "Batt": d.batteria or "-",
                    "Porta": d.porta_aperta or "-",
                    "Sotto C.": "SC" if d.is_sotto_corona else "",
                    "Trend": trend_str(d.trend_7d),
                    "Giorni": str(d.days_in_current) if d.days_in_current else "-",
                    "Malf.": d.tipo_malfunzionamento or "-",
                    "Ticket": d.ticket_id or "-",
                })

            def render_dev(row, col):
                val = row.get(col, "")
                if col == "DeviceID":
                    item = QTableWidgetItem(val)
                    item.setData(Qt.UserRole, row.get("_full_did"))
                    item.setForeground(QColor("#0066CC"))
                    font = item.font(); font.setBold(True); item.setFont(font)
                    return item
                elif col == "Health":
                    return colored_item(val, HEALTH_BG.get(val, ""), bold=True)
                elif col in ("LTE", "SSH", "Mongo", "Batt", "Porta"):
                    return check_item(val)
                elif col == "Sotto C." and val == "SC":
                    return colored_item("SC", "#E3F2FD", "#1565C0", True)
                elif col == "Trend":
                    item = QTableWidgetItem(val)
                    item.setFont(QFont("Consolas", 10))
                    return item
                elif col == "Ticket" and val != "-":
                    return colored_item(val, "#FFEBEE", "#C62828")
                return QTableWidgetItem(str(val))

            self.dev_table.set_data(data, render_dev)
            self.dev_count_label.setText(f"{len(data)} dispositivi")
        finally:
            session.close()

    def refresh_overview(self):
        session = get_session()
        try:
            # Fornitore stats
            fornitori = ["INDRA", "MII", "SIRTI"]
            self.ov_forn_table.setRowCount(len(fornitori))
            for i, f in enumerate(fornitori):
                devs = session.query(Device).filter(Device.fornitore == f).all()
                total = len(devs)
                ok = sum(1 for d in devs if d.current_health == "OK")
                ko = sum(1 for d in devs if d.current_health == "KO")
                deg = sum(1 for d in devs if d.current_health == "DEGRADED")
                tickets = sum(1 for d in devs if d.ticket_stato == "Aperto")
                sc = sum(1 for d in devs if d.is_sotto_corona)
                pct = round(ok / total * 100, 1) if total else 0
                self.ov_forn_table.setItem(i, 0, colored_item(f, bold=True))
                self.ov_forn_table.setItem(i, 1, colored_item(total, bold=True))
                self.ov_forn_table.setItem(i, 2, colored_item(ok, "#E8F5E9", "#2E7D32", True))
                self.ov_forn_table.setItem(i, 3, colored_item(ko, "#FFEBEE", "#C62828", True))
                self.ov_forn_table.setItem(i, 4, colored_item(deg, "#FFF3E0", "#E65100"))
                self.ov_forn_table.setItem(i, 5, colored_item(f"{pct}%", bold=True))
                self.ov_forn_table.setItem(i, 6, colored_item(tickets))
                self.ov_forn_table.setItem(i, 7, colored_item(sc, "#E3F2FD"))
            self.ov_forn_table.resizeRowsToContents()

            # DT stats
            from sqlalchemy import func
            dts = session.query(Device.dt, func.count()).group_by(Device.dt).all()
            dts = [(dt, cnt) for dt, cnt in dts if dt]
            dts.sort(key=lambda x: x[1], reverse=True)
            self.ov_dt_table.setRowCount(len(dts))
            for i, (dt, total) in enumerate(dts):
                ok = session.query(Device).filter(Device.dt == dt, Device.current_health == "OK").count()
                ko = total - ok
                pct = round(ok / total * 100, 1) if total else 0
                self.ov_dt_table.setItem(i, 0, colored_item(dt, bold=True))
                self.ov_dt_table.setItem(i, 1, colored_item(total))
                self.ov_dt_table.setItem(i, 2, colored_item(ok, "#E8F5E9", "#2E7D32"))
                self.ov_dt_table.setItem(i, 3, colored_item(ko, "#FFEBEE", "#C62828"))
                self.ov_dt_table.setItem(i, 4, colored_item(f"{pct}%", bold=True))
            self.ov_dt_table.resizeRowsToContents()

            # Correlation
            self.ov_corr_table.setRowCount(len(fornitori))
            for i, f in enumerate(fornitori):
                devs = session.query(Device).filter(Device.fornitore == f).all()
                total = len(devs)
                lte_ko = sum(1 for d in devs if d.check_lte == "KO")
                ssh_ko = sum(1 for d in devs if d.check_ssh == "KO")
                mongo_ko = sum(1 for d in devs if d.check_mongo == "KO")
                porta_ko = sum(1 for d in devs if d.porta_aperta == "KO")
                batt_ko = sum(1 for d in devs if d.batteria == "KO")
                disconn = sum(1 for d in devs if d.check_lte == "KO" and d.check_ssh == "KO")
                self.ov_corr_table.setItem(i, 0, colored_item(f, bold=True))
                self.ov_corr_table.setItem(i, 1, colored_item(total))
                self.ov_corr_table.setItem(i, 2, colored_item(lte_ko, bold=True))
                self.ov_corr_table.setItem(i, 3, colored_item(ssh_ko, bold=True))
                self.ov_corr_table.setItem(i, 4, colored_item(mongo_ko, bold=True))
                self.ov_corr_table.setItem(i, 5, colored_item(porta_ko, bold=True))
                self.ov_corr_table.setItem(i, 6, colored_item(batt_ko, bold=True))
                self.ov_corr_table.setItem(i, 7, colored_item(disconn, "#FFEBEE", "#C62828", True))
            self.ov_corr_table.resizeRowsToContents()
        finally:
            session.close()

    def _on_tab_changed(self, index):
        if index == 0: self.refresh_alerts()
        elif index == 1: self.refresh_devices()
        elif index == 2: self.refresh_overview()

    # ========== IMPORT ==========
    def do_import(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleziona File Excel Monitoraggio", "",
                                                    "Excel Files (*.xlsx *.xls);;All Files (*)")
        if not file_path: return
        self.import_btn.setEnabled(False)
        self.status_label.setText("Importazione in corso...")
        self.import_thread = ImportThread(file_path)
        self.import_thread.finished.connect(self._on_import_done)
        self.import_thread.error.connect(self._on_import_error)
        self.import_thread.start()

    def _on_import_done(self, stats, alert_count):
        self.import_btn.setEnabled(True)
        self.refresh_data()
        QMessageBox.information(self, "Import Completato",
            f"Dispositivi: {stats['devices_imported']}\n"
            f"Record availability: {stats['availability_records']}\n"
            f"Alert generati: {alert_count}")

    def _on_import_error(self, error):
        self.import_btn.setEnabled(True)
        self.status_label.setText(f"Errore import: {error}")
        QMessageBox.critical(self, "Errore Import", error)

    def export_overview(self):
        """Esporta le tabelle Overview in un file Excel con 3 sheet"""
        import pandas as pd
        from pathlib import Path

        default_name = f"DIGIL_Overview_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Salva Overview", str(Path.home() / "Downloads" / default_name), "Excel Files (*.xlsx)")
        if not file_path:
            return

        session = get_session()
        try:
            # 1. Fornitore data
            forn_data = []
            for f in ["INDRA", "MII", "SIRTI"]:
                devs = session.query(Device).filter(Device.fornitore == f).all()
                total = len(devs)
                ok = sum(1 for d in devs if d.current_health == "OK")
                ko = sum(1 for d in devs if d.current_health == "KO")
                deg = sum(1 for d in devs if d.current_health == "DEGRADED")
                tickets = sum(1 for d in devs if d.ticket_stato == "Aperto")
                sc = sum(1 for d in devs if d.is_sotto_corona)
                forn_data.append({"Fornitore": f, "Totale": total, "OK": ok, "KO": ko,
                                  "Degraded": deg, "% OK": round(ok/total*100, 1) if total else 0,
                                  "Ticket Aperti": tickets, "Sotto Corona": sc})

            # 2. DT data
            from sqlalchemy import func
            dt_data = []
            dts = session.query(Device.dt, func.count()).group_by(Device.dt).all()
            for dt_name, total in sorted(dts, key=lambda x: x[1], reverse=True):
                if not dt_name: continue
                ok = session.query(Device).filter(Device.dt == dt_name, Device.current_health == "OK").count()
                ko = total - ok
                dt_data.append({"DT": dt_name, "Totale": total, "OK": ok, "KO": ko,
                                "% OK": round(ok/total*100, 1) if total else 0})

            # 3. Correlation data
            corr_data = []
            for f in ["INDRA", "MII", "SIRTI"]:
                devs = session.query(Device).filter(Device.fornitore == f).all()
                total = len(devs)
                corr_data.append({
                    "Fornitore": f, "Totale": total,
                    "LTE KO": sum(1 for d in devs if d.check_lte == "KO"),
                    "SSH KO": sum(1 for d in devs if d.check_ssh == "KO"),
                    "Mongo KO": sum(1 for d in devs if d.check_mongo == "KO"),
                    "Porta KO": sum(1 for d in devs if d.porta_aperta == "KO"),
                    "Batt KO": sum(1 for d in devs if d.batteria == "KO"),
                    "Disconnessi": sum(1 for d in devs if d.check_lte == "KO" and d.check_ssh == "KO"),
                })

            # Write Excel
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                pd.DataFrame(forn_data).to_excel(writer, index=False, sheet_name='Stato Fornitore')
                pd.DataFrame(dt_data).to_excel(writer, index=False, sheet_name='Stato DT')
                pd.DataFrame(corr_data).to_excel(writer, index=False, sheet_name='Correlazione Diagnostica')

                wb = writer.book
                hdr_fmt = wb.add_format({'bold': True, 'bg_color': '#0066CC', 'font_color': 'white', 'border': 1, 'align': 'center'})
                for sheet_name in ['Stato Fornitore', 'Stato DT', 'Correlazione Diagnostica']:
                    ws = writer.sheets[sheet_name]
                    ws.freeze_panes(1, 0)
                    ws.set_column(0, 0, 18)
                    ws.set_column(1, 10, 14)

            self.status_label.setText(f"Overview esportata: {file_path}")
            QMessageBox.information(self, "Export Completato", f"File salvato:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Errore Export", str(e))
        finally:
            session.close()


def func_count():
    from sqlalchemy import func
    return func.count()


def main():
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setApplicationName("DIGIL Monitoring")
    app.setOrganizationName("Terna")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
