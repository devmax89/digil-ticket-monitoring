"""
DIGIL Monitoring Dashboard - PyQt5 - Terna IoT Team
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
    QDialog, QTextEdit, QScrollArea, QSplitter, QSizePolicy,
    QFormLayout, QDateEdit, QDialogButtonBox, QCheckBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDate
from PyQt5.QtGui import QColor, QFont, QBrush, QPixmap
from database import get_session, init_db, Device, AvailabilityDaily, AnomalyEvent, ImportLog, TicketHistory
from importer import run_import
from detection import run_detection

JIRA_USERS = {
    "Festa Rosa": "60705508126db9006f3be9e8",
    "Massimiliano Tavernese": "5e8ae84a84dec20b8159e37a",
    "Paolo Marino": "712020:18a68569-e00b-414c-bd78-2ef0e43c0534",
    "Vittorio Mitri": "5e86f312b39dbf0c114bdefa",
    "Team AMS": "622f434533fb840069656a1a",
}
JIRA_USER_NAMES = list(JIRA_USERS.keys())

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
QPushButton#success { background: #2E7D32; }
QPushButton#jira { background: #0052CC; color: white; }
QPushButton#jira:hover { background: #0747A6; }
QPushButton#toggle_on { background: #E65100; color: white; border: 2px solid #BF360C; font-weight: bold; }
QPushButton#toggle_off { background: white; color: #666; border: 1px solid #CCC; font-weight: normal; }
QTableWidget { border: 1px solid #CCCCCC; gridline-color: #E0E0E0; background: white; alternate-background-color: #F8FBFF; font-size: 11px; }
QTableWidget::item { padding: 3px 5px; }
QTableWidget::item:selected { background: #CCE5FF; color: black; }
QHeaderView::section { background: #0066CC; color: white; padding: 6px; border: none; font-weight: bold; font-size: 11px; }
QTabWidget::pane { border: 1px solid #CCCCCC; background: white; }
QTabBar::tab { background: #F0F0F0; border: 1px solid #CCCCCC; padding: 12px 30px; font-weight: bold; font-size: 13px; min-width: 100px; }
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
AVAIL_COLORS = {"COMPLETE": "#2E7D32", "AVAILABLE": "#66BB6A", "NOT AVAILABLE": "#F9A825", "NO DATA": "#C62828"}

JIRA_LABELS = ["Misure_fuori_range", "Porta_aperta", "Misure_assenti", "Disconnesso", "Comandi", "Misure_parziali", "Allarme_batteria", "DigilApp", "PiattaformaIoT"]

def avail_color(raw_status):
    """Map any raw availability status to its color. Handles CODE_3, CODE_4, etc."""
    if not raw_status: return "#BDBDBD"
    v = raw_status.strip().upper()
    if v in ("COMPLETE", "1", "CODE_1"): return "#2E7D32"
    if v in ("AVAILABLE", "ON", "2", "CODE_2"): return "#66BB6A"
    if v in ("NOT AVAILABLE", "OFF", "KO", "3", "CODE_3"): return "#F9A825"
    if v in ("NO DATA", "4", "CODE_4"): return "#C62828"
    return AVAIL_COLORS.get(v, "#BDBDBD")

def colored_item(text, bg=None, fg=None, bold=False):
    item = QTableWidgetItem(str(text) if text else "-"); item.setTextAlignment(Qt.AlignCenter)
    if bg: item.setBackground(QColor(bg))
    if fg: item.setForeground(QColor(fg))
    if bold: item.setFont(QFont("Segoe UI", 11, QFont.Bold))
    return item

def check_item(val):
    if val == "OK": return colored_item("OK", "#E8F5E9", "#2E7D32", True)
    elif val == "KO": return colored_item("KO", "#FFEBEE", "#C62828", True)
    return colored_item(val or "-", "#F5F5F5", "#757575")

def trend_str(t):
    if not t: return "-"
    return "".join("\u25a0" if c == "O" else "\u25a1" for c in t)

def func_count():
    from sqlalchemy import func
    return func.count()

class ImportThread(QThread):
    progress = pyqtSignal(str); finished = pyqtSignal(dict, int); error = pyqtSignal(str)
    def __init__(self, fp): super().__init__(); self.file_path = fp
    def run(self):
        try:
            self.progress.emit("Importazione..."); stats = run_import(self.file_path)
            self.progress.emit("Alert..."); ac = run_detection(); self.finished.emit(stats, ac)
        except Exception as e: self.error.emit(str(e))

class FilterableTable(QWidget):
    def __init__(self, columns, parent=None):
        super().__init__(parent); self.columns = columns; self._all = []; self._filt = []; self._rfn = None
        lo = QVBoxLayout(self); lo.setContentsMargins(0,0,0,0); lo.setSpacing(2)
        fw = QWidget(); fl = QHBoxLayout(fw); fl.setContentsMargins(2,2,2,2); fl.setSpacing(2)
        self.filters = {}
        for col in columns:
            le = QLineEdit(); le.setPlaceholderText(f"Filtra {col}"); le.setMaximumHeight(22)
            le.setStyleSheet("font-size:10px;padding:1px 4px;"); le.textChanged.connect(self._apply); self.filters[col] = le; fl.addWidget(le)
        lo.addWidget(fw)
        self.table = QTableWidget(); self.table.setColumnCount(len(columns)); self.table.setHorizontalHeaderLabels(columns)
        self.table.setAlternatingRowColors(True); self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection); self.table.setSortingEnabled(True)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers); self.table.verticalHeader().setDefaultSectionSize(24)
        self.table.horizontalHeader().setStretchLastSection(True); lo.addWidget(self.table)
    def clear_filters(self):
        for le in self.filters.values(): le.blockSignals(True); le.clear(); le.blockSignals(False)
        self._apply()
    def set_data(self, data, rfn=None): self._all = data; self._rfn = rfn; self._apply()
    def _apply(self):
        act = {c: le.text().strip().lower() for c, le in self.filters.items() if le.text().strip()}
        self._filt = [r for r in self._all if all(t in str(r.get(c,"")).lower() for c,t in act.items())] if act else list(self._all)
        self.table.setSortingEnabled(False); self.table.setRowCount(len(self._filt))
        for ri, rd in enumerate(self._filt):
            for ci, cn in enumerate(self.columns):
                it = self._rfn(rd, cn) if self._rfn else QTableWidgetItem(str(rd.get(cn,"")))
                if it: self.table.setItem(ri, ci, it)
        self.table.setSortingEnabled(True)
    def get_selected_rows_data(self):
        rows = sorted(set(idx.row() for idx in self.table.selectedIndexes()))
        return [self._filt[r] for r in rows if r < len(self._filt)]

class JiraTicketDialog(QDialog):
    def __init__(self, tickets_data, parent=None):
        super().__init__(parent); n = len(tickets_data)
        self.setWindowTitle(f"Genera Ticket Jira - {n} dispositivi"); self.setMinimumSize(1050, 750); self.setStyleSheet(STYLE)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"<b style='color:#0052CC;font-size:14px'>Preparazione {n} ticket Jira</b>"))
        form = QFormLayout()
        self.issue_type_edit = QLineEdit("Bug in esercizio"); form.addRow("Issue Type:", self.issue_type_edit)
        self.priority_combo = QComboBox(); self.priority_combo.addItems(["Highest","High","Medium","Low","Lowest"]); self.priority_combo.setCurrentText("Medium"); form.addRow("Priorita:", self.priority_combo)
        self.assignee_combo = QComboBox(); self.assignee_combo.setEditable(True)
        for name in JIRA_USER_NAMES: self.assignee_combo.addItem(name, JIRA_USERS[name])
        self.assignee_combo.setCurrentIndex(1); form.addRow("Assignee:", self.assignee_combo)
        self.reporter_combo = QComboBox(); self.reporter_combo.setEditable(True); self.reporter_combo.addItem("(nessuno)", "")
        for name in JIRA_USER_NAMES: self.reporter_combo.addItem(name, JIRA_USERS[name])
        form.addRow("Reporter:", self.reporter_combo)
        wg = QWidget(); wl = QHBoxLayout(wg); wl.setContentsMargins(0,0,0,0); self.watcher_cbs = {}
        for name in JIRA_USER_NAMES: cb = QCheckBox(name); self.watcher_cbs[name] = cb; wl.addWidget(cb)
        form.addRow("Watchers:", wg)
        lg = QWidget(); lgl = QHBoxLayout(lg); lgl.setContentsMargins(0,0,0,0); lgl.setSpacing(4); self.label_cbs = {}
        for lname in JIRA_LABELS: cb = QCheckBox(lname); self.label_cbs[lname] = cb; lgl.addWidget(cb)
        lscroll = QScrollArea(); lscroll.setWidgetResizable(True); lscroll.setMaximumHeight(36); lscroll.setWidget(lg); form.addRow("Labels:", lscroll)
        self.due_date_edit = QDateEdit(); self.due_date_edit.setCalendarPopup(True); self.due_date_edit.setDate(QDate.currentDate().addDays(14)); self.due_date_edit.setDisplayFormat("dd/MM/yyyy"); form.addRow("Due Date:", self.due_date_edit)
        self.reply_date = QLineEdit(datetime.now().strftime("%d-%m-%y")); form.addRow("Data Reply:", self.reply_date)
        layout.addLayout(form)
        cols = ["DeviceID","Summary","Label1","Label2","Description"]
        self.tt = QTableWidget(); self.tt.setColumnCount(5); self.tt.setHorizontalHeaderLabels(cols)
        self.tt.setRowCount(n); self.tt.horizontalHeader().setStretchLastSection(True); self.tt.setAlternatingRowColors(True)
        self.tt.setColumnWidth(0,180); self.tt.setColumnWidth(1,250); self.tt.setColumnWidth(2,100); self.tt.setColumnWidth(3,100)
        for i, td in enumerate(tickets_data):
            did = td.get("_full_did", td.get("DeviceID",""))
            tm = td.get("tipo_malf","") or ""; tm = "" if tm=="-" else tm
            l1 = td.get("tipo_malf_jira","") or ""; l1 = "" if l1=="-" else l1
            l2 = td.get("cluster_jira","") or ""; l2 = "" if l2=="-" else l2
            lte=td.get("lte",td.get("LTE","-")); ssh=td.get("ssh",td.get("SSH","-"))
            batt=td.get("batteria",td.get("Batt","-")); porta=td.get("porta",td.get("Porta","-"))
            mongo=td.get("mongo",td.get("Mongo","-")); note=td.get("note","") or ""; note="" if note=="-" else note
            desc = f"Reply {self.reply_date.text()}: Valori recuperati: Check LTE:{lte}, check SSH:{ssh}, Batteria:{batt}, Porta aperta:{porta}, Check Mongo:{mongo}"
            if note: desc += f"\r\n{note}"
            self.tt.setItem(i,0,QTableWidgetItem(did)); self.tt.setItem(i,1,QTableWidgetItem(f"Device {did} {tm}".strip()))
            self.tt.setItem(i,2,QTableWidgetItem(l1)); self.tt.setItem(i,3,QTableWidgetItem(l2)); self.tt.setItem(i,4,QTableWidgetItem(desc))
        layout.addWidget(QLabel("<i>Puoi modificare qualsiasi cella prima dell'export</i>")); layout.addWidget(self.tt)
        br = QHBoxLayout()
        cpb = QPushButton("Copia tutto"); cpb.setObjectName("secondary"); cpb.clicked.connect(self._copy); br.addWidget(cpb); br.addStretch()
        xb = QPushButton("Annulla"); xb.setObjectName("secondary"); xb.clicked.connect(self.reject); br.addWidget(xb)
        eb = QPushButton(f"Esporta CSV Jira ({n})"); eb.setObjectName("jira"); eb.clicked.connect(self._export); br.addWidget(eb); layout.addLayout(br)

    def _rows(self):
        return [{c: (self.tt.item(i,j).text() if self.tt.item(i,j) else "") for j,c in enumerate(["DeviceID","Summary","Label1","Label2","Description"])} for i in range(self.tt.rowCount())]
    def _copy(self):
        QApplication.clipboard().setText("\n---\n".join(f"DeviceID: {r['DeviceID']}\nSummary: {r['Summary']}\nLabels: {r['Label1']} {r['Label2']}\nDescription: {r['Description']}" for r in self._rows()))
        QMessageBox.information(self, "Copiato", f"{self.tt.rowCount()} ticket copiati!")
    def _export(self):
        n = self.tt.rowCount()
        fp, _ = QFileDialog.getSaveFileName(self, "Salva CSV Jira", str(Path.home()/"Downloads"/f"ticket_jira_{n}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"), "CSV (*.csv)")
        if not fp: return
        it = self.issue_type_edit.text(); pr = self.priority_combo.currentText()
        aidx = self.assignee_combo.currentIndex(); aid = self.assignee_combo.itemData(aidx) if aidx >= 0 and self.assignee_combo.itemData(aidx) else self.assignee_combo.currentText()
        ridx = self.reporter_combo.currentIndex(); rid = self.reporter_combo.itemData(ridx) if ridx >= 0 and self.reporter_combo.itemData(ridx) else ""
        wids = [JIRA_USERS[nm] for nm, cb in self.watcher_cbs.items() if cb.isChecked()]
        dd = self.due_date_edit.date().toString("dd/MM/yyyy") + " 00:00"
        el = [nm for nm, cb in self.label_cbs.items() if cb.isChecked()]
        ml = 2 + len(el); mw = max(len(wids), 1)
        hdr = ["Summary","Issue Type","Priority","Assignee"]
        if rid: hdr.append("Reporter")
        hdr.append("Due date")
        hdr.extend(["Labels"] * ml); hdr.extend(["Watchers"] * mw); hdr.append("Description")
        with open(fp, 'w', encoding='utf-8-sig') as f:
            f.write(";".join(hdr) + "\n")
            for r in self._rows():
                s = r["Summary"].replace('"','""'); d = r["Description"].replace('"','""')
                labs = [r["Label1"], r["Label2"]] + el
                while len(labs) < ml: labs.append("")
                ws = list(wids)
                while len(ws) < mw: ws.append("")
                parts = [f'"{s}"', it, pr, aid]
                if rid: parts.append(rid)
                parts.append(dd); parts.extend(labs); parts.extend(ws); parts.append(f'"{d}"')
                f.write(";".join(parts) + "\n")
        QMessageBox.information(self, "Export CSV Jira", f"{n} ticket salvati:\n{fp}"); self.accept()

class DeviceDetailDialog(QDialog):
    def __init__(self, device_id, parent=None):
        super().__init__(parent); self.setWindowTitle(f"Dettaglio: {device_id}"); self.setMinimumSize(780, 720); self.setStyleSheet(STYLE)
        session = get_session()
        try:
            device = session.get(Device, device_id)
            if not device: QMessageBox.warning(self, "Errore", f"Device {device_id} non trovato"); return
            layout = QVBoxLayout(self)
            top = QHBoxLayout()
            tl = QLabel(device_id); tl.setStyleSheet("font-size:16px;font-weight:bold;color:#0066CC;"); tl.setTextInteractionFlags(Qt.TextSelectableByMouse); top.addWidget(tl)
            hl = QLabel(device.current_health or "?"); hbg = HEALTH_BG.get(device.current_health,"#F5F5F5")
            hl.setStyleSheet(f"background:{hbg};padding:4px 12px;border-radius:3px;font-weight:bold;"); top.addWidget(hl)
            if device.is_sotto_corona:
                sc = QLabel("SOTTO CORONA"); sc.setStyleSheet("background:#E3F2FD;color:#1565C0;padding:4px 8px;border-radius:3px;font-weight:bold;"); top.addWidget(sc)
            top.addStretch()
            cpb = QPushButton("Copia Info"); cpb.setObjectName("secondary"); cpb.clicked.connect(lambda: self._copy_info(device, session)); top.addWidget(cpb)
            jb = QPushButton("Ticket Jira"); jb.setObjectName("jira"); jb.clicked.connect(lambda: self._open_jira(device)); top.addWidget(jb)
            layout.addLayout(top)
            scroll = QScrollArea(); scroll.setWidgetResizable(True); content = QWidget(); cl = QVBoxLayout(content)
            grid = QGridLayout(); grid.setSpacing(4); r = 0
            for lbl, val in [("DeviceID",device.device_id),("Linea",device.linea),("Sostegno",device.st_sostegno),("Fornitore",device.fornitore),("Tipo",device.sistema_digil),("DT",device.dt),("Denominazione",device.denominazione),("Regione",f"{device.regione or ''} / {device.provincia or ''}"),("IP",device.ip_address),("Installazione",str(device.data_install) if device.data_install else "-"),("Tipo Install.",device.tipo_install),("Da file master",device.da_file_master)]:
                if val:
                    lb = QLabel(f"<b style='color:#666'>{lbl}:</b>"); lb.setTextInteractionFlags(Qt.TextSelectableByMouse)
                    vl = QLabel(str(val)); vl.setTextInteractionFlags(Qt.TextSelectableByMouse); grid.addWidget(lb, r, 0); grid.addWidget(vl, r, 1); r += 1
            ga = QGroupBox("Anagrafica"); ga.setLayout(grid); cl.addWidget(ga)
            dl = QHBoxLayout()
            for name, val in [("LTE",device.check_lte),("SSH",device.check_ssh),("Mongo",device.check_mongo),("Batteria",device.batteria),("Porta",device.porta_aperta)]:
                bx = QLabel(f"<center><small>{name}</small><br><b>{val or '-'}</b></center>")
                bg = "#E8F5E9" if val=="OK" else "#FFEBEE" if val=="KO" else "#F5F5F5"
                bx.setStyleSheet(f"background:{bg};border:1px solid #CCC;border-radius:4px;padding:6px;min-width:60px;"); dl.addWidget(bx)
            dl.addWidget(QLabel(f"<b>Trend 7d:</b> {trend_str(device.trend_7d)}")); dl.addWidget(QLabel(f"<b>Giorni:</b> {device.days_in_current}")); dl.addStretch()
            gd = QGroupBox("Diagnostica"); gd.setLayout(dl); cl.addWidget(gd)
            if device.last_avail_status:
                av_bg = avail_color(device.last_avail_status)
                cl.addWidget(QLabel(f"<b>Ultimo availability:</b> {device.last_avail_status} ({device.last_avail_date})").setStyleSheet(f"padding:4px 8px;border-left:4px solid {av_bg};background:#FAFAFA;") or QLabel(f"<b>Ultimo availability:</b> {device.last_avail_status} ({device.last_avail_date})"))
            malf = [("Tipo",device.tipo_malfunzionamento),("Cluster",device.cluster_analisi),("Analisi",device.analisi_malfunzionamento),("Intervento",device.tipologia_intervento),("Strategia",device.strategia_risolutiva),("Cause",device.cause_anomalie),("Risoluzione",device.risoluzione_attuata),("Note",device.note)]
            if any(v for _,v in malf):
                mg = QGridLayout(); mg.setSpacing(4); mr = 0
                for lbl, val in malf:
                    if val:
                        mg.addWidget(QLabel(f"<b style='color:#666'>{lbl}:</b>"), mr, 0, Qt.AlignTop)
                        vl = QLabel(str(val)); vl.setWordWrap(True); vl.setTextInteractionFlags(Qt.TextSelectableByMouse); mg.addWidget(vl, mr, 1); mr += 1
                gm = QGroupBox("Malfunzionamento"); gm.setLayout(mg); cl.addWidget(gm)
            if device.ticket_id:
                tgl = QGridLayout()
                tgl.addWidget(QLabel("<b style='color:#666'>ID:</b>"), 0, 0); tid = QLabel(f"<b>{device.ticket_id}</b>"); tid.setTextInteractionFlags(Qt.TextSelectableByMouse); tgl.addWidget(tid, 0, 1)
                tgl.addWidget(QLabel("<b style='color:#666'>Stato:</b>"), 1, 0)
                sbg = "#FFEBEE" if device.ticket_stato=="Aperto" else "#E8F5E9" if device.ticket_stato in ("Chiuso","Risolto") else "#FFF3E0"
                sll = QLabel(device.ticket_stato or "-"); sll.setStyleSheet(f"background:{sbg};padding:2px 8px;border-radius:3px;"); tgl.addWidget(sll, 1, 1)
                if device.ticket_data_apertura: tgl.addWidget(QLabel("<b style='color:#666'>Apertura:</b>"), 2, 0); tgl.addWidget(QLabel(str(device.ticket_data_apertura)), 2, 1)
                if device.ticket_data_risoluzione: tgl.addWidget(QLabel("<b style='color:#666'>Risoluzione:</b>"), 3, 0); tgl.addWidget(QLabel(str(device.ticket_data_risoluzione)), 3, 1)
                gt = QGroupBox("Ticket Corrente"); gt.setLayout(tgl); cl.addWidget(gt)
            thist = session.query(TicketHistory).filter(TicketHistory.device_id==device_id).order_by(TicketHistory.first_seen.desc()).all()
            if thist:
                ht = QTableWidget(); hc = ["Ticket","Stato","Apertura","Risoluzione","Tipo Malf.","Cluster","Note","Prima volta","Ultima volta"]
                ht.setColumnCount(9); ht.setHorizontalHeaderLabels(hc); ht.setRowCount(len(thist)); ht.horizontalHeader().setStretchLastSection(True); ht.setAlternatingRowColors(True); ht.setEditTriggers(QAbstractItemView.NoEditTriggers)
                for i, th in enumerate(thist):
                    ht.setItem(i,0,QTableWidgetItem(th.ticket_id or "-"))
                    st = th.ticket_stato or "-"; sb = "#FFEBEE" if st=="Aperto" else "#E8F5E9" if st in ("Chiuso","Risolto") else "#FFF3E0" if st=="Interno" else "#F5F5F5"
                    ht.setItem(i,1,colored_item(st, sb, bold=True))
                    ht.setItem(i,2,QTableWidgetItem(str(th.ticket_data_apertura) if th.ticket_data_apertura else "-"))
                    ht.setItem(i,3,QTableWidgetItem(str(th.ticket_data_risoluzione) if th.ticket_data_risoluzione else "-"))
                    ht.setItem(i,4,QTableWidgetItem(th.tipo_malfunzionamento or "-")); ht.setItem(i,5,QTableWidgetItem(th.cluster_analisi or "-"))
                    ht.setItem(i,6,QTableWidgetItem((th.note or "-")[:80]))
                    ht.setItem(i,7,QTableWidgetItem(th.first_seen.strftime("%Y-%m-%d %H:%M") if th.first_seen else "-"))
                    ht.setItem(i,8,QTableWidgetItem(th.last_seen.strftime("%Y-%m-%d %H:%M") if th.last_seen else "-"))
                ht.resizeColumnsToContents(); gh = QGroupBox(f"Storico Ticket ({len(thist)})"); ghl = QVBoxLayout(); ghl.addWidget(ht); gh.setLayout(ghl); cl.addWidget(gh)
            avail = session.query(AvailabilityDaily).filter(AvailabilityDaily.device_id==device_id).order_by(AvailabilityDaily.check_date).all()
            if avail:
                tll = QHBoxLayout()
                for a in avail:
                    bg = avail_color(a.raw_status)
                    bx = QLabel(""); bx.setFixedSize(14,14); bx.setStyleSheet(f"background:{bg};border-radius:2px;"); bx.setToolTip(f"{a.check_date}: {a.raw_status}"); tll.addWidget(bx)
                tll.addStretch(); leg = QHBoxLayout()
                for lbl, clr in [("COMPLETE","#2E7D32"),("AVAILABLE","#66BB6A"),("NOT AVAILABLE","#F9A825"),("NO DATA","#C62828")]:
                    sq = QLabel(""); sq.setFixedSize(10,10); sq.setStyleSheet(f"background:{clr};border-radius:1px;"); leg.addWidget(sq); leg.addWidget(QLabel(f"<small>{lbl}</small>")); leg.addSpacing(8)
                leg.addStretch(); tlo = QVBoxLayout(); tlo.addLayout(tll); tlo.addLayout(leg)
                gtl = QGroupBox(f"Timeline Availability ({len(avail)} giorni)"); gtl.setLayout(tlo); cl.addWidget(gtl)
            events = session.query(AnomalyEvent).filter(AnomalyEvent.device_id==device_id).order_by(AnomalyEvent.created_at.desc()).limit(10).all()
            if events:
                et = QTableWidget(); et.setColumnCount(4); et.setHorizontalHeaderLabels(["Data","Tipo","Severity","Descrizione"]); et.setRowCount(len(events)); et.horizontalHeader().setStretchLastSection(True)
                for i, e in enumerate(events):
                    et.setItem(i,0,QTableWidgetItem(str(e.event_date))); et.setItem(i,1,QTableWidgetItem(e.event_type))
                    et.setItem(i,2,colored_item(e.severity, SEV_BG.get(e.severity), SEV_COLORS.get(e.severity), True)); et.setItem(i,3,QTableWidgetItem(e.description or ""))
                ge = QGroupBox("Alert Recenti"); el = QVBoxLayout(); el.addWidget(et); ge.setLayout(el); cl.addWidget(ge)
            cl.addStretch(); scroll.setWidget(content); layout.addWidget(scroll)
            cb = QPushButton("Chiudi"); cb.clicked.connect(self.close); layout.addWidget(cb, alignment=Qt.AlignRight)
        finally: session.close()

    def _copy_info(self, device, session):
        lines = [f"DeviceID: {device.device_id}",f"Fornitore: {device.fornitore}",f"Linea: {device.linea}",f"Sostegno: {device.st_sostegno}",f"DT: {device.dt}",f"Denominazione: {device.denominazione}",f"Regione: {device.regione} / {device.provincia}",f"IP: {device.ip_address}",f"Tipo: {device.sistema_digil}",f"Tipo Install: {device.tipo_install}",f"Sotto Corona: {'Si' if device.is_sotto_corona else 'No'}",f"Data Install: {device.data_install}","","--- Diagnostica ---",f"LTE: {device.check_lte}",f"SSH: {device.check_ssh}",f"Mongo: {device.check_mongo}",f"Batteria: {device.batteria}",f"Porta: {device.porta_aperta}",f"Health: {device.current_health}",f"Ultimo Avail: {device.last_avail_status} ({device.last_avail_date})",f"Trend 7d: {trend_str(device.trend_7d)}",f"Giorni: {device.days_in_current}","","--- Ticket ---",f"Ticket: {device.ticket_id or 'Nessuno'}",f"Stato: {device.ticket_stato or '-'}",f"Apertura: {device.ticket_data_apertura or '-'}",f"Risoluzione: {device.ticket_data_risoluzione or '-'}"]
        if device.tipo_malfunzionamento: lines += ["","--- Malfunzionamento ---",f"Tipo: {device.tipo_malfunzionamento}",f"Cluster: {device.cluster_analisi or '-'}",f"Analisi: {device.analisi_malfunzionamento or '-'}",f"Cause: {device.cause_anomalie or '-'}",f"Note: {device.note or '-'}"]
        hist = session.query(TicketHistory).filter(TicketHistory.device_id==device.device_id).order_by(TicketHistory.first_seen.desc()).all()
        if hist:
            lines += ["",f"--- Storico Ticket ({len(hist)}) ---"]
            for h in hist: lines.append(f"  {h.ticket_id} [{h.ticket_stato}] Apt:{h.ticket_data_apertura or '-'} Ris:{h.ticket_data_risoluzione or '-'} Tipo:{h.tipo_malfunzionamento or '-'}")
        QApplication.clipboard().setText("\n".join(lines)); QMessageBox.information(self, "Copiato", "Info copiate!")

    def _open_jira(self, device):
        JiraTicketDialog([{"_full_did":device.device_id,"DeviceID":device.device_id,"lte":device.check_lte,"ssh":device.check_ssh,"mongo":device.check_mongo,"batteria":device.batteria,"porta":device.porta_aperta,"tipo_malf":device.tipo_malfunzionamento,"tipo_malf_jira":device.tipo_malf_jira,"cluster_jira":device.cluster_jira,"note":device.note}], self).exec_()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__(); self.import_thread = None; self._alert_no_ticket = False; self._dev_no_ticket = False
        self.init_ui(); self.setStyleSheet(STYLE); init_db(); self.refresh_data()

    def init_ui(self):
        self.setWindowTitle("DIGIL Monitoring Dashboard - Terna IoT Team"); self.setMinimumSize(1400, 900)
        central = QWidget(); self.setCentralWidget(central); ml = QVBoxLayout(central); ml.setSpacing(8); ml.setContentsMargins(12,12,12,12)
        hdr = QHBoxLayout()
        self.logo_label = QLabel("TERNA"); self.logo_label.setStyleSheet("background:#0066CC;color:white;font-size:18px;font-weight:bold;padding:8px 16px;border-radius:5px;"); self.logo_label.setFixedHeight(40); self.logo_label.setFixedWidth(120); self.logo_label.setAlignment(Qt.AlignCenter)
        lp = Path(__file__).parent/"assets"/"logo_terna.png"
        if lp.exists():
            px = QPixmap(str(lp))
            if not px.isNull(): self.logo_label.setPixmap(px.scaled(120,40,Qt.KeepAspectRatio,Qt.SmoothTransformation)); self.logo_label.setText(""); self.logo_label.setStyleSheet("background:transparent;")
        hdr.addWidget(self.logo_label)
        tw = QWidget(); tvl = QVBoxLayout(tw); tvl.setContentsMargins(10,0,0,0); tvl.setSpacing(0)
        tvl.addWidget(QLabel("DIGIL Monitoring Dashboard")); tvl.addWidget(QLabel("Monitoraggio IoT - Terna S.p.A."))
        hdr.addWidget(tw); hdr.addStretch()
        self.import_btn = QPushButton("Importa Excel"); self.import_btn.setObjectName("secondary"); self.import_btn.clicked.connect(self.do_import); hdr.addWidget(self.import_btn)
        ml.addLayout(hdr)
        sep = QFrame(); sep.setFrameShape(QFrame.HLine); sep.setStyleSheet("background:#CCCCCC;max-height:1px;"); ml.addWidget(sep)
        cl = QHBoxLayout()
        self.card_total = self._make_card("0","Dispositivi"); self.card_ok = self._make_card("0","OK","#E8F5E9"); self.card_ko = self._make_card("0","KO","#FFEBEE"); self.card_deg = self._make_card("0","Degraded","#FFF3E0"); self.card_crit = self._make_card("0","Alert Critical","#FFEBEE"); self.card_high = self._make_card("0","Alert High","#FFF3E0"); self.card_tickets = self._make_card("0","Ticket Aperti")
        for c in [self.card_total,self.card_ok,self.card_ko,self.card_deg,self.card_crit,self.card_high,self.card_tickets]: cl.addWidget(c)
        ml.addLayout(cl)
        self.tabs = QTabWidget(); self.tabs.addTab(self._create_alerts_tab(), "\u26a0 Alert"); self.tabs.addTab(self._create_devices_tab(), "\U0001f4cb Dispositivi"); self.tabs.addTab(self._create_overview_tab(), "\U0001f4ca Overview")
        self.tabs.currentChanged.connect(self._on_tab_changed); ml.addWidget(self.tabs)
        self.status_bar = QStatusBar(); self.setStatusBar(self.status_bar); self.status_label = QLabel("Pronto"); self.status_bar.addWidget(self.status_label, stretch=1)

    def _make_card(self, value, label, bg="#FFFFFF"):
        card = QWidget(); card.setStyleSheet(f"background:{bg};border:1px solid #CCC;border-radius:5px;"); card.setFixedHeight(78); card.setMinimumWidth(280)
        vl = QVBoxLayout(card); vl.setSpacing(2); vl.setContentsMargins(12,8,12,8)
        v = QLabel(value); v.setFont(QFont("Segoe UI",30,QFont.Bold)); v.setStyleSheet("background:transparent;border:none;color:#333;"); vl.addWidget(v)
        l = QLabel(label); l.setStyleSheet("font-size:10px;color:#666;background:transparent;border:none;"); vl.addWidget(l); card._value_label = v; return card
    def _update_card(self, card, val): card._value_label.setText(str(val))

    def _create_alerts_tab(self):
        tab = QWidget(); layout = QVBoxLayout(tab); flt = QHBoxLayout()
        flt.addWidget(QLabel("Severity:")); self.alert_sev = QComboBox(); self.alert_sev.addItems(["Tutti","CRITICAL","HIGH","MEDIUM","LOW"]); self.alert_sev.currentTextChanged.connect(self.refresh_alerts); flt.addWidget(self.alert_sev)
        flt.addWidget(QLabel("Tipo:")); self.alert_type = QComboBox(); self.alert_type.addItem("Tutti"); self.alert_type.currentTextChanged.connect(self.refresh_alerts); flt.addWidget(self.alert_type)
        flt.addWidget(QLabel("Fornitore:")); self.alert_forn = QComboBox(); self.alert_forn.addItems(["Tutti","INDRA","MII","SIRTI"]); self.alert_forn.currentTextChanged.connect(self.refresh_alerts); flt.addWidget(self.alert_forn)
        self.alert_no_ticket_btn = QPushButton("Solo senza ticket"); self.alert_no_ticket_btn.setObjectName("toggle_off"); self.alert_no_ticket_btn.setCheckable(True); self.alert_no_ticket_btn.clicked.connect(self._toggle_alert_nt); flt.addWidget(self.alert_no_ticket_btn)
        flt.addStretch(); self.alert_count_label = QLabel(""); self.alert_count_label.setStyleSheet("color:#666;"); flt.addWidget(self.alert_count_label)
        jb = QPushButton("Jira Selezionati"); jb.setObjectName("jira"); jb.clicked.connect(self._jira_from_alerts); flt.addWidget(jb)
        cb = QPushButton("Pulisci Filtri"); cb.setObjectName("secondary"); cb.clicked.connect(self._clear_alert_filters); flt.addWidget(cb); layout.addLayout(flt)
        cols = ["Severity","Tipo","DeviceID","Fornitore","DT","Trend","LTE","SSH","Mongo","Batt","Porta","Sotto C.","Descrizione","Ticket"]
        self.alert_table = FilterableTable(cols)
        self.alert_table.table.setColumnWidth(0,75); self.alert_table.table.setColumnWidth(1,180); self.alert_table.table.setColumnWidth(2,100); self.alert_table.table.setColumnWidth(3,60); self.alert_table.table.setColumnWidth(4,55); self.alert_table.table.setColumnWidth(5,60)
        for i in range(6,11): self.alert_table.table.setColumnWidth(i,45)
        self.alert_table.table.setColumnWidth(11,50); self.alert_table.table.setColumnWidth(12,300)
        self.alert_table.table.doubleClicked.connect(self._on_alert_dblclick); layout.addWidget(self.alert_table); return tab

    def _toggle_alert_nt(self):
        self._alert_no_ticket = self.alert_no_ticket_btn.isChecked()
        self.alert_no_ticket_btn.setObjectName("toggle_on" if self._alert_no_ticket else "toggle_off"); self.alert_no_ticket_btn.setStyle(self.alert_no_ticket_btn.style()); self.refresh_alerts()
    def _clear_alert_filters(self):
        self.alert_sev.setCurrentIndex(0); self.alert_type.setCurrentIndex(0); self.alert_forn.setCurrentIndex(0)
        self._alert_no_ticket = False; self.alert_no_ticket_btn.setChecked(False); self.alert_no_ticket_btn.setObjectName("toggle_off"); self.alert_no_ticket_btn.setStyle(self.alert_no_ticket_btn.style()); self.alert_table.clear_filters()
    def _on_alert_dblclick(self, index):
        it = self.alert_table.table.item(index.row(), 2)
        if it: DeviceDetailDialog(it.data(Qt.UserRole) or it.text(), self).exec_()
    def _jira_from_alerts(self):
        sel = self.alert_table.get_selected_rows_data()
        if not sel: QMessageBox.warning(self, "Nessuna selezione", "Seleziona righe."); return
        session = get_session()
        try:
            enriched = []
            for row in sel:
                did = row.get("_full_did",""); d = session.get(Device, did); td = dict(row); td["_full_did"] = did
                if d:
                    for k,a in [("tipo_malf","tipo_malfunzionamento"),("tipo_malf_jira","tipo_malf_jira"),("cluster_jira","cluster_jira"),("note","note"),("lte","check_lte"),("ssh","check_ssh"),("mongo","check_mongo"),("batteria","batteria"),("porta","porta_aperta")]: td[k] = getattr(d, a, None)
                enriched.append(td)
        finally: session.close()
        JiraTicketDialog(enriched, self).exec_()

    def _create_devices_tab(self):
        tab = QWidget(); layout = QVBoxLayout(tab); flt = QHBoxLayout()
        flt.addWidget(QLabel("Fornitore:")); self.dev_forn = QComboBox(); self.dev_forn.addItems(["Tutti","INDRA","MII","SIRTI"]); self.dev_forn.currentTextChanged.connect(self.refresh_devices); flt.addWidget(self.dev_forn)
        flt.addWidget(QLabel("Health:")); self.dev_health = QComboBox(); self.dev_health.addItems(["Tutti","OK","KO","DEGRADED"]); self.dev_health.currentTextChanged.connect(self.refresh_devices); flt.addWidget(self.dev_health)
        flt.addWidget(QLabel("Tipo:")); self.dev_tipo = QComboBox(); self.dev_tipo.addItems(["Tutti","master","slave"]); self.dev_tipo.currentTextChanged.connect(self.refresh_devices); flt.addWidget(self.dev_tipo)
        flt.addWidget(QLabel("Install:")); self.dev_install = QComboBox(); self.dev_install.addItems(["Tutti","Completa","Sotto corona"]); self.dev_install.currentTextChanged.connect(self.refresh_devices); flt.addWidget(self.dev_install)
        flt.addWidget(QLabel("Ticket:")); self.dev_ticket = QComboBox(); self.dev_ticket.addItems(["Tutti","Vuoto","Aperto","Chiuso","Scartato","Interno","Risolto"]); self.dev_ticket.currentTextChanged.connect(self.refresh_devices); flt.addWidget(self.dev_ticket)
        self.dev_no_ticket_btn = QPushButton("Solo senza ticket"); self.dev_no_ticket_btn.setObjectName("toggle_off"); self.dev_no_ticket_btn.setCheckable(True); self.dev_no_ticket_btn.clicked.connect(self._toggle_dev_nt); flt.addWidget(self.dev_no_ticket_btn)
        flt.addStretch(); self.dev_count_label = QLabel(""); self.dev_count_label.setStyleSheet("color:#666;"); flt.addWidget(self.dev_count_label)
        jb2 = QPushButton("Jira Selezionati"); jb2.setObjectName("jira"); jb2.clicked.connect(self._jira_from_devices); flt.addWidget(jb2)
        cb2 = QPushButton("Pulisci Filtri"); cb2.setObjectName("secondary"); cb2.clicked.connect(self._clear_dev_filters); flt.addWidget(cb2); layout.addLayout(flt)
        cols = ["DeviceID","Linea","Fornitore","Tipo","DT","Health","LTE","SSH","Mongo","Batt","Porta","Sotto C.","Trend","Giorni","Malf.","Ticket"]
        self.dev_table = FilterableTable(cols)
        self.dev_table.table.setColumnWidth(0,100); self.dev_table.table.setColumnWidth(1,70); self.dev_table.table.setColumnWidth(2,60); self.dev_table.table.setColumnWidth(3,48); self.dev_table.table.setColumnWidth(4,50)
        for i in range(5,12): self.dev_table.table.setColumnWidth(i,48)
        self.dev_table.table.setColumnWidth(12,60); self.dev_table.table.setColumnWidth(13,42)
        self.dev_table.table.doubleClicked.connect(self._on_dev_dblclick); layout.addWidget(self.dev_table); return tab

    def _toggle_dev_nt(self):
        self._dev_no_ticket = self.dev_no_ticket_btn.isChecked()
        self.dev_no_ticket_btn.setObjectName("toggle_on" if self._dev_no_ticket else "toggle_off"); self.dev_no_ticket_btn.setStyle(self.dev_no_ticket_btn.style()); self.refresh_devices()
    def _clear_dev_filters(self):
        self.dev_forn.setCurrentIndex(0); self.dev_health.setCurrentIndex(0); self.dev_tipo.setCurrentIndex(0); self.dev_install.setCurrentIndex(0); self.dev_ticket.setCurrentIndex(0)
        self._dev_no_ticket = False; self.dev_no_ticket_btn.setChecked(False); self.dev_no_ticket_btn.setObjectName("toggle_off"); self.dev_no_ticket_btn.setStyle(self.dev_no_ticket_btn.style()); self.dev_table.clear_filters()
    def _on_dev_dblclick(self, index):
        it = self.dev_table.table.item(index.row(), 0)
        if it: DeviceDetailDialog(it.data(Qt.UserRole) or it.text(), self).exec_()
    def _jira_from_devices(self):
        sel = self.dev_table.get_selected_rows_data()
        if not sel: QMessageBox.warning(self, "Nessuna selezione", "Seleziona righe."); return
        session = get_session()
        try:
            enriched = []
            for row in sel:
                did = row.get("_full_did",""); d = session.get(Device, did); td = dict(row); td["_full_did"] = did
                if d:
                    for k,a in [("tipo_malf","tipo_malfunzionamento"),("tipo_malf_jira","tipo_malf_jira"),("cluster_jira","cluster_jira"),("note","note"),("lte","check_lte"),("ssh","check_ssh"),("mongo","check_mongo"),("batteria","batteria"),("porta","porta_aperta")]: td[k] = getattr(d, a, None)
                enriched.append(td)
        finally: session.close()
        JiraTicketDialog(enriched, self).exec_()

    def _create_overview_tab(self):
        tab = QWidget(); layout = QVBoxLayout(tab); er = QHBoxLayout(); er.addStretch()
        self.ov_export_btn = QPushButton("Esporta Overview Excel"); self.ov_export_btn.setObjectName("success"); self.ov_export_btn.setStyleSheet("background:#2E7D32;color:white;padding:8px 16px;font-weight:bold;"); self.ov_export_btn.clicked.connect(self.export_overview); er.addWidget(self.ov_export_btn); layout.addLayout(er)
        scroll = QScrollArea(); scroll.setWidgetResizable(True); content = QWidget(); ccl = QVBoxLayout(content)
        self.ov_forn_table = QTableWidget(); self.ov_forn_table.setColumnCount(8); self.ov_forn_table.setHorizontalHeaderLabels(["Fornitore","Totale","OK","KO","Degraded","% OK","Ticket Aperti","Sotto Corona"]); self.ov_forn_table.horizontalHeader().setStretchLastSection(True)
        gf = QGroupBox("Stato per Fornitore"); fl = QVBoxLayout(); fl.addWidget(self.ov_forn_table); gf.setLayout(fl); ccl.addWidget(gf)
        self.ov_dt_table = QTableWidget(); self.ov_dt_table.setColumnCount(5); self.ov_dt_table.setHorizontalHeaderLabels(["DT","Totale","OK","KO","% OK"]); self.ov_dt_table.horizontalHeader().setStretchLastSection(True)
        gdt = QGroupBox("Stato per DT"); dtl = QVBoxLayout(); dtl.addWidget(self.ov_dt_table); gdt.setLayout(dtl); ccl.addWidget(gdt)
        self.ov_corr_table = QTableWidget(); self.ov_corr_table.setColumnCount(8); self.ov_corr_table.setHorizontalHeaderLabels(["Fornitore","Totale","LTE KO","SSH KO","Mongo KO","Porta KO","Batt KO","Disconnessi"]); self.ov_corr_table.horizontalHeader().setStretchLastSection(True)
        gc = QGroupBox("Correlazione Diagnostica"); gcl = QVBoxLayout(); gcl.addWidget(self.ov_corr_table); gc.setLayout(gcl); ccl.addWidget(gc)
        ccl.addStretch(); scroll.setWidget(content); layout.addWidget(scroll); return tab

    def refresh_data(self):
        session = get_session()
        try:
            total = session.query(Device).count()
            if total == 0: self.status_label.setText("Nessun dato. Importa un file Excel."); return
            health = dict(session.query(Device.current_health, func_count()).group_by(Device.current_health).all())
            self._update_card(self.card_total, total); self._update_card(self.card_ok, health.get("OK",0)); self._update_card(self.card_ko, health.get("KO",0)); self._update_card(self.card_deg, health.get("DEGRADED",0))
            self._update_card(self.card_crit, session.query(AnomalyEvent).filter(AnomalyEvent.severity=="CRITICAL",AnomalyEvent.acknowledged==False).count())
            self._update_card(self.card_high, session.query(AnomalyEvent).filter(AnomalyEvent.severity=="HIGH",AnomalyEvent.acknowledged==False).count())
            self._update_card(self.card_tickets, session.query(Device).filter(Device.ticket_stato=="Aperto").count())
            types = [r[0] for r in session.query(AnomalyEvent.event_type).distinct().all() if r[0]]
            self.alert_type.blockSignals(True); self.alert_type.clear(); self.alert_type.addItem("Tutti")
            for tp in sorted(types): self.alert_type.addItem(tp)
            self.alert_type.blockSignals(False)
            self.refresh_alerts(); self.refresh_devices(); self.refresh_overview(); self.status_label.setText(f"Dati: {total} dispositivi")
        except Exception as e: self.status_label.setText(f"Errore: {e}")
        finally: session.close()

    def refresh_alerts(self):
        session = get_session()
        try:
            from sqlalchemy import func, and_
            q = session.query(AnomalyEvent, Device).join(Device, AnomalyEvent.device_id == Device.device_id).filter(AnomalyEvent.acknowledged == False)
            sev = self.alert_sev.currentText()
            if sev != "Tutti": q = q.filter(AnomalyEvent.severity == sev)
            typ = self.alert_type.currentText()
            if typ != "Tutti": q = q.filter(AnomalyEvent.event_type == typ)
            forn = self.alert_forn.currentText()
            if forn != "Tutti": q = q.filter(Device.fornitore == forn)
            # Toggle: solo senza ticket
            if self._alert_no_ticket:
                q = q.filter((Device.ticket_id == None) | (Device.ticket_id == ""))
            data = []
            for event, device in q.all():
                sdid = ":".join(device.device_id.split(":")[-2:]) if ":" in device.device_id else device.device_id
                data.append({"Severity":event.severity,"Tipo":(event.event_type or "").replace("_"," "),"DeviceID":sdid,"_full_did":device.device_id,"Fornitore":device.fornitore or "-","DT":device.dt or "-","Trend":trend_str(device.trend_7d),"LTE":device.check_lte or "-","SSH":device.check_ssh or "-","Mongo":device.check_mongo or "-","Batt":device.batteria or "-","Porta":device.porta_aperta or "-","Sotto C.":"SC" if device.is_sotto_corona else "","Descrizione":event.description or "","Ticket":f"{device.ticket_id} ({device.ticket_stato})" if device.ticket_id else "-"})
            def ra(row, col):
                val = row.get(col, "")
                if col == "Severity": return colored_item(val, SEV_BG.get(val,""), SEV_COLORS.get(val,""), True)
                elif col == "DeviceID": it = QTableWidgetItem(val); it.setData(Qt.UserRole, row.get("_full_did")); it.setForeground(QColor("#0066CC")); f = it.font(); f.setBold(True); it.setFont(f); return it
                elif col in ("LTE","SSH","Mongo","Batt","Porta"): return check_item(val)
                elif col == "Sotto C." and val == "SC": return colored_item("SC","#E3F2FD","#1565C0",True)
                elif col == "Trend": it = QTableWidgetItem(val); it.setFont(QFont("Consolas",10)); return it
                elif col == "Ticket" and val != "-": return colored_item(val, "#FFEBEE", "#C62828")
                return QTableWidgetItem(str(val))
            self.alert_table.set_data(data, ra); self.alert_count_label.setText(f"{len(data)} alert")
        finally: session.close()

    def refresh_devices(self):
        session = get_session()
        try:
            q = session.query(Device)
            forn = self.dev_forn.currentText()
            if forn != "Tutti": q = q.filter(Device.fornitore == forn)
            hlth = self.dev_health.currentText()
            if hlth != "Tutti": q = q.filter(Device.current_health == hlth)
            tipo = self.dev_tipo.currentText()
            if tipo != "Tutti": q = q.filter(Device.sistema_digil.contains(tipo))
            inst = self.dev_install.currentText()
            if inst == "Completa": q = q.filter(Device.is_sotto_corona == False)
            elif inst == "Sotto corona": q = q.filter(Device.is_sotto_corona == True)
            tkt = self.dev_ticket.currentText()
            if tkt == "Vuoto": q = q.filter((Device.ticket_id == None) | (Device.ticket_id == ""))
            elif tkt != "Tutti": q = q.filter(Device.ticket_stato == tkt)
            if self._dev_no_ticket: q = q.filter((Device.ticket_id == None) | (Device.ticket_id == ""))
            data = []
            for d in q.all():
                sdid = ":".join(d.device_id.split(":")[-2:]) if ":" in d.device_id else d.device_id
                sd = d.sistema_digil or ""; ts = "M" if "master" in sd else "S" if "slave" in sd else "?"
                data.append({"DeviceID":sdid,"_full_did":d.device_id,"Linea":d.linea or "-","Fornitore":d.fornitore or "-","Tipo":ts,"DT":d.dt or "-","Health":d.current_health or "-","LTE":d.check_lte or "-","SSH":d.check_ssh or "-","Mongo":d.check_mongo or "-","Batt":d.batteria or "-","Porta":d.porta_aperta or "-","Sotto C.":"SC" if d.is_sotto_corona else "","Trend":trend_str(d.trend_7d),"Giorni":str(d.days_in_current) if d.days_in_current else "-","Malf.":d.tipo_malfunzionamento or "-","Ticket":d.ticket_id or "-"})
            def rd(row, col):
                val = row.get(col, "")
                if col == "DeviceID": it = QTableWidgetItem(val); it.setData(Qt.UserRole, row.get("_full_did")); it.setForeground(QColor("#0066CC")); f = it.font(); f.setBold(True); it.setFont(f); return it
                elif col == "Health": return colored_item(val, HEALTH_BG.get(val,""), bold=True)
                elif col in ("LTE","SSH","Mongo","Batt","Porta"): return check_item(val)
                elif col == "Sotto C." and val == "SC": return colored_item("SC","#E3F2FD","#1565C0",True)
                elif col == "Trend": it = QTableWidgetItem(val); it.setFont(QFont("Consolas",10)); return it
                elif col == "Ticket" and val != "-": return colored_item(val, "#FFEBEE", "#C62828")
                return QTableWidgetItem(str(val))
            self.dev_table.set_data(data, rd); self.dev_count_label.setText(f"{len(data)} dispositivi")
        finally: session.close()

    def refresh_overview(self):
        session = get_session()
        try:
            fornitori = ["INDRA","MII","SIRTI"]; self.ov_forn_table.setRowCount(3)
            for i, f in enumerate(fornitori):
                devs = session.query(Device).filter(Device.fornitore==f).all(); total = len(devs)
                ok = sum(1 for d in devs if d.current_health=="OK"); ko = sum(1 for d in devs if d.current_health=="KO"); deg = sum(1 for d in devs if d.current_health=="DEGRADED"); tickets = sum(1 for d in devs if d.ticket_stato=="Aperto"); sc = sum(1 for d in devs if d.is_sotto_corona); pct = round(ok/total*100,1) if total else 0
                self.ov_forn_table.setItem(i,0,colored_item(f,bold=True)); self.ov_forn_table.setItem(i,1,colored_item(total,bold=True)); self.ov_forn_table.setItem(i,2,colored_item(ok,"#E8F5E9","#2E7D32",True)); self.ov_forn_table.setItem(i,3,colored_item(ko,"#FFEBEE","#C62828",True)); self.ov_forn_table.setItem(i,4,colored_item(deg,"#FFF3E0","#E65100")); self.ov_forn_table.setItem(i,5,colored_item(f"{pct}%",bold=True)); self.ov_forn_table.setItem(i,6,colored_item(tickets)); self.ov_forn_table.setItem(i,7,colored_item(sc,"#E3F2FD"))
            self.ov_forn_table.resizeRowsToContents()
            from sqlalchemy import func
            dts = [(dt,cnt) for dt,cnt in session.query(Device.dt, func.count()).group_by(Device.dt).all() if dt]; dts.sort(key=lambda x:x[1], reverse=True); self.ov_dt_table.setRowCount(len(dts))
            for i, (dt, total) in enumerate(dts):
                ok = session.query(Device).filter(Device.dt==dt, Device.current_health=="OK").count(); pct = round(ok/total*100,1) if total else 0
                self.ov_dt_table.setItem(i,0,colored_item(dt,bold=True)); self.ov_dt_table.setItem(i,1,colored_item(total)); self.ov_dt_table.setItem(i,2,colored_item(ok,"#E8F5E9","#2E7D32")); self.ov_dt_table.setItem(i,3,colored_item(total-ok,"#FFEBEE","#C62828")); self.ov_dt_table.setItem(i,4,colored_item(f"{pct}%",bold=True))
            self.ov_dt_table.resizeRowsToContents()
            self.ov_corr_table.setRowCount(3)
            for i, f in enumerate(fornitori):
                devs = session.query(Device).filter(Device.fornitore==f).all(); total = len(devs)
                self.ov_corr_table.setItem(i,0,colored_item(f,bold=True)); self.ov_corr_table.setItem(i,1,colored_item(total))
                self.ov_corr_table.setItem(i,2,colored_item(sum(1 for d in devs if d.check_lte=="KO"),bold=True)); self.ov_corr_table.setItem(i,3,colored_item(sum(1 for d in devs if d.check_ssh=="KO"),bold=True)); self.ov_corr_table.setItem(i,4,colored_item(sum(1 for d in devs if d.check_mongo=="KO"),bold=True)); self.ov_corr_table.setItem(i,5,colored_item(sum(1 for d in devs if d.porta_aperta=="KO"),bold=True)); self.ov_corr_table.setItem(i,6,colored_item(sum(1 for d in devs if d.batteria=="KO"),bold=True)); self.ov_corr_table.setItem(i,7,colored_item(sum(1 for d in devs if d.check_lte=="KO" and d.check_ssh=="KO"),"#FFEBEE","#C62828",True))
            self.ov_corr_table.resizeRowsToContents()
        finally: session.close()

    def _on_tab_changed(self, idx):
        if idx==0: self.refresh_alerts()
        elif idx==1: self.refresh_devices()
        elif idx==2: self.refresh_overview()

    def do_import(self):
        fp, _ = QFileDialog.getOpenFileName(self, "Seleziona File Excel", "", "Excel (*.xlsx *.xls);;All (*)"); 
        if not fp: return
        self.import_btn.setEnabled(False); self.status_label.setText("Importazione...")
        self.import_thread = ImportThread(fp); self.import_thread.finished.connect(self._on_import_done); self.import_thread.error.connect(self._on_import_error); self.import_thread.start()
    def _on_import_done(self, stats, ac):
        self.import_btn.setEnabled(True); self.refresh_data()
        QMessageBox.information(self, "Import", f"Dispositivi: {stats['devices_imported']}\nAvailability: {stats['availability_records']}\nTicket nuovi: {stats.get('tickets_new',0)}\nTicket aggiornati: {stats.get('tickets_updated',0)}\nAlert: {ac}")
    def _on_import_error(self, error):
        self.import_btn.setEnabled(True); self.status_label.setText(f"Errore: {error}"); QMessageBox.critical(self, "Errore", error)
    def export_overview(self):
        import pandas as pd
        fp, _ = QFileDialog.getSaveFileName(self, "Salva Overview", str(Path.home()/"Downloads"/f"DIGIL_Overview_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"), "Excel (*.xlsx)")
        if not fp: return
        session = get_session()
        try:
            fd = []; 
            for f in ["INDRA","MII","SIRTI"]:
                devs = session.query(Device).filter(Device.fornitore==f).all(); t = len(devs); ok = sum(1 for d in devs if d.current_health=="OK"); ko = sum(1 for d in devs if d.current_health=="KO"); deg = sum(1 for d in devs if d.current_health=="DEGRADED"); tix = sum(1 for d in devs if d.ticket_stato=="Aperto"); sc = sum(1 for d in devs if d.is_sotto_corona)
                fd.append({"Fornitore":f,"Totale":t,"OK":ok,"KO":ko,"Degraded":deg,"% OK":round(ok/t*100,1) if t else 0,"Ticket Aperti":tix,"Sotto Corona":sc})
            from sqlalchemy import func
            dd = [{"DT":dt,"Totale":t,"OK":session.query(Device).filter(Device.dt==dt,Device.current_health=="OK").count(),"KO":t-session.query(Device).filter(Device.dt==dt,Device.current_health=="OK").count(),"% OK":round(session.query(Device).filter(Device.dt==dt,Device.current_health=="OK").count()/t*100,1) if t else 0} for dt,t in sorted(session.query(Device.dt,func.count()).group_by(Device.dt).all(),key=lambda x:x[1],reverse=True) if dt]
            cd = []
            for f in ["INDRA","MII","SIRTI"]:
                devs = session.query(Device).filter(Device.fornitore==f).all(); t = len(devs)
                cd.append({"Fornitore":f,"Totale":t,"LTE KO":sum(1 for d in devs if d.check_lte=="KO"),"SSH KO":sum(1 for d in devs if d.check_ssh=="KO"),"Mongo KO":sum(1 for d in devs if d.check_mongo=="KO"),"Porta KO":sum(1 for d in devs if d.porta_aperta=="KO"),"Batt KO":sum(1 for d in devs if d.batteria=="KO"),"Disconnessi":sum(1 for d in devs if d.check_lte=="KO" and d.check_ssh=="KO")})
            with pd.ExcelWriter(fp, engine='xlsxwriter') as w:
                pd.DataFrame(fd).to_excel(w, index=False, sheet_name='Stato Fornitore'); pd.DataFrame(dd).to_excel(w, index=False, sheet_name='Stato DT'); pd.DataFrame(cd).to_excel(w, index=False, sheet_name='Correlazione')
            self.status_label.setText(f"Esportato: {fp}"); QMessageBox.information(self, "Export", f"Salvato:\n{fp}")
        except Exception as e: QMessageBox.critical(self, "Errore", str(e))
        finally: session.close()

def main():
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True); QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv); app.setApplicationName("DIGIL Monitoring"); app.setOrganizationName("Terna")
    w = MainWindow(); w.show(); sys.exit(app.exec_())

if __name__ == "__main__":
    main()
