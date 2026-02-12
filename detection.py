"""
DIGIL Monitoring - Alert Generator
9 regole. NO_DATA: ultimo giorno = NO DATA → alert.
4 stati: COMPLETE(OK), AVAILABLE(OK), NOT AVAILABLE(KO), NO DATA(KO).
"""
import json
from datetime import date
from database import get_session, Device, AvailabilityDaily, AnomalyEvent

ACTIVE_TICKET_STATI = {"Aperto", "Interno"}

def _has_active_ticket(device) -> bool:
    if not device.ticket_id: return False
    return device.ticket_stato in ACTIVE_TICKET_STATI


class AlertGenerator:
    def __init__(self, target_date=None):
        self.target_date = target_date or date.today()
        self.count = 0

    def run(self) -> int:
        session = get_session()
        try:
            session.query(AnomalyEvent).filter(
                AnomalyEvent.event_date == self.target_date,
                AnomalyEvent.acknowledged == False).delete()
            session.flush()
            for device in session.query(Device).all():
                avail = self._get_avail(session, device.device_id)
                if not avail: continue
                self._rule_new_ko(session, device, avail)
                self._rule_recovered(session, device, avail)
                self._rule_intermittent(session, device, avail)
                self._rule_ko_no_ticket(session, device, avail)
                self._rule_open_ticket_ok(session, device, avail)
                self._rule_connectivity_lost(session, device)
                self._rule_door_alarm(session, device)
                self._rule_battery_alarm(session, device)
                self._rule_no_data(session, device, avail)
            session.commit()
            return self.count
        finally: session.close()

    def _get_avail(self, session, device_id, days=14):
        records = (session.query(AvailabilityDaily).filter(AvailabilityDaily.device_id == device_id)
                   .order_by(AvailabilityDaily.check_date.asc()).all())
        return [{"date": r.check_date, "norm": r.norm_status, "raw": r.raw_status} for r in records[-days:]]

    def _add(self, session, device, event_type, severity, description, context=None):
        session.add(AnomalyEvent(device_id=device.device_id, event_date=self.target_date,
            event_type=event_type, severity=severity, description=description,
            context_json=json.dumps(context or {}, default=str), related_ticket=device.ticket_id))
        self.count += 1

    def _rule_new_ko(self, s, d, av):
        if len(av) < 3 or av[-1]["norm"] != "KO": return
        ok_streak = sum(1 for i in range(len(av)-2, -1, -1) if av[i]["norm"] == "OK") if av[-2]["norm"] == "OK" else 0
        if ok_streak < 2: return
        has_ticket = _has_active_ticket(d)
        self._add(s, d, "NEW_KO", "MEDIUM" if has_ticket else "HIGH",
                  f"Device passato da OK a KO dopo {ok_streak} giorni" + (" (ticket aperto)" if has_ticket else " SENZA ticket"))

    def _rule_recovered(self, s, d, av):
        if len(av) < 3 or av[-1]["norm"] != "OK": return
        ko_streak = sum(1 for i in range(len(av)-2, -1, -1) if av[i]["norm"] == "KO") if av[-2]["norm"] == "KO" else 0
        if ko_streak < 2: return
        self._add(s, d, "RECOVERED", "LOW",
                  f"Tornato OK dopo {ko_streak} giorni KO" + (f" — chiudere ticket {d.ticket_id}?" if _has_active_ticket(d) else ""))

    def _rule_intermittent(self, s, d, av):
        r7 = av[-7:] if len(av) >= 7 else av
        if len(r7) < 5: return
        changes = sum(1 for i in range(1, len(r7)) if r7[i]["norm"] != r7[i-1]["norm"])
        if changes < 3: return
        self._add(s, d, "INTERMITTENT", "MEDIUM", f"{changes} cambi stato in {len(r7)} giorni")

    def _rule_ko_no_ticket(self, s, d, av):
        if not av or av[-1]["norm"] != "KO": return
        ko_days = 0
        for i in range(len(av)-1, -1, -1):
            if av[i]["norm"] == "KO": ko_days += 1
            else: break
        if ko_days < 3: return
        if _has_active_ticket(d): return
        sev = "CRITICAL" if ko_days >= 7 else "HIGH"
        ticket_note = ""
        if d.ticket_id and d.ticket_stato:
            ticket_note = f" (ticket {d.ticket_id} stato: {d.ticket_stato})"
        if d.is_sotto_corona:
            mongo_ok = d.check_mongo in (None, "OK", "-", "")
            if mongo_ok:
                self._add(s, d, "KO_NO_TICKET", "LOW",
                          f"Sotto corona Mongo OK — KO da sensori tiro assenti ({ko_days}gg){ticket_note}")
                return
        self._add(s, d, "KO_NO_TICKET", sev,
                  f"Device KO da {ko_days} giorni consecutivi SENZA ticket attivo{ticket_note}")

    def _rule_open_ticket_ok(self, s, d, av):
        if d.ticket_stato != "Aperto" or not av or av[-1]["norm"] != "OK": return
        ok_days = 0
        for i in range(len(av)-1, -1, -1):
            if av[i]["norm"] == "OK": ok_days += 1
            else: break
        if ok_days < 5: return
        self._add(s, d, "OPEN_TICKET_OK", "LOW", f"OK da {ok_days} giorni ma ticket {d.ticket_id} ancora aperto")

    def _rule_connectivity_lost(self, s, d):
        if d.check_lte != "KO" or d.check_ssh != "KO" or _has_active_ticket(d): return
        self._add(s, d, "CONNECTIVITY_LOST", "HIGH",
                  f"Disconnesso (LTE=KO, SSH=KO, Mongo={d.check_mongo or '?'}) — Nessun ticket attivo")

    def _rule_door_alarm(self, s, d):
        if d.porta_aperta != "KO": return
        has_door = _has_active_ticket(d) and d.tipo_malfunzionamento and "porta" in d.tipo_malfunzionamento.lower()
        if has_door: return
        has_any = _has_active_ticket(d)
        self._add(s, d, "DOOR_ALARM", "LOW" if has_any else "MEDIUM",
                  f"Porta aperta" + (f" (ticket {d.ticket_id} per altro)" if has_any else " — Nessun ticket"))

    def _rule_battery_alarm(self, s, d):
        if d.batteria != "KO": return
        self._add(s, d, "BATTERY_ALARM", "MEDIUM" if _has_active_ticket(d) else "HIGH", "Batteria KO — Rischio disconnessione")

    def _rule_no_data(self, s, d, av):
        """NO_DATA: se l'ultimo giorno di availability è NO DATA → alert.
        NO DATA = il dispositivo non comunica misure (ma potrebbe arrivare allarmi e diagnostiche)."""
        if not av: return
        last_raw = (av[-1].get("raw", "") or "").upper()
        if last_raw != "NO DATA": return
        if _has_active_ticket(d): return
        # Conta giorni consecutivi di NO DATA per la severity
        nd_days = 0
        for i in range(len(av)-1, -1, -1):
            raw = (av[i].get("raw", "") or "").upper()
            if raw == "NO DATA": nd_days += 1
            else: break
        sev = "HIGH" if nd_days >= 5 else "MEDIUM"
        self._add(s, d, "NO_DATA", sev,
                  f"NO DATA da {nd_days} giorn{'o' if nd_days==1 else 'i'} — il dispositivo non comunica misure")


def run_detection(target_date=None) -> int:
    return AlertGenerator(target_date).run()
