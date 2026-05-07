"""
DIGIL Monitoring - Maintenance API Client
==========================================
Client OAuth2 per recuperare lo stato maintenanceMode dai DigIL via API.
"""

import os
import json
import time
import threading
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Optional, Tuple, Callable

import requests
import urllib3
from dotenv import load_dotenv

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
load_dotenv()


class TokenManager:
    def __init__(self):
        self.auth_url = os.getenv("AUTH_URL")
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self._token: Optional[str] = None
        self._token_expiry: float = 0
        self._lock = threading.Lock()
        self.TOKEN_LIFETIME = 300
        self.REFRESH_MARGIN = 30

    def _is_valid(self) -> bool:
        return bool(self._token) and time.time() < (self._token_expiry - self.REFRESH_MARGIN)

    def _fetch(self) -> str:
        r = requests.post(
            self.auth_url,
            data={"grant_type": "client_credentials", "client_id": self.client_id, "client_secret": self.client_secret},
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            verify=False, timeout=30,
        )
        r.raise_for_status()
        return r.json().get("access_token")

    def get_token(self) -> str:
        with self._lock:
            if not self._is_valid():
                self._token = self._fetch()
                self._token_expiry = time.time() + self.TOKEN_LIFETIME
            return self._token

    def invalidate(self):
        with self._lock:
            self._token = None
            self._token_expiry = 0

    def is_configured(self) -> bool:
        return bool(self.auth_url and self.client_id and self.client_secret)


class MaintenanceApiClient:
    def __init__(self, tm: TokenManager):
        self.tm = tm
        base = os.getenv("BASE_URL", "https://digil-back-end-onesait.servizi.prv")
        self.config_url = f"{base}/api/v1/digils/{{deviceid}}/configuration"

    def _headers(self) -> Dict[str, str]:
        return {"Authorization": f"Bearer {self.tm.get_token()}", "Accept": "application/json"}

    def get_maintenance_status(self, deviceid: str) -> Tuple[bool, str, str]:
        """Returns (success, status, error). status: ON|OFF|NULL|ERR"""
        url = self.config_url.format(deviceid=deviceid)
        try:
            r = requests.get(url, headers=self._headers(), verify=False, timeout=20)
            if r.status_code in (401, 403):
                self.tm.invalidate()
                r = requests.get(url, headers=self._headers(), verify=False, timeout=20)
            r.raise_for_status()
            data = r.json()
            mm = (data.get("application") or {}).get("maintenanceMode")
            if mm == "ON": return True, "ON", ""
            if mm == "OFF": return True, "OFF", ""
            if mm is None: return True, "NULL", ""
            return True, str(mm), ""
        except requests.exceptions.HTTPError as e:
            code = e.response.status_code if e.response is not None else "?"
            return False, "ERR", f"HTTP {code}"
        except requests.exceptions.Timeout:
            return False, "ERR", "Timeout"
        except requests.exceptions.ConnectionError:
            return False, "ERR", "Conn fallita"
        except Exception as e:
            return False, "ERR", str(e)


def device_name_to_clientid(name: str) -> str:
    """Converte device_name DB (es. '1:1:2:15:22:DIGIL_MRN_0136') in clientid API (es. '1121522_0136').
    Se il formato non corrisponde, ritorna l'input invariato."""
    if not name:
        return name
    parts = name.split(":")
    if len(parts) >= 6:
        try:
            prefix = "".join(parts[:5])
            last_seg = parts[5]
            if "_" in last_seg:
                suffix = last_seg.split("_")[-1]
                return f"{prefix}_{suffix}"
        except Exception:
            pass
    return name


_tm: Optional[TokenManager] = None
_client: Optional[MaintenanceApiClient] = None


def get_token_manager() -> TokenManager:
    global _tm
    if _tm is None:
        _tm = TokenManager()
    return _tm


def get_client() -> MaintenanceApiClient:
    global _client
    if _client is None:
        _client = MaintenanceApiClient(get_token_manager())
    return _client


def fetch_maintenance_bulk(device_ids: List[str],
                            progress_cb: Optional[Callable[[int, int], None]] = None,
                            stop_flag: Optional[threading.Event] = None,
                            max_threads: Optional[int] = None) -> Dict[str, str]:
    """
    Recupera maintenance status per una lista di device.
    Ritorna {device_id: "ON"|"OFF"|"NULL"|"ERR"|"SKIP"}.
    """
    client = get_client()
    if not client.tm.is_configured():
        return {did: "ERR" for did in device_ids}

    if max_threads is None:
        try: max_threads = int(os.getenv("MAINT_MAX_THREADS", "80"))
        except: max_threads = 80

    results: Dict[str, str] = {}
    total = len(device_ids)
    done = 0
    lock = threading.Lock()

    def worker(did: str) -> Tuple[str, str]:
        if stop_flag is not None and stop_flag.is_set():
            return did, "SKIP"
        clientid = device_name_to_clientid(did)
        _, status, _ = client.get_maintenance_status(clientid)
        return did, status

    with ThreadPoolExecutor(max_workers=max_threads) as ex:
        futs = {ex.submit(worker, did): did for did in device_ids}
        for f in as_completed(futs):
            did = futs[f]
            try:
                _, status = f.result()
            except Exception:
                status = "ERR"
            results[did] = status
            with lock:
                done += 1
                if progress_cb:
                    try: progress_cb(done, total)
                    except: pass
            if stop_flag is not None and stop_flag.is_set():
                break

    for did in device_ids:
        results.setdefault(did, "SKIP")
    return results


# === Persistenza cache su disco ===

def _cache_path() -> Path:
    p = Path(__file__).parent / "data" / "maintenance_cache.json"
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def load_cache() -> Dict[str, str]:
    """Carica cache {device_id: status} da disco. Ritorna {} se assente o corrotta."""
    fp = _cache_path()
    if not fp.exists():
        return {}
    try:
        with open(fp, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            payload = data.get("devices") if "devices" in data else data
            return {k: str(v) for k, v in payload.items() if isinstance(k, str)}
    except Exception:
        pass
    return {}


def save_cache(cache: Dict[str, str]) -> bool:
    """Salva cache {device_id: status} su disco con timestamp."""
    try:
        fp = _cache_path()
        payload = {
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "count": len(cache),
            "devices": cache,
        }
        tmp = fp.with_suffix(".json.tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        tmp.replace(fp)
        return True
    except Exception:
        return False


def cache_last_updated() -> Optional[str]:
    """Ritorna timestamp ISO dell'ultimo aggiornamento cache, o None."""
    fp = _cache_path()
    if not fp.exists():
        return None
    try:
        with open(fp, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("updated_at") if isinstance(data, dict) else None
    except Exception:
        return None
