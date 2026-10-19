"""
Local bridge between the app's HTML screen (resources/ui) and the Python engine.

The app window shows the HTML screen in a WebView. The screen talks to this module over a
small HTTP server bound to 127.0.0.1 on a random port. Every /api call must carry a secret
token that only the app's own page receives, so other programs and websites on the computer
cannot use it. All real work (Excel, forms, HLA scanning, settings) is done here in Python.
"""
from __future__ import annotations
import asyncio
import base64
import dataclasses
import json
import os
import secrets
import shutil
import sys
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any, Callable, Dict, List, Optional, Tuple
from urllib.parse import urlparse

from .engine.models import BuyerRecord, extract_surname, parse_float, parse_int, parse_date
from .engine.excel import ExcelManager
from .engine.filler import FormFiller
from .engine.rules import get_forms_for_stage
from .engine.number_to_words import amount_words_and_figures
from .engine.tech_desc import extract_short_tech_desc
from .engine.anthropic_client import AnthropicHLAClient, map_hla_json_to_buyer, join_photos_and_pdfs_to_pdf
from .settings import SettingsManager, DEFAULT_MODEL, get_documents_folder

UI_DIR = os.path.join(os.path.dirname(__file__), "resources", "ui")
MAX_BODY_BYTES = 80 * 1024 * 1024  # several phone photos as base64

AMOUNT_FIELDS = ("consideration", "loanable_amount", "mri_sri", "non_life", "principal_interest", "lot_area")
INT_FIELDS = ("loan_term_years", "repricing_years")
FIELD_LABELS = {
    "principal_borrower": "Principal Borrower",
    "block_lot": "Block & Lot",
    "title_no": "Title Number (TCT)",
    "tech_desc": "Technical Description",
    "interest_rate": "Interest Rate",
    "consideration": "Consideration (Price)",
    "loanable_amount": "Loanable Amount",
    "date_notarization": "Date of Notarization",
    "address": "Present Address",
    "mri_sri": "MRI / SRI",
    "non_life": "Non-Life",
    "principal_interest": "Principal & Interest",
    "lot_area": "Lot Area",
    "loan_term_years": "Loan Term",
    "repricing_years": "Re-pricing",
}
TEXT_FIELDS = [f.name for f in dataclasses.fields(BuyerRecord)
               if f.name not in AMOUNT_FIELDS + INT_FIELDS + ("interest_rate", "row_idx", "extra_fields")]


class ApiError(Exception):
    def __init__(self, status: int, message: str, details: Optional[List[str]] = None):
        super().__init__(message)
        self.status = status
        self.message = message
        self.details = details or []


def safe_filename(name: str) -> str:
    """Removes characters Windows does not allow in file and folder names."""
    import re
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", name).strip(" .")
    return cleaned or "UNNAMED"


def buyer_to_dict(b: BuyerRecord) -> Dict[str, Any]:
    d = dataclasses.asdict(b)
    d["extra_fields"] = {k: v for k, v in b.extra_fields.items() if isinstance(v, (str, int, float, bool)) or v is None}
    # Values the screen shows but never edits
    d["is_affordable"] = b.is_affordable
    d["lot_label"] = b.lot_label
    d["marital_status"] = b.marital_status
    return d


def buyer_from_dict(d: Dict[str, Any]) -> Tuple[BuyerRecord, List[str]]:
    """Builds a BuyerRecord from what the screen sent; returns it with a list of fields it could not read."""
    b = BuyerRecord()
    bad: List[str] = []
    b.row_idx = parse_int(d.get("row_idx"), 0)
    for key in TEXT_FIELDS:
        val = d.get(key)
        setattr(b, key, "" if val is None else str(val).strip())
    for key in AMOUNT_FIELDS:
        raw = str(d.get(key, "") if d.get(key) is not None else "").strip()
        val = parse_float(raw, -1.0) if raw else 0.0
        if val < 0:
            bad.append(f"{FIELD_LABELS.get(key, key)}: '{raw}' is not a number")
            val = 0.0
        setattr(b, key, val)
    raw_rate = str(d.get("interest_rate", "") if d.get("interest_rate") is not None else "").strip()
    rate = parse_float(raw_rate, -1.0) if raw_rate else 0.0
    if rate < 0:
        bad.append(f"Interest Rate: '{raw_rate}' is not a number")
        rate = 0.0
    # '6.25', '6.25%' and 0.0625 all mean 6.25% (stored as 0.0625)
    b.interest_rate = rate / 100.0 if rate >= 1.0 else rate
    for key in INT_FIELDS:
        default = 30 if key == "loan_term_years" else 1
        raw = str(d.get(key, "") if d.get(key) is not None else "").strip()
        if raw and parse_float(raw, -1.0) < 0:
            bad.append(f"{FIELD_LABELS[key]}: '{raw}' is not a whole number")
        setattr(b, key, parse_int(raw, default) if raw else default)
    if b.date_notarization and not parse_date(b.date_notarization):
        bad.append(f"Date of Notarization: '{b.date_notarization}' is not a date (use mm/dd/yyyy)")
    if not b.status:
        b.status = "Single"
    if not b.co_borrower:
        b.co_borrower = "N/A"
    extra = d.get("extra_fields") or {}
    if isinstance(extra, dict):
        b.extra_fields = {str(k): v for k, v in extra.items()}
    return b, bad


class Backend:
    """Everything the screen can ask for. Independent of Toga so it can be tested directly."""

    def __init__(self, settings: SettingsManager, base_dir: str,
                 ui_call: Optional[Callable[[Callable[[], Any]], Any]] = None):
        self.settings = settings
        self.base_dir = base_dir
        self.ui_call = ui_call  # runs a coroutine factory on the app's UI loop (native dialogs)
        self.lock = threading.RLock()

    # --- File locations ---

    def excel_path(self) -> str:
        chosen = str(self.settings.get("excel_path", "") or "").strip()
        if chosen:
            return chosen
        dev_copy = os.path.join(self.base_dir, "BookData.xlsx")
        if os.path.exists(dev_copy):
            return dev_copy
        return os.path.join(get_documents_folder(), "BookData.xlsx")

    def forms_dir(self) -> str:
        chosen = str(self.settings.get("forms_folder", "") or "").strip()
        if chosen:
            return chosen
        dev_forms = os.path.join(self.base_dir, "forms")
        if os.path.isdir(dev_forms):
            return dev_forms
        return os.path.join(get_documents_folder(), "Blank Forms")

    def save_folder(self) -> str:
        return str(self.settings.get("save_folder", "") or "").strip() or get_documents_folder()

    def excel(self) -> ExcelManager:
        return ExcelManager(self.excel_path())

    # --- Buyers ---

    def excel_mtime(self) -> float:
        path = self.excel_path()
        return os.path.getmtime(path) if os.path.exists(path) else 0.0

    def get_buyers(self, _body=None) -> Dict[str, Any]:
        path = self.excel_path()
        if not os.path.exists(path):
            return {"buyers": [], "mtime": 0.0, "excel_path": path,
                    "notice": f"BookData.xlsx not found at {path}. Choose it with Load Excel, "
                              "or save a new buyer to create it there."}
        try:
            buyers = self.excel().load_buyers()
        except Exception as e:
            raise ApiError(500, f"Could not load buyers from {path}: {e}")
        return {"buyers": [buyer_to_dict(b) for b in buyers], "mtime": self.excel_mtime(),
                "excel_path": path, "notice": ""}

    def poll(self, _body=None) -> Dict[str, Any]:
        return {"mtime": self.excel_mtime()}

    def _buyer_or_error(self, body: Dict[str, Any]) -> BuyerRecord:
        buyer, bad = buyer_from_dict(body.get("buyer") or {})
        if bad:
            raise ApiError(400, "Please check these fields", bad)
        if not buyer.principal_borrower:
            raise ApiError(400, "Principal Borrower name cannot be empty.")
        return buyer

    def _save(self, buyer: BuyerRecord) -> int:
        try:
            with self.lock:
                return self.excel().save_buyer(buyer)
        except PermissionError as pe:
            raise ApiError(409, str(pe))
        except Exception as e:
            raise ApiError(500, f"Error saving to Excel: {e}")

    def save_buyer(self, body: Dict[str, Any]) -> Dict[str, Any]:
        buyer = self._buyer_or_error(body)
        row = self._save(buyer)
        return {"buyer": buyer_to_dict(buyer), "row": row, "mtime": self.excel_mtime()}

    def preview_text(self, body: Dict[str, Any]) -> Dict[str, Any]:
        """Short technical description and amount in words, as the forms will print them."""
        buyer, _bad = buyer_from_dict(body.get("buyer") or {})
        return {
            "short_tech_desc": extract_short_tech_desc(buyer.full_tech_desc, buyer.block_no, buyer.lot_no, buyer.lot_area),
            "loan_words": amount_words_and_figures(buyer.loanable_amount) if buyer.loanable_amount > 0 else "",
            "equity": buyer.equity,
            "is_affordable": buyer.is_affordable,
            "lot_label": buyer.lot_label,
        }

    # --- Forms ---

    def _filler(self) -> FormFiller:
        return FormFiller(self.forms_dir(), self.settings.settings)

    def list_forms(self, body: Dict[str, Any]) -> Dict[str, Any]:
        buyer, _bad = buyer_from_dict(body.get("buyer") or {})
        stage = body.get("stage") or "PRE-APPROVAL"
        filler = self._filler()
        forms = [{"index": i, "code": m["code"], "name": m["name"], "official": filler.has_official_template(m)}
                 for i, m in enumerate(get_forms_for_stage(stage, buyer))]
        return {"forms": forms, "forms_dir": self.forms_dir()}

    def preview_form(self, body: Dict[str, Any]) -> bytes:
        buyer, _bad = buyer_from_dict(body.get("buyer") or {})
        stage = body.get("stage") or "PRE-APPROVAL"
        forms = get_forms_for_stage(stage, buyer)
        idx = parse_int(body.get("index"), 0)
        if not 0 <= idx < len(forms):
            raise ApiError(400, "Unknown form")
        return self._filler().fill_form(forms[idx], buyer)

    def check_forms(self, body: Dict[str, Any]) -> Dict[str, Any]:
        """What the staff member should confirm before forms are made."""
        buyer = self._buyer_or_error(body)
        stage = body.get("stage") or "PRE-APPROVAL"
        required = ["block_lot", "title_no", "consideration", "loanable_amount", "interest_rate", "address"]
        if stage == "PRE-APPROVAL":
            required += ["tech_desc", "date_notarization"]
        missing = [FIELD_LABELS[k] for k in required if not getattr(buyer, k)]
        filler = self._filler()
        drafts = [m["name"] for m in get_forms_for_stage(stage, buyer) if not filler.has_official_template(m)]
        return {"missing": missing, "drafts": drafts, "forms_dir": self.forms_dir()}

    def make_forms(self, body: Dict[str, Any]) -> Dict[str, Any]:
        buyer = self._buyer_or_error(body)
        stage = body.get("stage") or "PRE-APPROVAL"
        # Save to Excel first so the forms always match the saved record
        self._save(buyer)

        stage_label = "Pre-approval" if stage == "PRE-APPROVAL" else "Post-approval"
        folder = os.path.join(self.save_folder(), safe_filename(f"{buyer.principal_borrower.upper()} - {stage_label}"))
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError as e:
            raise ApiError(500, f"Cannot create folder {folder}: {e}")
        surname = safe_filename(extract_surname(buyer.principal_borrower))

        filler = self._filler()
        files, errors = [], []
        for meta in get_forms_for_stage(stage, buyer):
            out_path = os.path.join(folder, safe_filename(f"{meta['name']} - {surname}") + ".pdf")
            try:
                pdf_bytes = filler.fill_form(meta, buyer)
                with open(out_path, "wb") as f:
                    f.write(pdf_bytes)
                files.append(out_path)
            except PermissionError:
                errors.append(f"{os.path.basename(out_path)}: close it in your PDF viewer and try again")
            except Exception as e:
                errors.append(f"{meta['name']}: {e}")

        scanned_hla = str(buyer.extra_fields.get("hla_scan_path", "") or "")
        if scanned_hla and os.path.exists(scanned_hla):
            dest = os.path.join(folder, f"HLA (scanned) - {surname}.pdf")
            try:
                shutil.copy2(scanned_hla, dest)
                files.append(dest)
            except Exception as e:
                errors.append(f"HLA scan: {e}")

        return {"folder": folder, "files": [{"path": p, "name": os.path.basename(p)} for p in files],
                "errors": errors, "buyer": buyer_to_dict(buyer), "mtime": self.excel_mtime()}

    def open_path(self, body: Dict[str, Any]) -> Dict[str, Any]:
        """Opens a made form or its folder; only paths inside the save folder are allowed."""
        path = os.path.realpath(str(body.get("path") or ""))
        root = os.path.realpath(self.save_folder())
        if not (path == root or path.startswith(root + os.sep)) or not os.path.exists(path):
            raise ApiError(403, "That file is not in the forms folder.")
        if sys.platform == "win32":
            os.startfile(path)
        else:
            import subprocess
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", path])
        return {"ok": True}

    # --- HLA scanning ---

    def scan_hla(self, body: Dict[str, Any]) -> Dict[str, Any]:
        api_key = self.settings.get("anthropic_api_key", "")
        if not api_key:
            raise ApiError(400, "Please enter your Anthropic API key in Settings before scanning an HLA.")
        files = body.get("files") or []
        if not files:
            raise ApiError(400, "Choose HLA photos (JPG/PNG) or a PDF.")
        tmp = tempfile.mkdtemp(prefix="hla_")
        try:
            paths = []
            for i, f in enumerate(files):
                name = safe_filename(os.path.basename(str(f.get("name") or f"page{i}")))
                path = os.path.join(tmp, f"{i:02d}_{name}")
                with open(path, "wb") as out:
                    out.write(base64.b64decode(f.get("data") or ""))
                paths.append(path)
            client = AnthropicHLAClient(api_key=api_key, model=self.settings.get("anthropic_model", DEFAULT_MODEL))
            try:
                data, flags = client.process_hla_files(paths)
            except Exception as e:
                raise ApiError(502, str(e))
            buyer = map_hla_json_to_buyer(data, row_idx=0)

            # Keep the scanned HLA with this buyer (named after the full name)
            try:
                scans_dir = os.path.join(self.save_folder(), "HLA Scans")
                scan_name = safe_filename(buyer.principal_borrower.upper() or "UNKNOWN BUYER")
                scan_pdf = os.path.join(scans_dir, f"HLA (scanned) - {scan_name}.pdf")
                n = 2
                while os.path.exists(scan_pdf):
                    scan_pdf = os.path.join(scans_dir, f"HLA (scanned) - {scan_name} ({n}).pdf")
                    n += 1
                join_photos_and_pdfs_to_pdf(paths, scan_pdf)
                buyer.extra_fields["hla_scan_path"] = scan_pdf
            except Exception as e:
                flags.append(f"Could not save the combined HLA scan PDF: {e}")
            return {"buyer": buyer_to_dict(buyer), "flags": flags}
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    # --- Settings ---

    def get_settings(self, _body=None) -> Dict[str, Any]:
        s = self.settings
        return {
            "excel_path": self.excel_path(),
            "forms_folder": self.forms_dir(),
            "save_folder": self.save_folder(),
            "developer_name": s.get("developer_name", ""),
            "developer_rep": s.get("developer_rep", ""),
            "project_address": s.get("project_address", ""),
            "bvs_municipality": s.get("bvs_municipality", ""),
            "witnesses_affordable": ", ".join(s.get("witnesses_affordable", [])),
            "witnesses_end_user": ", ".join(s.get("witnesses_end_user", [])),
            "anthropic_model": s.get("anthropic_model", DEFAULT_MODEL),
            "has_api_key": bool(s.get("anthropic_api_key", "")),
            "setup_complete": self.setup_complete(),
        }

    # --- First-run setup ---

    def setup_complete(self) -> bool:
        """True once this computer has been set up (or an older version already chose a workbook)."""
        if self.settings.get("setup_complete"):
            return True
        chosen = str(self.settings.get("excel_path", "") or "").strip()
        return bool(chosen) and os.path.exists(chosen)

    def finish_setup(self, body: Dict[str, Any]) -> Dict[str, Any]:
        """
        Saves the first-run choices: where BookData.xlsx is (an existing file, or a new one
        created in a chosen folder), the official blank forms folder, and where to save forms.
        """
        problems: List[str] = []

        mode = body.get("excel_mode")
        excel = ""
        if mode == "existing":
            excel = str(body.get("excel_path") or "").strip()
            if not excel:
                problems.append("Choose your BookData.xlsx file.")
            elif not excel.lower().endswith(".xlsx") or not os.path.isfile(excel):
                problems.append(f"BookData.xlsx was not found at: {excel}")
            else:
                try:
                    ExcelManager(excel).load_buyers()
                except Exception as e:
                    problems.append(f"That Excel file could not be read: {e}")
        elif mode == "new":
            folder = str(body.get("excel_folder") or "").strip()
            if not folder:
                problems.append("Choose the folder where the new BookData.xlsx should be created.")
            else:
                excel = os.path.join(folder, "BookData.xlsx")
        else:
            problems.append("Choose whether to use an existing BookData.xlsx or create a new one.")

        forms = str(body.get("forms_folder") or "").strip()
        if not forms:
            problems.append("Choose the folder with the official Pag-IBIG blank forms.")
        save = str(body.get("save_folder") or "").strip()
        if not save:
            problems.append("Choose where filled forms should be saved.")
        if problems:
            raise ApiError(400, "Please complete these before continuing", problems)

        try:
            for folder in (forms, save):
                os.makedirs(folder, exist_ok=True)
            if mode == "new":
                os.makedirs(os.path.dirname(excel), exist_ok=True)
                # Never overwrite a workbook that is already there
                if not os.path.exists(excel):
                    ExcelManager.create_empty_workbook(excel)
        except OSError as e:
            raise ApiError(500, f"Could not create the folders or workbook: {e}")

        s = self.settings
        s.set("excel_path", excel)
        s.set("forms_folder", forms)
        s.set("save_folder", save)
        for key in ("developer_name", "developer_rep", "project_address", "bvs_municipality"):
            if body.get(key) is not None:
                s.set(key, str(body.get(key)).strip())
        key = str(body.get("anthropic_api_key") or "").strip()
        if key:
            s.set("anthropic_api_key", key)
        s.set("setup_complete", True)
        return self.get_settings()

    def save_settings(self, body: Dict[str, Any]) -> Dict[str, Any]:
        s = self.settings
        # Only remember a location the user actually changed; otherwise keep it automatic
        excel = str(body.get("excel_path", "")).strip()
        if excel and excel != self.excel_path():
            s.set("excel_path", excel)
        forms = str(body.get("forms_folder", "")).strip()
        if forms and forms != self.forms_dir():
            s.set("forms_folder", forms)
        for key in ("save_folder", "developer_name", "developer_rep", "project_address", "bvs_municipality"):
            if key in body:
                s.set(key, str(body.get(key) or "").strip())
        for key in ("witnesses_affordable", "witnesses_end_user"):
            if key in body:
                s.set(key, [w.strip() for w in str(body.get(key) or "").split(",") if w.strip()])
        key = str(body.get("anthropic_api_key") or "").strip()
        if key:
            s.set("anthropic_api_key", key)
        if "anthropic_model" in body:
            s.set("anthropic_model", str(body.get("anthropic_model") or "").strip() or DEFAULT_MODEL)
        return self.get_settings()

    def clear_api_key(self, _body=None) -> Dict[str, Any]:
        self.settings.set("anthropic_api_key", "")
        return self.get_settings()

    def browse(self, body: Dict[str, Any]) -> Dict[str, Any]:
        """Native file/folder picker (the WebView cannot see real file paths)."""
        if not self.ui_call:
            raise ApiError(501, "Browsing is not available here.")
        import toga
        kind = body.get("kind")
        if kind == "excel":
            make = lambda: toga.OpenFileDialog("Choose BookData.xlsx", file_types=["xlsx"])
        else:
            make = lambda: toga.SelectFolderDialog("Choose a folder")
        path = self.ui_call(make)
        return {"path": str(path) if path else ""}

    def choose_excel(self, _body=None) -> Dict[str, Any]:
        """'Load Excel': pick BookData.xlsx, remember it, and return its buyers."""
        picked = self.browse({"kind": "excel"})["path"]
        if picked:
            self.settings.set("excel_path", picked)
        return {"picked": bool(picked), **self.get_buyers()}

    # --- Routing ---

    def routes(self) -> Dict[str, Callable[[Dict[str, Any]], Any]]:
        return {
            "/api/buyers": self.get_buyers,
            "/api/poll": self.poll,
            "/api/buyers/save": self.save_buyer,
            "/api/buyers/preview-text": self.preview_text,
            "/api/forms/list": self.list_forms,
            "/api/forms/preview": self.preview_form,
            "/api/forms/check": self.check_forms,
            "/api/forms/make": self.make_forms,
            "/api/open": self.open_path,
            "/api/hla/scan": self.scan_hla,
            "/api/settings": self.get_settings,
            "/api/settings/save": self.save_settings,
            "/api/settings/clear-key": self.clear_api_key,
            "/api/browse": self.browse,
            "/api/excel/choose": self.choose_excel,
            "/api/setup/finish": self.finish_setup,
        }


STATIC_TYPES = {".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8",
                ".css": "text/css; charset=utf-8", ".woff2": "font/woff2", ".png": "image/png",
                ".svg": "image/svg+xml", ".ico": "image/x-icon"}


class UIServer:
    """Serves the HTML screen and the /api routes on 127.0.0.1 for this app only."""

    def __init__(self, backend: Backend, ui_dir: str = UI_DIR):
        self.backend = backend
        self.ui_dir = os.path.realpath(ui_dir)
        self.token = secrets.token_urlsafe(32)
        self.routes = backend.routes()
        self.httpd = ThreadingHTTPServer(("127.0.0.1", 0), self._handler_class())
        self.httpd.daemon_threads = True
        self.port = self.httpd.server_address[1]
        self.url = f"http://127.0.0.1:{self.port}/"
        self.thread: Optional[threading.Thread] = None

    def start(self):
        self.thread = threading.Thread(target=self.httpd.serve_forever, name="pagibig-ui", daemon=True)
        self.thread.start()

    def stop(self):
        self.httpd.shutdown()
        self.httpd.server_close()

    def _handler_class(self):
        server = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, fmt, *args):  # keep the console quiet
                pass

            def _send(self, status: int, body: bytes, ctype: str):
                self.send_response(status)
                self.send_header("Content-Type", ctype)
                self.send_header("Content-Length", str(len(body)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("X-Content-Type-Options", "nosniff")
                self.end_headers()
                self.wfile.write(body)

            def _json(self, status: int, data: Any):
                self._send(status, json.dumps(data, default=str).encode("utf-8"), "application/json")

            def _host_ok(self) -> bool:
                # Blocks DNS-rebinding: only the exact address the app opened is accepted
                return self.headers.get("Host", "") == f"127.0.0.1:{server.port}"

            def do_GET(self):
                if not self._host_ok():
                    return self._send(403, b"Forbidden", "text/plain")
                path = urlparse(self.path).path
                if path.startswith("/api/"):
                    return self._api(path, {})
                if path == "/":
                    path = "/index.html"
                full = os.path.realpath(os.path.join(server.ui_dir, path.lstrip("/")))
                if not full.startswith(server.ui_dir + os.sep) or not os.path.isfile(full):
                    return self._send(404, b"Not found", "text/plain")
                with open(full, "rb") as f:
                    body = f.read()
                if full.endswith("index.html"):
                    body = body.replace(b"__APP_TOKEN__", server.token.encode("ascii"))
                self._send(200, body, STATIC_TYPES.get(os.path.splitext(full)[1], "application/octet-stream"))

            def do_POST(self):
                if not self._host_ok():
                    return self._send(403, b"Forbidden", "text/plain")
                length = int(self.headers.get("Content-Length") or 0)
                if length > MAX_BODY_BYTES:
                    return self._json(413, {"error": "The files are too large. Use fewer or smaller photos."})
                raw = self.rfile.read(length) if length else b""
                try:
                    body = json.loads(raw.decode("utf-8")) if raw else {}
                except ValueError:
                    return self._json(400, {"error": "Bad request"})
                self._api(urlparse(self.path).path, body)

            def _api(self, path: str, body: Dict[str, Any]):
                if not secrets.compare_digest(self.headers.get("X-App-Token", ""), server.token):
                    return self._json(403, {"error": "Forbidden"})
                handler = server.routes.get(path)
                if not handler:
                    return self._json(404, {"error": "Not found"})
                try:
                    result = handler(body)
                except ApiError as e:
                    return self._json(e.status, {"error": e.message, "details": e.details})
                except Exception as e:
                    return self._json(500, {"error": f"Unexpected error: {e}"})
                if isinstance(result, (bytes, bytearray)):
                    return self._send(200, bytes(result), "application/pdf")
                self._json(200, result)

        return Handler


def make_ui_call(app) -> Callable[[Callable[[], Any]], Any]:
    """Runs a native Toga dialog on the app's UI loop from a server thread and waits for the answer."""
    def call(make_dialog: Callable[[], Any]):
        async def show():
            return await app.main_window.dialog(make_dialog())
        return asyncio.run_coroutine_threadsafe(show(), app.loop).result(timeout=900)
    return call
