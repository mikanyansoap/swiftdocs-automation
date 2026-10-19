"""
Settings manager for Pag-IBIG Forms.
Stores device-specific configurations in a local JSON file:
- Save folder for filled forms, BookData.xlsx location, blank forms folder
- Project address
- Developer representative
- Witnesses (Affordable and End-User)
- Developer name
- BVS municipality
- Anthropic API key (encrypted on Windows) and model
"""
from __future__ import annotations
import base64
import json
import os
import sys
from typing import Any, Dict


def get_documents_folder() -> str:
    if sys.platform == "win32":
        return os.path.join(os.path.expanduser("~"), "Documents", "Pag-IBIG Forms")
    # Android / Linux / fallback
    return os.path.join(os.path.expanduser("~"), "Download", "Pag-IBIG Forms")


def get_default_save_folder() -> str:
    return get_documents_folder()


DEFAULT_MODEL = "claude-opus-5-5"
# Model IDs that no longer work; replaced with DEFAULT_MODEL when settings load
RETIRED_MODEL_PREFIXES = ("claude-3", "claude-2", "claude-instant")

DEFAULT_SETTINGS: Dict[str, Any] = {
    "save_folder": get_default_save_folder(),
    "excel_path": "",      # blank = automatic (see app.resolve_excel_path)
    "forms_folder": "",    # blank = automatic (see app.resolve_forms_folder)
    "project_address": "BLOOMFIELD TERRACES, Brgy. Maybancal, Morong, Rizal",
    "developer_name": "SUMMERFIELD RESOURCES CORPORATION",
    "developer_rep": "JENNIFER B. PEDUCA",
    "bvs_municipality": "MORONG, RIZAL",
    "witnesses_affordable": ["DENIELLE S. TOLENTINO", "SALVACION M. AZCARRAGA"],
    "witnesses_end_user": ["VALENTINO P. HALFON", "SALVACION M. AZCARRAGA"],
    "anthropic_model": DEFAULT_MODEL
}

SECRET_KEY = "anthropic_api_key"
SECRET_STORE_KEY = "anthropic_api_key_protected"


# --- API key protection ---------------------------------------------------
# On Windows the key is encrypted with DPAPI, so only the same Windows user on the
# same PC can decrypt it; copying settings.json elsewhere does not reveal the key.
# On Android the settings file lives in the app's private storage, which other
# apps cannot read.

def _dpapi(data: bytes, protect: bool) -> bytes:
    import ctypes
    from ctypes import wintypes

    class DATA_BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    buf = ctypes.create_string_buffer(data, len(data))
    blob_in = DATA_BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = DATA_BLOB()
    fn = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    # CRYPTPROTECT_UI_FORBIDDEN = 0x1
    ok = fn(ctypes.byref(blob_in), None, None, None, None, 0x1, ctypes.byref(blob_out))
    if not ok:
        raise OSError("Windows could not encrypt/decrypt the API key")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        kernel32.LocalFree(blob_out.pbData)


def protect_secret(value: str) -> str:
    if not value:
        return ""
    if sys.platform == "win32":
        return "dpapi:" + base64.b64encode(_dpapi(value.encode("utf-8"), True)).decode("ascii")
    return "plain:" + value


def unprotect_secret(stored: str) -> str:
    if not stored:
        return ""
    if stored.startswith("dpapi:"):
        try:
            return _dpapi(base64.b64decode(stored[6:]), False).decode("utf-8")
        except Exception:
            return ""  # encrypted by another Windows user/PC: ask for the key again
    if stored.startswith("plain:"):
        return stored[6:]
    return ""


class SettingsManager:
    def __init__(self, config_dir: str):
        self.config_dir = config_dir
        self.config_file = os.path.join(config_dir, "settings.json")
        self.settings = dict(DEFAULT_SETTINGS)
        self.load()

    def load(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.settings.update(data)
            except Exception:
                pass

        changed = False
        # Move a key saved in plain text by older versions into protected storage
        legacy_key = self.settings.pop(SECRET_KEY, "")
        if legacy_key:
            self.settings[SECRET_STORE_KEY] = protect_secret(legacy_key)
            changed = True
        model = str(self.settings.get("anthropic_model", "")).strip()
        if not model or model.startswith(RETIRED_MODEL_PREFIXES):
            self.settings["anthropic_model"] = DEFAULT_MODEL
            changed = True
        if changed:
            self.save()

    def save(self):
        os.makedirs(self.config_dir, exist_ok=True)
        tmp = self.config_file + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=2)
            os.replace(tmp, self.config_file)
            try:
                os.chmod(self.config_file, 0o600)
            except OSError:
                pass
        except Exception:
            pass

    def get(self, key: str, default: Any = None) -> Any:
        if key == SECRET_KEY:
            return unprotect_secret(self.settings.get(SECRET_STORE_KEY, "")) or (default or "")
        return self.settings.get(key, default)

    def set(self, key: str, value: Any):
        if key == SECRET_KEY:
            self.settings[SECRET_STORE_KEY] = protect_secret(str(value or "").strip())
        else:
            self.settings[key] = value
        self.save()
