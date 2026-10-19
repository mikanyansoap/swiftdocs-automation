"""
Pag-IBIG Forms desktop/Android app. Requires Toga 0.5.

The window shows the app's HTML screen (resources/ui/index.html) in a WebView, so the app
looks the same as the web tool. Everything the screen does (reading and saving
BookData.xlsx, making forms, HLA scanning, settings) runs in Python, through the private
local bridge in webui.py.
"""
from __future__ import annotations
import os
import toga

from .settings import SettingsManager
from .webui import Backend, UIServer, make_ui_call


class PagIBIGFormsApp(toga.App):
    def startup(self):
        # Folder holding the source checkout (used for BookData.xlsx / forms during development)
        base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        config_dir = os.environ.get("PAGIBIG_FORMS_CONFIG_DIR") or os.path.join(os.path.expanduser("~"), ".pagibig_forms")
        self.settings_manager = SettingsManager(config_dir)

        self.backend = Backend(self.settings_manager, base_dir, ui_call=make_ui_call(self))
        self.server = UIServer(self.backend)
        self.server.start()

        self.webview = toga.WebView(url=self.server.url, style=toga.style.Pack(flex=1))
        self.main_window = toga.MainWindow(title=self.formal_name, size=(1280, 860))
        self.main_window.content = self.webview
        self.main_window.show()

    def on_exit(self):
        try:
            self.server.stop()
        except Exception:
            pass
        return True


def main():
    return PagIBIGFormsApp("Pag-IBIG Forms", "org.summerfield.pagibigforms", icon="resources/pagibig_forms")
