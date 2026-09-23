import sys
import os
import subprocess
import json
import asyncio
import win32com.client  # Used to close Explorer and read shortcut targets

from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                             QListWidget, QFileDialog, QWidget, QHBoxLayout, QLabel,
                             QTabWidget, QListWidgetItem, QMessageBox)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QFont

import keyboard
from winsdk.windows.security.credentials.ui import (
    UserConsentVerifier, 
    UserConsentVerificationResult, 
    UserConsentVerifierAvailability
)

CONFIG_FILE = os.path.expanduser("~/.hidden_items_config.json")

# Pure Black and White Theme (QSS)
STYLE_SHEET = """
QMainWindow, QWidget {
    background-color: #000000;
    color: #FFFFFF;
    font-family: 'Segoe UI', sans-serif;
    font-size: 14px;
}
QLabel {
    color: #FFFFFF;
    margin-bottom: 5px;
}
QPushButton {
    background-color: #000000;
    color: #FFFFFF;
    border: 2px solid #FFFFFF;
    padding: 8px 16px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #FFFFFF;
    color: #000000;
}
QListWidget {
    background-color: #000000;
    color: #FFFFFF;
    border: 1px solid #FFFFFF;
    padding: 5px;
    outline: none;
}
QListWidget::item:selected {
    background-color: #FFFFFF;
    color: #000000;
}
QListWidget::indicator {
    border: 1px solid #FFFFFF;
    width: 14px;
    height: 14px;
    background-color: #000000;
}
QListWidget::indicator:checked {
    background-color: #FFFFFF;
}
QTabWidget::pane {
    border: 1px solid #FFFFFF;
    top: -1px;
}
QTabBar::tab {
    background-color: #000000;
    color: #FFFFFF;
    border: 1px solid #FFFFFF;
    padding: 8px 20px;
    margin-right: 2px;
}
QTabBar::tab:selected {
    background-color: #FFFFFF;
    color: #000000;
    border-bottom: 1px solid #000000;
}
"""

class FileHiderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Windows Vault")
        self.resize(600, 480)
        
        # Apply Icon
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icon.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            
        self.config = self.load_config()
        self.is_hidden = False
        self.app_dict = self.scan_system_shortcuts()

        # UI Setup
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        
        self.header_label = QLabel("Press <b>Alt + ~</b> globally to toggle visibility and terminate apps.")
        self.header_label.setTextFormat(Qt.TextFormat.RichText)
        self.header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.header_label)

        self.tabs = QTabWidget()
        
        # --- TAB 1: Custom Files & Folders ---
        self.custom_tab = QWidget()
        custom_layout = QVBoxLayout()
        self.custom_list_widget = QListWidget()
        self.custom_list_widget.addItems(self.config.get("custom_items", []))
        
        btn_layout = QHBoxLayout()
        add_file_btn = QPushButton("Add File")
        add_folder_btn = QPushButton("Add Folder")
        remove_btn = QPushButton("Remove Selected")

        add_file_btn.clicked.connect(self.add_file)
        add_folder_btn.clicked.connect(self.add_folder)
        remove_btn.clicked.connect(self.remove_custom_item)

        btn_layout.addWidget(add_file_btn)
        btn_layout.addWidget(add_folder_btn)
        btn_layout.addWidget(remove_btn)
        
        custom_layout.addWidget(QLabel("Custom Vault Paths:"))
        custom_layout.addWidget(self.custom_list_widget)
        custom_layout.addLayout(btn_layout)
        self.custom_tab.setLayout(custom_layout)

        # --- TAB 2: System Apps ---
        self.apps_tab = QWidget()
        apps_layout = QVBoxLayout()
        self.app_list_widget = QListWidget()
        
        saved_apps = self.config.get("system_apps", [])
        for app_name in sorted(self.app_dict.keys()):
            item = QListWidgetItem(app_name)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if app_name in saved_apps else Qt.CheckState.Unchecked)
            self.app_list_widget.addItem(item)
            
        self.app_list_widget.itemChanged.connect(self.save_apps_state)
        
        apps_layout.addWidget(QLabel("Select apps to hide and terminate:"))
        apps_layout.addWidget(self.app_list_widget)
        self.apps_tab.setLayout(apps_layout)

        self.tabs.addTab(self.custom_tab, "Files & Folders")
        self.tabs.addTab(self.apps_tab, "System Apps")
        main_layout.addWidget(self.tabs)

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        # Register hotkey mapping backtick (`) to represent the Tilde key press 
        # so it doesn't require holding Shift
        keyboard.add_hotkey('alt+`', self.toggle_visibility)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                data = json.load(f)
                if isinstance(data, list): return {"custom_items": data, "system_apps": []}
                return data
        return {"custom_items": [], "system_apps": []}

    def save_config(self):
        with open(CONFIG_FILE, 'w') as f:
            json.dump(self.config, f)

    def scan_system_shortcuts(self):
        paths_to_scan = [
            os.path.join(os.environ.get('APPDATA', ''), r"Microsoft\Windows\Start Menu\Programs"),
            os.path.join(os.environ.get('PROGRAMDATA', ''), r"Microsoft\Windows\Start Menu\Programs"),
            os.path.join(os.environ.get('APPDATA', ''), r"Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar")
        ]
        apps = {}
        for base_path in paths_to_scan:
            if not os.path.exists(base_path): continue
            for root, _, files in os.walk(base_path):
                for file in files:
                    if file.lower().endswith(".lnk"):
                        app_name = os.path.splitext(file)[0]
                        full_path = os.path.join(root, file)
                        if app_name not in apps: apps[app_name] = []
                        if full_path not in apps[app_name]: apps[app_name].append(full_path)
        return apps

    def save_apps_state(self, item=None):
        checked_apps = [self.app_list_widget.item(i).text() for i in range(self.app_list_widget.count()) 
                        if self.app_list_widget.item(i).checkState() == Qt.CheckState.Checked]
        self.config["system_apps"] = checked_apps
        self.save_config()

    def add_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File")
        if file_path and file_path not in self.config["custom_items"]:
            self.config["custom_items"].append(file_path)
            self.custom_list_widget.addItem(file_path)
            self.save_config()

    def add_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder_path and folder_path not in self.config["custom_items"]:
            self.config["custom_items"].append(folder_path)
            self.custom_list_widget.addItem(folder_path)
            self.save_config()

    def remove_custom_item(self):
        for item in self.custom_list_widget.selectedItems():
            self.config["custom_items"].remove(item.text())
            self.custom_list_widget.takeItem(self.custom_list_widget.row(item))
        self.save_config()

    def terminate_open_windows(self):
        """Closes open File Explorer windows and terminates selected apps."""
        try:
            shell_app = win32com.client.Dispatch("Shell.Application")
            for window in shell_app.Windows():
                if os.path.basename(window.FullName).lower() == 'explorer.exe':
                    window.Quit()
        except Exception as e:
            print(f"File Explorer close error: {e}")

        try:
            wscript = win32com.client.Dispatch("WScript.Shell")
            apps_to_kill = set()
            
            # Find .exe targets from custom items
            for path in self.config.get("custom_items", []):
                if path.lower().endswith('.exe'):
                    apps_to_kill.add(os.path.basename(path))
                elif path.lower().endswith('.lnk'):
                    try:
                        target = wscript.CreateShortcut(path).TargetPath
                        if target.lower().endswith('.exe'): apps_to_kill.add(os.path.basename(target))
                    except: pass
            
            # Find .exe targets from system apps
            for app_name in self.config.get("system_apps", []):
                if app_name in self.app_dict:
                    for shortcut_path in self.app_dict[app_name]:
                        try:
                            target = wscript.CreateShortcut(shortcut_path).TargetPath
                            if target.lower().endswith('.exe'): apps_to_kill.add(os.path.basename(target))
                        except: pass
            
            # Force kill running executables
            for exe in apps_to_kill:
                subprocess.run(f'taskkill /F /IM "{exe}"', shell=True, capture_output=True)
        except Exception as e:
            print(f"App termination error: {e}")

    def toggle_visibility(self):
        print(f"\n--- Hotkey triggered! Target hidden status will be: {not self.is_hidden} ---")
        flag = "+h +s" if not self.is_hidden else "-h -s"
        
        # If we are transitioning to "Hidden", close open windows and apps
        if not self.is_hidden:
            self.terminate_open_windows()
        
        # 1. Toggle custom files and folders
        for path in self.config.get("custom_items", []):
            if os.path.exists(path):
                subprocess.run(f'attrib {flag} "{path}"', shell=True, capture_output=True)
                
        # 2. Toggle checked system apps
        for app_name in self.config.get("system_apps", []):
            if app_name in self.app_dict:
                for shortcut_path in self.app_dict[app_name]:
                    if os.path.exists(shortcut_path):
                        subprocess.run(f'attrib {flag} "{shortcut_path}"', shell=True, capture_output=True)
                        
        self.is_hidden = not self.is_hidden
        print("--- Finished applying attributes ---")


async def authenticate_user():
    try:
        availability = await UserConsentVerifier.check_availability_async()
        if availability != UserConsentVerifierAvailability.AVAILABLE:
            return False, "Windows Hello is not configured."

        result = await UserConsentVerifier.request_verification_async("Unlock Windows Vault")
        if result == UserConsentVerificationResult.VERIFIED:
            return True, "Success"
        return False, "Authentication failed."
    except Exception as e:
        return False, str(e)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLE_SHEET)
    
    is_authenticated, message = asyncio.run(authenticate_user())
    
    if not is_authenticated:
        QMessageBox.critical(None, "Access Denied", f"Verification failed:\n{message}")
        sys.exit(1)
        
    window = FileHiderApp()
    window.show()
    sys.exit(app.exec())