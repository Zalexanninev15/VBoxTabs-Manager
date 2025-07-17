# VBoxTabs Manager
# Copyright (c) 2025 Zalexanninev15
# https://github.com/Zalexanninev15/VBoxTabs-Manager
#
# Licensed under the MIT License. See LICENSE file in the project root for full license information.
"""
VBoxTabs Manager - Combining VirtualBox, Hyper-V, and other windows into a single tabbed interface.
"""

import sys
import os
import subprocess
import json
import platform
import locale
import re
from datetime import datetime
from functools import partial
import time

# --- Platform-specific imports ---
IS_WINDOWS = platform.system() == "Windows"
IS_LINUX = platform.system() == "Linux"

if IS_WINDOWS:
    import win32gui
    import win32process
    import win32con
    import win32api
    try:
        import winreg
        WINREG_AVAILABLE = True
    except ImportError:
        WINREG_AVAILABLE = False
else:
    WINREG_AVAILABLE = False

from PySide6.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget,
                               QVBoxLayout, QPushButton, QMessageBox, QLabel,
                               QDialog, QHBoxLayout, QInputDialog, QCheckBox,
                               QStyleFactory, QComboBox, QToolButton, QMenu,
                               QStyle, QTabBar, QSpinBox, QLineEdit, QGridLayout,
                               QGroupBox, QFileDialog, QTableWidget,
                               QTableWidgetItem, QHeaderView, QFrame)
from PySide6.QtCore import Qt, QTimer, Signal, QPoint, QSize
from PySide6.QtGui import QFont, QAction, QIcon, QPixmap, QMouseEvent, QPalette

# --- Helper Functions ---
# QDarkStyle
try:
    import qdarkstyle
    QDARKSTYLE_AVAILABLE = True
except ImportError:
    QDARKSTYLE_AVAILABLE = False

# QtAwesome
try:
    import qtawesome as qta
    QTAWESOME_AVAILABLE = True
except ImportError:
    QTAWESOME_AVAILABLE = False

# --- Constants ---
SCANCODE_MAP = {
    'A': '1e', 'B': '30', 'C': '2e', 'D': '20', 'E': '12', 'F': '21', 'G': '22',
    'H': '23', 'I': '17', 'J': '24', 'K': '25', 'L': '26', 'M': '32', 'N': '31',
    'O': '18', 'P': '19', 'Q': '10', 'R': '13', 'S': '1f', 'T': '14', 'U': '16',
    'V': '2f', 'W': '11', 'X': '2d', 'Y': '15', 'Z': '2c'
}
VERSION = '2.0-beta 2'

# --- Helper Functions (Global Scope) ---
def get_settings_path():
    return os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), 'settings.json')

def load_settings(path):
    defaults = {
        "auto_attach": True, 
        "attach_vbox_manager": False,
        "refresh_interval": 5,
        "vbox_path": r"C:\Program Files\Oracle\VirtualBox\VirtualBox.exe",
        "theme": "Fusion", 
        "dpi_scaling": "Auto",
        "icon_theme": "Awesome",
        "toolbar_button_size": 16,
        "custom_names": {}, 
        "use_custom_names": True,
        "custom_actions": [
            {
                "name": "VM Settings",
                "tooltip": "Opens the virtual machine settings.",
                "key": "S"
            },
            {
                "name": "Fullscreen",
                "tooltip": "Switching to full screen mode.",
                "key": "F"
            },
            {
                "name": "Adjust Size",
                "tooltip": "Size adjustment (helps to adapt the VM screen)",
                "key": "A"
            },
            {
                "name": "Screenshot",
                "tooltip": "Create a screenshot of the virtual machine screen.",
                "key": "E"
            }
        ]
    }
    if not os.path.exists(path):
        return defaults
    try:
        with open(path, 'r') as f:
            settings = json.load(f)
        settings.update({k: v for k, v in defaults.items() if k not in settings})
        return settings
    except Exception:
        return defaults

def is_windows_light_theme():
    if not IS_WINDOWS or not WINREG_AVAILABLE: return False
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        return value == 1
    except (FileNotFoundError, OSError): return False

# --- Abstract Base Classes for Platform-Specific Code ---
class AbstractWindowManager:
    def find_embeddable_windows(self): raise NotImplementedError
    def set_window_parent(self, child_hwnd, parent_hwnd): raise NotImplementedError
    def restore_window_style(self, hwnd, native_details): raise NotImplementedError
    def get_window_title(self, hwnd): raise NotImplementedError
    def get_process_id(self, hwnd): raise NotImplementedError
    def terminate_process_by_id(self, pid): raise NotImplementedError
    def is_window(self, hwnd): raise NotImplementedError
    def get_window_from_point(self, point): raise NotImplementedError
    def get_toplevel_parent(self, hwnd): raise NotImplementedError
    def get_windowed_processes(self, main_hwnd): raise NotImplementedError

class LinuxWindowManager(AbstractWindowManager):
    def __init__(self): print("Warning: Linux support is not yet implemented.")
    def find_embeddable_windows(self): return []
    def set_window_parent(self, c, p): QMessageBox.critical(None, "Not Implemented", "Window management is not yet supported on Linux.")
    def restore_window_style(self, c, n): pass
    def get_window_title(self, h): return "N/A"
    def get_process_id(self, h): return -1
    def terminate_process_by_id(self, p): pass
    def is_window(self, h): return False
    def get_window_from_point(self, p): return None
    def get_toplevel_parent(self, h): return h
    def get_windowed_processes(self, main_hwnd): return []

class WindowsWindowManager(AbstractWindowManager):
    def __init__(self):
        self.found_windows = []
        self.WS_CHILD = 0x40000000; self.GWL_STYLE = -16; self.SWP_NOSIZE = 0x0001
        self.SWP_NOMOVE = 0x0002; self.SWP_NOZORDER = 0x0004; self.SWP_FRAMECHANGED = 0x0020
        self.WS_CAPTION = 0x00C00000; self.WS_THICKFRAME = 0x00040000; self.WS_MINIMIZEBOX = 0x00020000
        self.WS_MAXIMIZEBOX = 0x00010000; self.WS_SYSMENU = 0x00080000

    def _enum_windows_callback(self, hwnd, _):
        if not win32gui.IsWindowVisible(hwnd): return True
        window_title = win32gui.GetWindowText(hwnd)
        if not window_title: return True
        
        is_vbox_vm = ("[Running]" in window_title or "[Работает]" in window_title) and " Oracle VirtualBox" in window_title
        if is_vbox_vm:
            vm_name_from_title = window_title.split(" [")[0]
            self.found_windows.append({'hwnd': hwnd, 'title': vm_name_from_title, 'type': 'VirtualBox'})
            return True
            
        is_vbox_manager = "Oracle VirtualBox " in window_title
        if is_vbox_manager:
            self.found_windows.append({'hwnd': hwnd, 'title': "VirtualBox Manager", 'type': 'VirtualBox Manager'})
            return True

        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        try:
            handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
            proc_name = win32process.GetModuleFileNameEx(handle, 0)
            win32api.CloseHandle(handle)
            if "vmconnect.exe" in proc_name.lower() and win32gui.GetParent(hwnd) == 0:
                self.found_windows.append({'hwnd': hwnd, 'title': window_title, 'type': 'Hyper-V'})
        except: pass
        return True

    def find_embeddable_windows(self):
        self.found_windows = []
        win32gui.EnumWindows(self._enum_windows_callback, None)
        return self.found_windows

    def set_window_parent(self, hwnd, parent_hwnd):
        style = win32gui.GetWindowLong(hwnd, self.GWL_STYLE)
        original_details = {'style': style, 'parent': win32gui.GetParent(hwnd)}
        new_style = (style & ~(self.WS_CAPTION | self.WS_THICKFRAME | self.WS_MINIMIZEBOX | self.WS_MAXIMIZEBOX | self.WS_SYSMENU)) | self.WS_CHILD
        win32gui.SetWindowLong(hwnd, self.GWL_STYLE, new_style)
        win32gui.SetParent(hwnd, parent_hwnd)
        win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0, self.SWP_NOMOVE | self.SWP_NOSIZE | self.SWP_NOZORDER | self.SWP_FRAMECHANGED)
        return original_details

    def restore_window_style(self, hwnd, original_details):
        if not self.is_window(hwnd): return
        current_style = win32gui.GetWindowLong(hwnd, self.GWL_STYLE)
        style_to_restore = original_details.get('style', current_style)
        new_style = (current_style & ~self.WS_CHILD) | (style_to_restore & (self.WS_CAPTION | self.WS_THICKFRAME | self.WS_MINIMIZEBOX | self.WS_MAXIMIZEBOX | self.WS_SYSMENU))
        win32gui.SetWindowLong(hwnd, self.GWL_STYLE, new_style)
        win32gui.SetParent(hwnd, original_details.get('parent', 0))
        win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0, self.SWP_NOMOVE | self.SWP_NOSIZE | self.SWP_NOZORDER | self.SWP_FRAMECHANGED)
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

    def get_window_title(self, hwnd): return win32gui.GetWindowText(hwnd)
    def get_process_id(self, hwnd):
        if self.is_window(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return pid
        return None
    def terminate_process_by_id(self, pid):
        try:
            h = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
            win32api.TerminateProcess(h, 0)
            win32api.CloseHandle(h)
            return True
        except: return False
    def is_window(self, hwnd): return win32gui.IsWindow(hwnd)
    def get_window_from_point(self, point): return win32gui.WindowFromPoint(point)
    def get_toplevel_parent(self, hwnd):
        while parent := win32gui.GetParent(hwnd): hwnd = parent
        return hwnd

    def get_windowed_processes(self, main_hwnd):
        results = []
        def callback(hwnd, extra):
            if hwnd == main_hwnd or not win32gui.IsWindowVisible(hwnd) or win32gui.GetParent(hwnd) != 0:
                return True
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return True
            
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc_name = "N/A"
            try:
                handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                proc_name = os.path.basename(win32process.GetModuleFileNameEx(handle, 0))
                win32api.CloseHandle(handle)
            except: pass
            
            results.append({'hwnd': hwnd, 'pid': pid, 'name': proc_name, 'title': title})
            return True
        
        win32gui.EnumWindows(callback, None)
        return results

# --- Application Classes ---

class ProcessListDialog(QDialog):
    def __init__(self, window_manager, main_hwnd, reload_icon=None, parent=None):
        super().__init__(parent)
        self.window_manager = window_manager
        self.main_hwnd = main_hwnd
        self.setWindowTitle("Attach Window from Process List")
        self.setMinimumSize(800, 600)
        
        layout = QVBoxLayout(self)
        
        search_layout = QHBoxLayout()
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search by PID, Name, or Title...")
        self.search_box.textChanged.connect(self.filter_table)
        self.search_box.setClearButtonEnabled(True)
        search_layout.addWidget(self.search_box)
        
        self.reload_button = QToolButton()
        if reload_icon:
            self.reload_button.setIcon(reload_icon)
        else:
            self.reload_button.setIcon(QApplication.style().standardIcon(QStyle.SP_BrowserReload))
        self.reload_button.setToolTip("Refresh Process List")
        self.reload_button.clicked.connect(self.populate_table)
        search_layout.addWidget(self.reload_button)
        layout.addLayout(search_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["HWND", "PID", "Process Name", "Window Title"])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        self.table.doubleClicked.connect(self.accept)
        layout.addWidget(self.table)
        
        button_box = QHBoxLayout()
        button_box.addStretch()
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        self.attach_button = QPushButton("Attach Selected")
        self.attach_button.clicked.connect(self.accept)
        self.attach_button.setDefault(True)
        button_box.addWidget(self.cancel_button)
        button_box.addWidget(self.attach_button)
        layout.addLayout(button_box)
        
        self.populate_table()

    def populate_table(self):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        processes = self.window_manager.get_windowed_processes(self.main_hwnd)
        for proc in processes:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(str(proc['hwnd'])))
            self.table.setItem(row, 1, QTableWidgetItem(str(proc['pid'])))
            self.table.setItem(row, 2, QTableWidgetItem(proc['name']))
            self.table.setItem(row, 3, QTableWidgetItem(proc['title']))
        self.table.setSortingEnabled(True)

    def filter_table(self, text):
        text = text.lower()
        for i in range(self.table.rowCount()):
            match = (text in self.table.item(i, 0).text().lower() or
                     text in self.table.item(i, 1).text().lower() or
                     text in self.table.item(i, 2).text().lower() or
                     text in self.table.item(i, 3).text().lower())
            self.table.setRowHidden(i, not match)
            
    def get_selected_window_info(self):
        current_row = self.table.currentRow()
        if current_row < 0:
            return None
        hwnd = int(self.table.item(current_row, 0).text())
        title = self.table.item(current_row, 3).text()
        return {'hwnd': hwnd, 'title': title, 'type': 'Listed'}

class CustomNamesDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("Manage Custom Tab Names")
        self.setMinimumSize(700, 500)
        
        layout = QVBoxLayout(self)
        info_label = QLabel(
            "Here you can pre-define custom names for windows based on their original title.\n"
            "These names will be applied automatically when the window is attached."
        )
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Original Window Title", "Custom Name"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.table)
        
        button_layout = QHBoxLayout()
        self.add_button = QPushButton("Add..."); self.add_button.clicked.connect(self.add_name)
        self.edit_button = QPushButton("Edit..."); self.edit_button.clicked.connect(self.edit_name)
        self.remove_button = QPushButton("Remove"); self.remove_button.clicked.connect(self.remove_name)
        button_layout.addWidget(self.add_button); button_layout.addWidget(self.edit_button); button_layout.addWidget(self.remove_button)
        button_layout.addStretch()
        layout.addLayout(button_layout)

        self.populate_table()
        
        button_box = QHBoxLayout()
        button_box.addStretch()
        self.close_button = QPushButton("Close"); self.close_button.clicked.connect(self.accept)
        button_box.addWidget(self.close_button)
        layout.addLayout(button_box)
        
    def populate_table(self):
        self.table.setRowCount(0)
        for original, custom in self.settings.get("custom_names", {}).items():
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(original))
            self.table.setItem(row, 1, QTableWidgetItem(custom))

    def add_name(self):
        original, ok = QInputDialog.getText(self, "Add Mapping", "Enter the exact Original Window Title:")
        if ok and original:
            custom, ok2 = QInputDialog.getText(self, "Add Mapping", f"Enter the Custom Name for '{original}':")
            if ok2 and custom:
                self.settings["custom_names"][original] = custom
                self.populate_table()

    def edit_name(self):
        row = self.table.currentRow()
        if row < 0: return
        original = self.table.item(row, 0).text()
        current_custom = self.table.item(row, 1).text()
        custom, ok = QInputDialog.getText(self, "Edit Mapping", f"Enter new Custom Name for '{original}':", text=current_custom)
        if ok and custom:
            self.settings["custom_names"][original] = custom
            self.populate_table()
            
    def remove_name(self):
        row = self.table.currentRow()
        if row < 0: return
        original = self.table.item(row, 0).text()
        if QMessageBox.question(self, "Confirm Remove", f"Remove mapping for '{original}'?") == QMessageBox.Yes:
            if original in self.settings["custom_names"]:
                del self.settings["custom_names"][original]
            self.populate_table()

class CustomActionDialog(QDialog):
    def __init__(self, action=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Custom Action" if action else "Add Custom Action")
        layout = QGridLayout(self)

        layout.addWidget(QLabel("Name:"), 0, 0)
        self.name_edit = QLineEdit(action['name'] if action else "")
        self.name_edit.setPlaceholderText("e.g., Fullscreen")
        layout.addWidget(self.name_edit, 0, 1)

        layout.addWidget(QLabel("Tooltip:"), 1, 0)
        self.tooltip_edit = QLineEdit(action['tooltip'] if action else "")
        self.tooltip_edit.setPlaceholderText("e.g., Toggle Fullscreen (Right Ctrl+F)")
        layout.addWidget(self.tooltip_edit, 1, 1)

        layout.addWidget(QLabel("Key (A-Z):"), 2, 0)
        self.key_edit = QLineEdit(action['key'] if action else "")
        self.key_edit.setMaxLength(1)
        self.key_edit.setPlaceholderText("F")
        layout.addWidget(self.key_edit, 2, 1)

        button_box = QHBoxLayout()
        button_box.addStretch()
        self.cancel_button = QPushButton("Cancel"); self.cancel_button.clicked.connect(self.reject)
        self.save_button = QPushButton("Save"); self.save_button.clicked.connect(self.accept)
        self.save_button.setDefault(True)
        button_box.addWidget(self.cancel_button)
        button_box.addWidget(self.save_button)
        layout.addLayout(button_box, 3, 0, 1, 2)

    def get_action(self):
        key = self.key_edit.text().upper()
        if not key or key not in SCANCODE_MAP:
            QMessageBox.warning(self, "Invalid Key", "Please enter a single letter (A-Z) for the key.")
            return None
        return {
            "name": self.name_edit.text(),
            "tooltip": self.tooltip_edit.text(),
            "key": key
        }

    def accept(self):
        if self.get_action() is not None:
            super().accept()

class SettingsDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.parent = parent
        self.setWindowTitle("Settings")
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15); main_layout.setSpacing(10)
        
        # --- General Group ---
        general_group = QGroupBox("General Settings")
        general_layout = QGridLayout(general_group)
        self.auto_attach_checkbox = QCheckBox("Automatically attach VirtualBox VM windows")
        self.auto_attach_checkbox.setChecked(self.settings.get("auto_attach", True))
        general_layout.addWidget(self.auto_attach_checkbox, 0, 0, 1, 2)
        
        self.attach_vbox_manager_checkbox = QCheckBox("Automatically attach VirtualBox Manager window")
        self.attach_vbox_manager_checkbox.setChecked(self.settings.get("attach_vbox_manager", False))
        general_layout.addWidget(self.attach_vbox_manager_checkbox, 1, 0, 1, 2)
        
        self.use_custom_names_map_checkbox = QCheckBox("Apply automatic tab names   ⇒   ")
        self.use_custom_names_map_checkbox.setChecked(self.settings.get("use_custom_names", False))
        general_layout.addWidget(self.use_custom_names_map_checkbox, 2, 0, 1, 1)

        self.custom_names_button = QPushButton("Manage Names..."); self.custom_names_button.clicked.connect(self.open_custom_names_dialog)
        general_layout.addWidget(self.custom_names_button, 2, 1, 1, 1)

        interval_label = QLabel("Window detection interval (seconds):")
        general_layout.addWidget(interval_label, 3, 0)
        self.refresh_interval_spinbox = QSpinBox(); self.refresh_interval_spinbox.setRange(1, 60); self.refresh_interval_spinbox.setValue(self.settings.get("refresh_interval", 5))
        general_layout.addWidget(self.refresh_interval_spinbox, 3, 1)
        path_label = QLabel("VirtualBox executable path:")
        general_layout.addWidget(path_label, 4, 0)
        vbox_path_layout = QHBoxLayout()
        self.vbox_path_edit = QLineEdit(self.settings.get("vbox_path", r"C:\Program Files\Oracle\VirtualBox\VirtualBox.exe")); self.vbox_path_edit.setCursorPosition(0)
        vbox_path_layout.addWidget(self.vbox_path_edit)
        self.browse_button = QPushButton("Browse..."); self.browse_button.clicked.connect(self.browse_vbox_path)
        vbox_path_layout.addWidget(self.browse_button)
        general_layout.addLayout(vbox_path_layout, 4, 1)
        general_layout.setColumnStretch(1, 1)
        main_layout.addWidget(general_group)

        # --- Display Group ---
        display_group = QGroupBox("Display Settings")
        display_layout = QGridLayout(display_group)
        theme_label = QLabel("Theme:"); display_layout.addWidget(theme_label, 0, 0)
        self.theme_combo = QComboBox(); theme_names = []
        if IS_WINDOWS and not is_windows_light_theme(): theme_names.append("Dark")
        theme_names.extend(["Light", "Classic", "Fusion"])
        if QDARKSTYLE_AVAILABLE: theme_names.append("QDark")
        self.theme_combo.addItems(theme_names); current_theme = self.settings.get("theme", "Fusion")
        if current_theme in theme_names: self.theme_combo.setCurrentText(current_theme)
        else: self.theme_combo.setCurrentText("Fusion" if "Fusion" in theme_names else (theme_names[0] if theme_names else ""))
        display_layout.addWidget(self.theme_combo, 0, 1)
        
        if QTAWESOME_AVAILABLE:
            icon_theme_label = QLabel("Icon Theme:"); display_layout.addWidget(icon_theme_label, 1, 0)
            self.icon_theme_combo = QComboBox(); self.icon_theme_combo.addItems(["Standard", "Awesome"])
            self.icon_theme_combo.setCurrentText(self.settings.get("icon_theme", "Standard"))
            display_layout.addWidget(self.icon_theme_combo, 1, 1)

        toolbar_size_label = QLabel("Toolbar Icon Size:"); display_layout.addWidget(toolbar_size_label, 2, 0)
        self.toolbar_size_spinbox = QSpinBox(); self.toolbar_size_spinbox.setRange(16, 48); self.toolbar_size_spinbox.setValue(self.settings.get("toolbar_button_size", 16))
        display_layout.addWidget(self.toolbar_size_spinbox, 2, 1)

        dpi_label = QLabel("DPI Scaling:"); display_layout.addWidget(dpi_label, 3, 0)
        self.dpi_scaling_combo = QComboBox(); self.dpi_scaling_combo.addItems(["Auto", "100%", "125%", "150%", "175%", "200%"])
        self.dpi_scaling_combo.setCurrentText(self.settings.get("dpi_scaling", "Auto"))
        display_layout.addWidget(self.dpi_scaling_combo, 3, 1)
        display_layout.setColumnStretch(1, 1)
        main_layout.addWidget(display_group)

        # --- Custom Actions Group ---
        actions_group = QGroupBox("Custom Toolbar Actions (Right Ctrl + Key)")
        actions_layout = QVBoxLayout(actions_group)
        
        self.host_key_warning = QLabel(
            "<b>Important:</b> These actions (and Ctrl+Alt+Del) may fail if your VirtualBox "
            "<b>Host Key</b> is set to the default <b>Right Ctrl</b>. <br>"
            "To fix this, go to VirtualBox -> File -> Preferences -> Input -> Virtual Machine "
            "and change the Host Key to something else (e.g., Left Ctrl+Shift)."
        )
        self.host_key_warning.setWordWrap(True)
        self.update_warning_style()
        actions_layout.addWidget(self.host_key_warning)
        actions_layout.addSpacing(10)

        self.actions_table = QTableWidget(); self.actions_table.setColumnCount(3); self.actions_table.setHorizontalHeaderLabels(["Name", "Tooltip", "Key"])
        self.actions_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.actions_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.actions_table.setEditTriggers(QTableWidget.NoEditTriggers); self.actions_table.setSelectionBehavior(QTableWidget.SelectRows)
        actions_layout.addWidget(self.actions_table)
        action_button_layout = QHBoxLayout()
        self.add_action_button = QPushButton("Add"); self.add_action_button.clicked.connect(self.add_action)
        self.edit_action_button = QPushButton("Edit"); self.edit_action_button.clicked.connect(self.edit_action)
        self.remove_action_button = QPushButton("Remove"); self.remove_action_button.clicked.connect(self.remove_action)
        action_button_layout.addStretch(); action_button_layout.addWidget(self.add_action_button); action_button_layout.addWidget(self.edit_action_button); action_button_layout.addWidget(self.remove_action_button)
        actions_layout.addLayout(action_button_layout)
        self.populate_actions_table()
        main_layout.addWidget(actions_group)

        button_layout = QHBoxLayout()
        button_layout.addStretch(1)
        self.cancel_button = QPushButton("Cancel"); self.cancel_button.clicked.connect(self.reject); button_layout.addWidget(self.cancel_button)
        self.save_button = QPushButton("Save"); self.save_button.clicked.connect(self.accept)
        button_layout.addWidget(self.save_button)
        main_layout.addLayout(button_layout)
        
        self.resize(600, 800)
        self.theme_combo.currentTextChanged.connect(self.update_warning_style)

    def update_warning_style(self, theme_name=None):
        if theme_name is None:
            theme_name = self.theme_combo.currentText()
            
        is_light_mode = (theme_name == "Light") or ((theme_name == "Fusion" or theme_name == "Classic") and is_windows_light_theme())
        if is_light_mode:
            self.host_key_warning.setStyleSheet("QLabel { background-color: #FFFACD; border: 1px solid #FFD700; padding: 5px; border-radius: 3px; }")
        else:
            self.host_key_warning.setStyleSheet("QLabel { background-color: #404020; border: 1px solid #808040; padding: 5px; border-radius: 3px; }")

    def open_custom_names_dialog(self):
        dialog = CustomNamesDialog(self.settings, self)
        dialog.exec()

    def populate_actions_table(self):
        actions = self.settings.get("custom_actions", [])
        self.actions_table.setRowCount(len(actions))
        for row, action in enumerate(actions):
            self.actions_table.setItem(row, 0, QTableWidgetItem(action.get("name", "")))
            self.actions_table.setItem(row, 1, QTableWidgetItem(action.get("tooltip", "")))
            self.actions_table.setItem(row, 2, QTableWidgetItem(action.get("key", "")))

    def add_action(self):
        dialog = CustomActionDialog(parent=self)
        if dialog.exec():
            action = dialog.get_action()
            if action:
                row = self.actions_table.rowCount()
                self.actions_table.insertRow(row)
                self.actions_table.setItem(row, 0, QTableWidgetItem(action["name"]))
                self.actions_table.setItem(row, 1, QTableWidgetItem(action["tooltip"]))
                self.actions_table.setItem(row, 2, QTableWidgetItem(action["key"]))

    def edit_action(self):
        current_row = self.actions_table.currentRow()
        if current_row < 0:
            return
        
        action_data = {
            "name": self.actions_table.item(current_row, 0).text(),
            "tooltip": self.actions_table.item(current_row, 1).text(),
            "key": self.actions_table.item(current_row, 2).text(),
        }

        dialog = CustomActionDialog(action=action_data, parent=self)
        if dialog.exec():
            new_action = dialog.get_action()
            if new_action:
                self.actions_table.setItem(current_row, 0, QTableWidgetItem(new_action["name"]))
                self.actions_table.setItem(current_row, 1, QTableWidgetItem(new_action["tooltip"]))
                self.actions_table.setItem(current_row, 2, QTableWidgetItem(new_action["key"]))

    def remove_action(self):
        current_row = self.actions_table.currentRow()
        if current_row >= 0:
            self.actions_table.removeRow(current_row)

    def browse_vbox_path(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select VirtualBox Executable", os.path.dirname(self.vbox_path_edit.text()), "VirtualBox Executable (VirtualBox.exe);;All Executable Files (*.exe)")
        if path:
            self.vbox_path_edit.setText(path)
            self.vbox_path_edit.setCursorPosition(0)
            
    def get_settings(self):
        s = self.settings 
        s.update({
            "auto_attach": self.auto_attach_checkbox.isChecked(),
            "attach_vbox_manager": self.attach_vbox_manager_checkbox.isChecked(),
            "use_custom_names": self.use_custom_names_map_checkbox.isChecked(),
            "refresh_interval": self.refresh_interval_spinbox.value(),
            "vbox_path": self.vbox_path_edit.text(),
            "theme": self.theme_combo.currentText(),
            "toolbar_button_size": self.toolbar_size_spinbox.value(),
            "dpi_scaling": self.dpi_scaling_combo.currentText(),
        })
        if QTAWESOME_AVAILABLE:
            s["icon_theme"] = self.icon_theme_combo.currentText()
        
        custom_actions = []
        for row in range(self.actions_table.rowCount()):
            custom_actions.append({
                "name": self.actions_table.item(row, 0).text(),
                "tooltip": self.actions_table.item(row, 1).text(),
                "key": self.actions_table.item(row, 2).text(),
            })
        s["custom_actions"] = custom_actions
        return s

class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About VBoxTabs Manager")
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)
        
        top_level_layout = QVBoxLayout(self)
        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        icon_label = QLabel()
        pixmap = QApplication.style().standardIcon(QStyle.SP_ComputerIcon).pixmap(64, 64)
        icon_label.setPixmap(pixmap)
        main_layout.addWidget(icon_label, 0, Qt.AlignTop)

        text_layout = QVBoxLayout()
        text_layout.setSpacing(5)
        title_label = QLabel("VBoxTabs Manager")
        title_font = self.font()
        title_font.setPointSize(16); title_font.setBold(True)
        title_label.setFont(title_font)
        text_layout.addWidget(title_label)
        version_label = QLabel(f"Version {VERSION}"); text_layout.addWidget(version_label)
        desc_label = QLabel("A tabbed window manager for VirtualBox and more."); text_layout.addWidget(desc_label)
        text_layout.addSpacing(10)
        author_label = QLabel("Copyright (c) 2025 Zalexanninev15"); text_layout.addWidget(author_label)
        github_label = QLabel('<a href="https://github.com/Zalexanninev15/VBoxTabs-Manager">GitHub Repository</a>')
        github_label.setOpenExternalLinks(True)
        text_layout.addWidget(github_label)
        license_label = QLabel("Licensed under the MIT License."); text_layout.addWidget(license_label)
        text_layout.addStretch()
        main_layout.addLayout(text_layout)
        
        button_box = QHBoxLayout()
        button_box.addStretch()
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        ok_button.setDefault(True)
        button_box.addWidget(ok_button)
        
        top_level_layout.addLayout(main_layout)
        top_level_layout.addLayout(button_box)
        self.setFixedSize(self.sizeHint())

class VmInfoDialog(QDialog):
    def __init__(self, registered_vm_name, settings, reload_icon=None, parent=None):
        super().__init__(parent)
        self.vm_name = registered_vm_name
        self.settings = settings
        self.setWindowTitle(f"Info: {self.vm_name}"); self.setMinimumSize(700, 500)
        
        layout = QVBoxLayout(self)
        
        search_layout = QHBoxLayout()
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search for parameter or value...")
        self.search_box.textChanged.connect(self.filter_table)
        self.search_box.setClearButtonEnabled(True)
        search_layout.addWidget(self.search_box)
        
        self.reload_button = QToolButton()
        if reload_icon:
            self.reload_button.setIcon(reload_icon)
        else:
            self.reload_button.setIcon(QApplication.style().standardIcon(QStyle.SP_BrowserReload))
        self.reload_button.setToolTip("Reload VM Information")
        self.reload_button.clicked.connect(self.fetch_info)
        search_layout.addWidget(self.reload_button)
        
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Parameter", "Value"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSectionsMovable(True)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_table_context_menu)
        
        layout.addWidget(QLabel(f"Showing information for registered VM: <b>{self.vm_name}</b>"))
        layout.addLayout(search_layout)
        layout.addWidget(self.table)
        
        QTimer.singleShot(50, self.fetch_info)

    def show_table_context_menu(self, pos):
        item = self.table.itemAt(pos)
        if not item: return
        menu = QMenu(self)
        copy_cell_action = menu.addAction("Copy Cell Text")
        copy_row_action = menu.addAction("Copy Row")
        action = menu.exec(self.table.mapToGlobal(pos))
        if action == copy_cell_action:
            QApplication.clipboard().setText(item.text())
        elif action == copy_row_action:
            row = item.row()
            param = self.table.item(row, 0).text() if self.table.item(row, 0) else ""
            value = self.table.item(row, 1).text() if self.table.item(row, 1) else ""
            QApplication.clipboard().setText(f"{param}: {value}")

    def filter_table(self, text):
        for i in range(self.table.rowCount()):
            item1 = self.table.item(i, 0)
            item2 = self.table.item(i, 1)
            match = False
            if item1 and item2:
                match = (text.lower() in item1.text().lower() or
                         text.lower() in item2.text().lower())
            self.table.setRowHidden(i, not match)

    def fetch_info(self):
        vbox_manage_path = os.path.join(os.path.dirname(self.settings.get("vbox_path", "")), "VBoxManage.exe")
        if not os.path.exists(vbox_manage_path):
            QMessageBox.critical(self, "Error", f"VBoxManage.exe not found.\nExpected at: {vbox_manage_path}"); return
        try:
            self.table.setRowCount(0)
            result = subprocess.run(
                [vbox_manage_path, "showvminfo", self.vm_name, "--machinereadable"],
                capture_output=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW if IS_WINDOWS else 0,
                encoding='utf-8', errors='replace'
            )
            self.table.setSortingEnabled(False)
            for line in result.stdout.splitlines():
                match = re.match(r'([^=]+)="([^"]*)"', line)
                if match:
                    key, value = match.groups()
                    row_position = self.table.rowCount()
                    self.table.insertRow(row_position)
                    self.table.setItem(row_position, 0, QTableWidgetItem(key))
                    self.table.setItem(row_position, 1, QTableWidgetItem(value))
            self.table.setSortingEnabled(True)
        except subprocess.CalledProcessError as e:
            error_output = e.stderr.decode('utf-8', errors='replace') if e.stderr else str(e)
            QMessageBox.critical(self, "Error", f"Error fetching VM info:\n\n{error_output}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred:\n\n{str(e)}")

class VBoxTab(QWidget):
    def __init__(self, window_info, window_manager, parent=None):
        super().__init__(parent)
        self.window_manager = window_manager; self.window_info = window_info
        self.hwnd = window_info['hwnd']
        self.orig_win_details = None; self.attached = False; self.detached_manually = False
        layout = QVBoxLayout(self); layout.setContentsMargins(0, 0, 0, 0)
        self.container = QWidget(self); layout.addWidget(self.container)
    def attach_window(self):
        if self.attached or not self.window_manager.is_window(self.hwnd): return
        self.orig_win_details = self.window_manager.set_window_parent(self.hwnd, int(self.container.winId()))
        self.resize_embedded_window(); self.attached = True; self.detached_manually = False
    def detach_window(self):
        if not self.attached or self.orig_win_details is None or not self.window_manager.is_window(self.hwnd): return
        self.window_manager.restore_window_style(self.hwnd, self.orig_win_details)
        self.attached = False; self.detached_manually = True
    def resizeEvent(self, event): super().resizeEvent(event); self.resize_embedded_window()
    def resize_embedded_window(self):
        if self.attached and self.window_manager.is_window(self.hwnd):
            if IS_WINDOWS: win32gui.MoveWindow(self.hwnd, 0, 0, self.container.width(), self.container.height(), True)

class MiddleClickCloseTabBar(QTabBar):
    middleCloseRequested = Signal(int)
    def __init__(self, parent=None): super().__init__(parent)
    def mouseReleaseEvent(self, event: QMouseEvent):
        if event.button() == Qt.MouseButton.MiddleButton:
            tab_index = self.tabAt(event.position().toPoint())
            if tab_index != -1: self.middleCloseRequested.emit(tab_index); event.accept(); return
        super().mouseReleaseEvent(event)

class VirtualBoxTabs(QMainWindow):
    def __init__(self, window_manager, settings):
        super().__init__()
        self.window_manager = window_manager
        self.settings = settings
        self.settings_file = get_settings_path()
        
        self.manually_detached_windows = set()
        self.custom_tab_names = {} 
        self.is_picking_window = False
        self.picker_timer = QTimer(self); self.picker_timer.setInterval(50); self.picker_timer.timeout.connect(self._on_picker_tick)
        
        self.theme_map = {"Dark": "windows11", "Light": "windowsvista", "Classic": "Windows", "Fusion": "Fusion", "QDark": "qdarkstyle"}
        self.icons = {}
        self.toolbar_buttons = []
        
        self.setWindowTitle("VBoxTabs Manager"); self.resize(1280, 800); self.setMinimumSize(640, 480)
        central_widget = QWidget(); self.setCentralWidget(central_widget); self.main_layout = QVBoxLayout(central_widget)
        
        self.button_layout = QHBoxLayout()
        self.main_layout.addLayout(self.button_layout)

        self.tab_widget = QTabWidget(); custom_tab_bar = MiddleClickCloseTabBar(self.tab_widget); custom_tab_bar.middleCloseRequested.connect(self.close_tab_by_index)
        self.tab_widget.setTabBar(custom_tab_bar); self.tab_widget.setMovable(True); self.tab_widget.setTabsClosable(False); self.main_layout.addWidget(self.tab_widget)
        self.tabs = {}; self.auto_refresh_timer = QTimer(self); self.auto_refresh_timer.timeout.connect(self.refresh_tabs); self.auto_refresh_timer.start(self.settings.get("refresh_interval", 5) * 1000)
        
        self._rebuild_toolbar()
        self.refresh_tabs()
        
        self.tab_widget.tabBar().setContextMenuPolicy(Qt.CustomContextMenu); self.tab_widget.tabBar().customContextMenuRequested.connect(self.show_tab_context_menu)
        self.tab_widget.currentChanged.connect(self.update_tool_buttons)

    def _clear_layout(self, layout):
        if layout is None: return
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
            else:
                layout_item = item.layout()
                if layout_item is not None:
                    self._clear_layout(layout_item)

    def _rebuild_toolbar(self):
        self._clear_layout(self.button_layout)
        self.toolbar_buttons.clear()
        self.custom_actions_layout = None
        self._setup_toolbar()
        self.update_ui_from_settings()

    def _setup_toolbar(self):
        self.refresh_button = QToolButton(); self.refresh_button.setToolTip("Refresh VM list"); self.refresh_button.clicked.connect(self.refresh_tabs)
        self.attach_all_button = QToolButton(); self.attach_all_button.setToolTip("Attach all available VMs"); self.attach_all_button.clicked.connect(self.refresh_tabs)
        self.picker_button = QToolButton(); self.picker_button.setToolTip("Attach any window by clicking it"); self.picker_button.clicked.connect(self.start_window_picker)
        self.process_list_button = QToolButton(); self.process_list_button.setToolTip("Attach window from a process list"); self.process_list_button.clicked.connect(self.show_process_list_dialog)
        
        self.detach_button = QToolButton(); self.detach_button.setToolTip("Detach current VM"); self.detach_button.clicked.connect(self.detach_current_tab)
        self.close_window_button = QToolButton(); self.close_window_button.setToolTip("Close current VM window"); self.close_window_button.clicked.connect(self.close_current_window)
        self.close_all_button = QToolButton(); self.close_all_button.setToolTip("Close ALL VM windows"); self.close_all_button.clicked.connect(self.close_all_vms)
        
        self.rename_button = QToolButton(); self.rename_button.setToolTip("Rename current tab"); self.rename_button.clicked.connect(self.rename_current_tab)
        self.info_button = QToolButton(); self.info_button.setToolTip("Show current VM Information"); self.info_button.clicked.connect(self.show_current_vm_info)
        
        self.snapshot_button = QToolButton(); self.snapshot_button.setToolTip("Take a snapshot of the current VM"); self.snapshot_button.clicked.connect(self.take_snapshot)
        self.cad_button = QToolButton(); self.cad_button.setToolTip("Send Ctrl+Alt+Del to current VM"); self.cad_button.clicked.connect(self.send_ctrl_alt_del);
        self.power_button = QToolButton(); self.power_button.setToolTip("Power options for current VM"); self.power_button.setPopupMode(QToolButton.InstantPopup)
        
        self.vbox_main_button = QToolButton(); self.vbox_main_button.setToolTip("Open VirtualBox main application"); self.vbox_main_button.clicked.connect(self.open_virtualbox_main)
        self.settings_button = QToolButton(); self.settings_button.setToolTip("Settings"); self.settings_button.clicked.connect(self.show_settings_dialog)
        self.about_button = QToolButton(); self.about_button.setToolTip("About"); self.about_button.clicked.connect(self.show_about_dialog)
        
        self.toolbar_buttons = [self.refresh_button, self.attach_all_button, self.picker_button, self.process_list_button, self.detach_button, self.close_window_button, self.close_all_button, self.rename_button, self.info_button, self.snapshot_button, self.cad_button, self.power_button, self.vbox_main_button, self.settings_button, self.about_button]
        
        self.button_layout.addWidget(self.refresh_button); self.button_layout.addWidget(self.attach_all_button); self.button_layout.addWidget(self.picker_button); self.button_layout.addWidget(self.process_list_button)
        sep1 = QFrame(); sep1.setFrameShape(QFrame.VLine); sep1.setFrameShadow(QFrame.Sunken); self.button_layout.addWidget(sep1)
        
        self.button_layout.addWidget(self.detach_button); self.button_layout.addWidget(self.close_window_button); self.button_layout.addWidget(self.close_all_button)
        sep2 = QFrame(); sep2.setFrameShape(QFrame.VLine); sep2.setFrameShadow(QFrame.Sunken); self.button_layout.addWidget(sep2)

        self.button_layout.addWidget(self.rename_button); self.button_layout.addWidget(self.info_button)
        self.button_layout.addWidget(self.snapshot_button); self.button_layout.addWidget(self.cad_button); self.button_layout.addWidget(self.power_button)
        
        sep_custom = QFrame(); sep_custom.setFrameShape(QFrame.VLine); sep_custom.setFrameShadow(QFrame.Sunken); self.button_layout.addWidget(sep_custom)
        self.custom_actions_layout = QHBoxLayout(); self.custom_actions_layout.setSpacing(2)
        self.button_layout.addLayout(self.custom_actions_layout)
        
        sep_main = QFrame(); sep_main.setFrameShape(QFrame.VLine); sep_main.setFrameShadow(QFrame.Sunken); self.button_layout.addWidget(sep_main)
        self.button_layout.addWidget(self.vbox_main_button); self.button_layout.addStretch(1)
        self.button_layout.addWidget(self.settings_button); self.button_layout.addWidget(self.about_button)

    def update_ui_from_settings(self):
        self._setup_icons()
        self._setup_custom_actions()
        self._apply_toolbar_size()
        self.update_tool_buttons(self.tab_widget.currentIndex())

    def _get_icon(self, awesome_name, standard_enum):
        use_awesome = self.settings.get("icon_theme") == "Awesome" and QTAWESOME_AVAILABLE
        if use_awesome:
            theme_name = self.settings.get("theme", "Fusion")
            is_light_mode = (theme_name == "Light") or ((theme_name == "Fusion" or theme_name == "Classic") and is_windows_light_theme())
            icon_color = '#404040' if is_light_mode else '#d0d0d0'
            return qta.icon(awesome_name, color=icon_color)
        return QApplication.style().standardIcon(standard_enum)

    def _setup_icons(self):
        self.icons = {
            'refresh': self._get_icon('fa5s.sync-alt', QStyle.SP_BrowserReload),
            'attach_all': self._get_icon('fa5s.link', QStyle.SP_ArrowUp),
            'picker': self._get_icon('fa5s.crosshairs', QStyle.SP_DesktopIcon),
            'process_list': self._get_icon('fa5s.list-ul', QStyle.SP_FileDialogListView),
            'detach': self._get_icon('fa5s.unlink', QStyle.SP_DialogCancelButton),
            'close': self._get_icon('fa5s.times-circle', QStyle.SP_DialogResetButton),
            'close_all': self._get_icon('fa5s.times', QStyle.SP_DialogCloseButton),
            'rename': self._get_icon('fa5s.edit', QStyle.SP_FileDialogNewFolder),
            'info': self._get_icon('fa5s.info-circle', QStyle.SP_MessageBoxInformation),
            'snapshot': self._get_icon('fa5s.camera-retro', QStyle.SP_DialogSaveButton),
            'cad': self._get_icon('fa5s.keyboard', QStyle.SP_DialogOkButton),
            'power': self._get_icon('fa5s.power-off', QStyle.SP_VistaShield),
            'vbox': self._get_icon('fa5s.desktop', QStyle.SP_ComputerIcon),
            'settings': self._get_icon('fa5s.cog', QStyle.SP_FileDialogDetailedView),
            'about': self._get_icon('fa5s.question-circle', QStyle.SP_MessageBoxInformation),
            'suspend': self._get_icon('fa5s.pause', QStyle.SP_BrowserStop),
            'restart': self._get_icon('fa5s.redo', QStyle.SP_DialogYesButton),
            'shutdown': self._get_icon('mdi.lightning-bolt', QStyle.SP_DialogNoButton),
            'custom_action': self._get_icon('fa5s.bolt', QStyle.SP_CommandLink)
        }
        
        self.refresh_button.setIcon(self.icons['refresh']); self.attach_all_button.setIcon(self.icons['attach_all'])
        self.picker_button.setIcon(self.icons['picker']); self.process_list_button.setIcon(self.icons['process_list'])
        self.detach_button.setIcon(self.icons['detach'])
        self.close_window_button.setIcon(self.icons['close']); self.close_all_button.setIcon(self.icons['close_all'])
        self.rename_button.setIcon(self.icons['rename']); self.info_button.setIcon(self.icons['info'])
        self.snapshot_button.setIcon(self.icons['snapshot']); self.cad_button.setIcon(self.icons['cad'])
        self.power_button.setIcon(self.icons['power']); self.vbox_main_button.setIcon(self.icons['vbox'])
        self.settings_button.setIcon(self.icons['settings']); self.about_button.setIcon(self.icons['about'])

        power_menu = QMenu(self)
        power_menu.addAction(self.icons['suspend'], "Suspend", self.suspend_vm)
        power_menu.addAction(self.icons['restart'], "Restart", self.restart_vm)
        power_menu.addAction(self.icons['shutdown'], "Shutdown", self.shutdown_vm)
        self.power_button.setMenu(power_menu)
        
        self.setWindowIcon(self.icons['vbox'])
    
    def _apply_toolbar_size(self):
        size = self.settings.get("toolbar_button_size", 16)
        icon_size = QSize(size, size)
        for button in self.toolbar_buttons:
            button.setIconSize(icon_size)

    def _setup_custom_actions(self):
        self._clear_layout(self.custom_actions_layout)
        
        actions = self.settings.get("custom_actions", [])
        for action in actions:
            key = action.get("key")
            if not key: continue
            
            button = QToolButton()
            button.setIcon(self.icons['custom_action'])
            tooltip_text = f"{action.get('tooltip', '')} (Right Ctrl+{key})"
            button.setToolTip(tooltip_text)
            button.clicked.connect(partial(self._execute_custom_action, key))
            
            self.custom_actions_layout.addWidget(button)
            self.toolbar_buttons.append(button)

    def _execute_custom_action(self, key):
        key = key.upper()
        vm_name, error = self._get_current_vm_name()
        if error:
            QMessageBox.warning(self, "Action Error", error)
            return

        if key == 'E':
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop', 'VBoxTabs Screenshots')
            os.makedirs(desktop_path, exist_ok=True)
            file_path = os.path.join(desktop_path, f"{vm_name}_Screenshot_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.png")
            command_args = ["screenshotpng", file_path]
            self._run_vboxmanage_command("controlvm", command_args)
            return
        
        scancode_press = SCANCODE_MAP.get(key)
        if not scancode_press:
            QMessageBox.warning(self, "Action Error", f"Invalid key '{key}' for custom action.")
            return

        scancode_release = f"{int(scancode_press, 16) | 0x80:02x}"
        command_args = ["keyboardputscancode", "e0", "1d", scancode_press, scancode_release, "e0", "9d"]
        self._run_vboxmanage_command("controlvm", command_args)

    def update_tool_buttons(self, index):
        has_tabs = index != -1
        is_vbox_tab = False
        if has_tabs:
            tab = self.tab_widget.widget(index)
            if isinstance(tab, VBoxTab) and tab.window_info.get('type') == 'VirtualBox':
                is_vbox_tab = True
        
        self.detach_button.setEnabled(has_tabs)
        self.close_window_button.setEnabled(has_tabs)
        self.close_all_button.setEnabled(has_tabs)
        self.rename_button.setEnabled(has_tabs)

        self.info_button.setEnabled(is_vbox_tab)
        self.snapshot_button.setEnabled(is_vbox_tab)
        self.cad_button.setEnabled(is_vbox_tab)
        self.power_button.setEnabled(is_vbox_tab)
        
        if self.custom_actions_layout:
            for i in range(self.custom_actions_layout.count()):
                widget = self.custom_actions_layout.itemAt(i).widget()
                if widget:
                    widget.setEnabled(is_vbox_tab)

    def save_settings(self):
        try:
            with open(self.settings_file, 'w') as f: json.dump(self.settings, f, indent=4)
        except Exception as e: QMessageBox.warning(self, "Error", f"Failed to save settings: {str(e)}")

    def show_settings_dialog(self):
        dialog = SettingsDialog(self.settings.copy(), self)
        if dialog.exec():
            new_settings = dialog.get_settings()
            old_settings = self.settings.copy()
            self.settings = new_settings
            self.save_settings()

            if new_settings["refresh_interval"] != old_settings["refresh_interval"]:
                self.auto_refresh_timer.start(new_settings["refresh_interval"] * 1000)
            
            if new_settings["theme"] != old_settings["theme"] or \
               new_settings["icon_theme"] != old_settings.get("icon_theme"):
                QTimer.singleShot(0, lambda: self.change_theme(new_settings["theme"]))

            if new_settings["dpi_scaling"] != old_settings["dpi_scaling"] or \
               new_settings["toolbar_button_size"] != old_settings["toolbar_button_size"]:
                QMessageBox.information(self, "Restart Required", 
                    "DPI Scaling and Toolbar Icon Size changes will take full effect after restarting the application.")

            if new_settings["custom_names"] != old_settings["custom_names"] or \
               new_settings["use_custom_names"] != old_settings["use_custom_names"]:
                for i in range(self.tab_widget.count()):
                    self._update_tab_text(i)
    
    def show_process_list_dialog(self):
        if not IS_WINDOWS:
            QMessageBox.information(self, "Not Supported", "This feature is only available on Windows.")
            return
        
        dialog = ProcessListDialog(self.window_manager, int(self.winId()), reload_icon=self.icons['refresh'], parent=self)
        if dialog.exec():
            window_info = dialog.get_selected_window_info()
            if window_info:
                self.add_tab_for_window(window_info, force_attach=True)

    def change_theme(self, theme_name):
        app = QApplication.instance()
        app.setStyleSheet("")
        app.setPalette(QPalette())
        
        theme_key = self.theme_map.get(theme_name)
        
        # This code allows you to bypass the Windows window title bug, as well as bypass the bug with the Fusion dark theme and the appearance of phantom places for icons in lists.
        if IS_WINDOWS and theme_name == "QDark":
            app.setStyle(QStyleFactory.create("windows11"))

        if theme_key == "qdarkstyle" and QDARKSTYLE_AVAILABLE:
            app.setStyleSheet(qdarkstyle.load_stylesheet(qt_api='pyside6'))
        elif theme_key:
            app.setStyle(QStyleFactory.create(theme_key))
        else:
            app.setStyle(QStyleFactory.create("Fusion"))
        
        self._rebuild_toolbar()

    def show_about_dialog(self): AboutDialog(self).exec()
    def start_window_picker(self):
        if not IS_WINDOWS: QMessageBox.information(self, "Not Supported", "Window picking is only supported on Windows."); return
        self.is_picking_window = True; QApplication.setOverrideCursor(Qt.CrossCursor); self.picker_timer.start()
    def _cancel_picker(self): self.is_picking_window = False; self.picker_timer.stop(); QApplication.restoreOverrideCursor()
    def keyPressEvent(self, event):
        if self.is_picking_window and event.key() == Qt.Key_Escape: self._cancel_picker()
        super().keyPressEvent(event)
    def _on_picker_tick(self):
        if not (win32api.GetKeyState(0x01) & 0x8000): return
        self._cancel_picker()
        pos = win32api.GetCursorPos()
        hwnd = self.window_manager.get_window_from_point(pos)
        if hwnd:
            toplevel_hwnd = self.window_manager.get_toplevel_parent(hwnd)
            if toplevel_hwnd != int(self.winId()):
                title = self.window_manager.get_window_title(toplevel_hwnd) or f"Window {toplevel_hwnd}"
                self.add_tab_for_window({'hwnd': toplevel_hwnd, 'title': title, 'type': 'Picked'}, force_attach=True)

    def add_tab_for_window(self, window_info, force_attach=False):
        hwnd = window_info['hwnd']
        if hwnd in self.tabs:
            for i in range(self.tab_widget.count()):
                if self.tab_widget.widget(i).hwnd == hwnd: self.tab_widget.setCurrentIndex(i); return
        
        window_info['original_title'] = window_info['title']
        
        session_name = self.custom_tab_names.get(hwnd)
        
        if session_name:
            window_info['title'] = session_name
        elif self.settings.get("use_custom_names", False):
            persistent_name = self.settings.get("custom_names", {}).get(window_info['original_title'])
            if persistent_name:
                window_info['title'] = persistent_name

        win_type = window_info.get('type', 'Unknown')
        is_manual_attach_all = hasattr(self, 'attach_all_button') and self.sender() == self.attach_all_button
        was_manually_detached = hwnd in self.manually_detached_windows
        
        auto_attach_enabled = False
        if win_type == 'VirtualBox':
            auto_attach_enabled = self.settings.get("auto_attach", True)
        elif win_type == 'VirtualBox Manager':
            auto_attach_enabled = self.settings.get("attach_vbox_manager", False)
        
        should_attach = force_attach or is_manual_attach_all or (auto_attach_enabled and not was_manually_detached)
        
        display_title = window_info['title']
        if win_type not in ['VirtualBox', 'VirtualBox Manager'] and len(display_title) > 20:
            display_title = display_title[:20] + '...'
        window_info['display_title'] = display_title

        tab = VBoxTab(window_info, self.window_manager)
        self.tabs[hwnd] = tab
        index = self.tab_widget.addTab(tab, "...") 
        
        self.tab_widget.setTabToolTip(index, f"Type: {win_type}\nOriginal: {window_info['original_title']}\nHandle: {hwnd}")
        self._update_tab_text(index)
        
        if win_type == 'VirtualBox':
            registered_name, _ = self._get_registered_vm_name(window_info['original_title'])
            if registered_name:
                snapshot_name = self._get_latest_snapshot_name(registered_name)
                if snapshot_name:
                    tab.window_info['snapshot_name'] = snapshot_name
                    self._update_tab_text(index)

        if should_attach:
            if was_manually_detached: self.manually_detached_windows.remove(hwnd)
            tab.attach_window()
        
        self.tab_widget.setCurrentIndex(index)
    
    def _update_tab_text(self, index):
        tab = self.tab_widget.widget(index)
        if not tab: return
        
        info = tab.window_info
        base_title = info.get('title', 'Tab')
        win_type = info.get('type', 'Unknown')
        snapshot_name = info.get('snapshot_name')
        
        final_text = base_title
        if win_type not in ['VirtualBox', 'VirtualBox Manager']:
            final_text = f"[{win_type}] {base_title}"
        
        if snapshot_name and win_type == 'VirtualBox':
            clean_base_title = re.sub(r" \([^)]*\)$", "", base_title).strip()
            final_text = f"{clean_base_title} ({snapshot_name})"
        
        self.tab_widget.setTabText(index, final_text)

    def refresh_tabs(self):
        is_attach_all_request = hasattr(self, 'attach_all_button') and self.sender() == self.attach_all_button
        found_windows = self.window_manager.find_embeddable_windows()
        
        current_hwnds = {w['hwnd'] for w in found_windows}
        for hwnd, tab in list(self.tabs.items()):
            if hwnd not in current_hwnds and not self.window_manager.is_window(hwnd):
                self.cleanup_closed_tab(hwnd)
        
        for window_info in found_windows:
            hwnd = window_info['hwnd']
            if hwnd in self.tabs:
                if is_attach_all_request and hwnd in self.manually_detached_windows:
                    self.manually_detached_windows.remove(hwnd)
                    tab = self.tabs[hwnd]
                    if not tab.attached:
                        tab.attach_window()
                continue

            if hwnd in self.manually_detached_windows and not is_attach_all_request:
                continue
            
            if window_info.get('type') == 'VirtualBox Manager' and not self.settings.get('attach_vbox_manager', False) and not is_attach_all_request:
                continue

            self.add_tab_for_window(window_info, force_attach=is_attach_all_request)
        self.update_tool_buttons(self.tab_widget.currentIndex())

    def detach_tab(self, index):
        if not (0 <= index < self.tab_widget.count()): return
        
        tab_to_detach = self.tab_widget.widget(index)
        if not isinstance(tab_to_detach, VBoxTab): return

        hwnd = tab_to_detach.hwnd
        tab_to_detach.detach_window()
        self.manually_detached_windows.add(hwnd)
        
        self.tab_widget.removeTab(index)
        if hwnd in self.tabs:
            del self.tabs[hwnd]
        tab_to_detach.deleteLater()

    def detach_current_tab(self): self.detach_tab(self.tab_widget.currentIndex())
    
    def close_all_vms(self):
        if self.tab_widget.count() == 0:
            return
        
        reply = QMessageBox.question(self, "Close All VMs",
                                     "Are you sure you want to forcefully close all VM windows?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.No:
            return
        
        for i in range(self.tab_widget.count() - 1, -1, -1):
            self.close_tab_by_index(i)
        
        QTimer.singleShot(500, self.refresh_tabs)

    def close_tab_by_index(self, index):
        if not (0 <= index < self.tab_widget.count()): return
        tab = self.tab_widget.widget(index); hwnd = tab.hwnd
        pid = self.window_manager.get_process_id(hwnd)
        if pid and self.window_manager.terminate_process_by_id(pid):
            QTimer.singleShot(200, lambda: self.cleanup_closed_tab(hwnd))
        else: self.cleanup_closed_tab(hwnd)
        
    def cleanup_closed_tab(self, hwnd):
        if hwnd in self.tabs:
            tab = self.tabs[hwnd]; index = self.tab_widget.indexOf(tab)
            if index != -1: self.tab_widget.removeTab(index)
            del self.tabs[hwnd]; tab.deleteLater()
            
            if hwnd in self.manually_detached_windows:
                self.manually_detached_windows.remove(hwnd)
            if hwnd in self.custom_tab_names:
                del self.custom_tab_names[hwnd]

    def close_current_window(self): self.close_tab_by_index(self.tab_widget.currentIndex())
    def closeEvent(self, event):
        for tab in self.tabs.values(): tab.detach_window()
        self.save_settings()
        event.accept()
    def open_virtualbox_main(self):
        path = self.settings.get("vbox_path")
        if path and os.path.exists(path):
            try: subprocess.Popen([path])
            except Exception as e: QMessageBox.warning(self, "Error", f"Failed to open VirtualBox: {str(e)}")
        else: QMessageBox.warning(self, "Error", f"VirtualBox executable not found.\nPlease check path in Settings:\n{path}")
        
    def rename_current_tab(self):
        index = self.tab_widget.currentIndex()
        if index >= 0: self._rename_tab_at_index(index)
        
    def _rename_tab_at_index(self, index):
        tab = self.tab_widget.widget(index)
        current_title = self.tab_widget.tabText(index)
        
        base_name_for_edit = re.sub(r" \([^)]*\)$", "", current_title).strip()
        
        new_name, ok = QInputDialog.getText(self, "Rename Tab", "Enter new name:", text=base_name_for_edit)
        
        if ok and new_name:
            hwnd = tab.window_info['hwnd']
            original_title = tab.window_info['original_title']
            
            tab.window_info['title'] = new_name
            self.custom_tab_names[hwnd] = new_name
            self.settings['custom_names_map'][original_title] = new_name
            self.save_settings()
            
            self._update_tab_text(index)

    def _get_current_vm_name(self):
        index = self.tab_widget.currentIndex()
        if index < 0: return None, "No tab selected."
        tab = self.tab_widget.widget(index)
        if not isinstance(tab, VBoxTab) or tab.window_info.get('type') != 'VirtualBox':
            return None, "Selected tab is not a VirtualBox VM."
        return self._get_registered_vm_name(tab.window_info['original_title'])

    def _get_registered_vm_name(self, window_title_part):
        vbox_manage_path = os.path.join(os.path.dirname(self.settings.get("vbox_path", "")), "VBoxManage.exe")
        if not os.path.exists(vbox_manage_path): return None, "VBoxManage.exe not found"
        try:
            result = subprocess.run(
                [vbox_manage_path, "list", "runningvms"],
                capture_output=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW if IS_WINDOWS else 0,
                encoding='utf-8', errors='replace'
            )
            for line in result.stdout.splitlines():
                match = re.match(r'"([^"]+)"', line)
                if match:
                    registered_name = match.group(1)
                    if window_title_part.startswith(registered_name): return registered_name, None
            return None, f"No running VM found matching '{window_title_part}'"
        except Exception as e: return None, f"Failed to list running VMs: {str(e)}"

    def _get_latest_snapshot_name(self, vm_name):
        vbox_manage_path = os.path.join(os.path.dirname(self.settings.get("vbox_path", "")), "VBoxManage.exe")
        if not os.path.exists(vbox_manage_path): return None
        try:
            result = subprocess.run(
                [vbox_manage_path, "showvminfo", vm_name, "--machinereadable"],
                capture_output=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW if IS_WINDOWS else 0,
                encoding='utf-8', errors='replace'
            )
            for line in result.stdout.splitlines():
                if line.startswith('currentsnapshotname='):
                    match = re.match(r'[^=]+="([^"]*)"', line)
                    if match and match.group(1):
                        return match.group(1)
            return None
        except Exception:
            return None

    def _run_vboxmanage_command(self, base_command, command_args, show_error=True):
        vm_name, error = self._get_current_vm_name()
        if error:
            if show_error: QMessageBox.warning(self, "Command Error", error)
            return False, None
        
        vbox_manage_path = os.path.join(os.path.dirname(self.settings.get("vbox_path", "")), "VBoxManage.exe")
        
        full_command = [vbox_manage_path, base_command, vm_name] + command_args
        
        try:
            subprocess.run(full_command, check=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW if IS_WINDOWS else 0)
            return True, vm_name
        except subprocess.CalledProcessError as e:
            if show_error:
                error_output = e.stderr.decode('utf-8', errors='replace') if e.stderr else str(e)
                QMessageBox.critical(self, "VBoxManage Error", f"Failed to execute command:\n{' '.join(full_command)}\n\n{error_output}")
            return False, vm_name
        except Exception as e:
            if show_error:
                QMessageBox.critical(self, "Execution Error", f"An unexpected error occurred:\n\n{str(e)}")
            return False, vm_name

    def take_snapshot(self):
        vm_name, error = self._get_current_vm_name()
        if error:
            QMessageBox.warning(self, "Snapshot Error", error)
            return

        default_name = f"Snapshot_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
        snap_name, ok = QInputDialog.getText(self, "Take Snapshot", f"Enter snapshot name for '{vm_name}':", text=default_name)
        
        if ok and snap_name:
            success, _ = self._run_vboxmanage_command("snapshot", ["take", snap_name])
            if success:
                index = self.tab_widget.currentIndex()
                if index != -1:
                    tab = self.tab_widget.widget(index)
                    tab.window_info['snapshot_name'] = snap_name
                    self._update_tab_text(index)

    def send_ctrl_alt_del(self):
        self._run_vboxmanage_command("controlvm", ["keyboardputscancode", "1d", "38", "e0", "53", "e0", "d3", "e0", "b8", "e0", "9d"])
    
    def suspend_vm(self): 
        success, _ = self._run_vboxmanage_command("controlvm", ["savestate"])
        if success:
            self.detach_current_tab()

    def restart_vm(self): 
        self._run_vboxmanage_command("controlvm", ["reset"])
            
    def shutdown_vm(self): 
        success, _ = self._run_vboxmanage_command("controlvm", ["acpipowerbutton"])
        if success:
            self.detach_current_tab()

    def show_current_vm_info(self):
        registered_name, error = self._get_current_vm_name()
        if registered_name:
            VmInfoDialog(registered_name, self.settings, reload_icon=self.icons['refresh'], parent=self).exec()

    def show_tab_context_menu(self, pos):
        tabBar = self.tab_widget.tabBar(); index = tabBar.tabAt(pos)
        if index < 0: return
        self.tab_widget.setCurrentIndex(index); tab = self.tab_widget.widget(index)
        menu = QMenu(self)
        rename_action = menu.addAction(self.icons['rename'], "Rename Tab")
        detach_action = menu.addAction(self.icons['detach'], "Detach Window")
        close_action = menu.addAction(self.icons['close'], "Close Window")
        
        info_action, snapshot_action = None, None
        if tab.window_info.get('type') == "VirtualBox":
            menu.addSeparator()
            info_action = menu.addAction(self.icons['info'], "Show VM Info")
            snapshot_action = menu.addAction(self.icons['snapshot'], "Take Snapshot")
        
        action = menu.exec(tabBar.mapToGlobal(pos))
        
        if action == rename_action: self._rename_tab_at_index(index)
        elif action == detach_action: self.detach_tab(index)
        elif action == close_action: self.close_tab_by_index(index)
        elif action == info_action: self.show_current_vm_info()
        elif action == snapshot_action: self.take_snapshot()

if __name__ == "__main__":
    settings_file = get_settings_path()
    settings = load_settings(settings_file)

    scaling = settings.get("dpi_scaling", "Auto")
    if scaling != "Auto":
        try:
            scale_factor = float(scaling.strip('%')) / 100.0
            os.environ['QT_SCALE_FACTOR'] = str(scale_factor)
            os.environ['QT_AUTO_SCREEN_SCALE_FACTOR'] = '0'
        except (ValueError, TypeError):
            os.environ['QT_AUTO_SCREEN_SCALE_FACTOR'] = '1'
    else:
        os.environ['QT_AUTO_SCREEN_SCALE_FACTOR'] = '1'

    app = QApplication(sys.argv)
    
    if IS_WINDOWS: window_manager = WindowsWindowManager()
    elif IS_LINUX: window_manager = LinuxWindowManager()
    else: QMessageBox.critical(None, "Unsupported OS", f"This application does not support {platform.system()}."); sys.exit(1)
    
    if not QDARKSTYLE_AVAILABLE: print("Warning: qdarkstyle not found. The 'QDark' theme may be unavailable.")
    if not QTAWESOME_AVAILABLE: print("Warning: qtawesome not found. 'Awesome' icons will be unavailable.")

    window = VirtualBoxTabs(window_manager, settings)
    window.show()
    
    QTimer.singleShot(50, lambda: window.change_theme(window.settings.get("theme", "Fusion")))
    
    sys.exit(app.exec())