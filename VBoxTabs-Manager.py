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
from PySide6.QtCore import Qt, QTimer, Signal, QPoint
from PySide6.QtGui import QFont, QAction, QIcon, QPixmap, QMouseEvent, QPalette

# QDarkStyle
try:
    import qdarkstyle
    QDARKSTYLE_AVAILABLE = True
except ImportError:
    QDARKSTYLE_AVAILABLE = False

# --- Helper Functions ---
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
            
        is_vbox_manager = "Oracle VM VirtualBox Manager" in window_title
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

# --- Application Classes ---
class SettingsDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("Settings")
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15); main_layout.setSpacing(10)
        
        general_group = QGroupBox("General Settings")
        general_layout = QGridLayout(general_group)

        self.auto_attach_checkbox = QCheckBox("Automatically attach VirtualBox windows")
        self.auto_attach_checkbox.setChecked(self.settings.get("auto_attach", True))
        general_layout.addWidget(self.auto_attach_checkbox, 0, 0, 1, 2)
        
        interval_label = QLabel("Window detection interval (seconds):")
        general_layout.addWidget(interval_label, 1, 0)
        
        self.refresh_interval_spinbox = QSpinBox()
        self.refresh_interval_spinbox.setRange(1, 60)
        self.refresh_interval_spinbox.setValue(self.settings.get("refresh_interval", 5))
        general_layout.addWidget(self.refresh_interval_spinbox, 1, 1)
        
        path_label = QLabel("VirtualBox executable path:")
        general_layout.addWidget(path_label, 2, 0)

        vbox_path_layout = QHBoxLayout()
        self.vbox_path_edit = QLineEdit(self.settings.get("vbox_path", r"C:\Program Files\Oracle\VirtualBox\VirtualBox.exe"))
        self.vbox_path_edit.setCursorPosition(0)
        vbox_path_layout.addWidget(self.vbox_path_edit)
        self.browse_button = QPushButton("Browse...")
        self.browse_button.clicked.connect(self.browse_vbox_path)
        vbox_path_layout.addWidget(self.browse_button)
        general_layout.addLayout(vbox_path_layout, 2, 1)
        general_layout.setColumnStretch(1, 1)
        main_layout.addWidget(general_group)

        display_group = QGroupBox("Display Settings")
        display_layout = QGridLayout(display_group)
        theme_label = QLabel("Theme:")
        display_layout.addWidget(theme_label, 0, 0)
        
        self.theme_combo = QComboBox()
        theme_names = []
        if IS_WINDOWS and not is_windows_light_theme():
            theme_names.append("Dark")
        theme_names.extend(["Light", "Classic", "Fusion"])
        if QDARKSTYLE_AVAILABLE: theme_names.append("QDark")
        
        self.theme_combo.addItems(theme_names)
        current_theme = self.settings.get("theme", "Fusion")
        if current_theme in theme_names:
            self.theme_combo.setCurrentText(current_theme)
        else:
            self.theme_combo.setCurrentText("Fusion" if "Fusion" in theme_names else (theme_names[0] if theme_names else ""))
        display_layout.addWidget(self.theme_combo, 0, 1)

        dpi_label = QLabel("DPI Scaling:")
        display_layout.addWidget(dpi_label, 1, 0)
        self.dpi_scaling_combo = QComboBox()
        self.dpi_scaling_combo.addItems(["Auto", "100%", "125%", "150%", "175%", "200%"])
        self.dpi_scaling_combo.setCurrentText(self.settings.get("dpi_scaling", "Auto"))
        display_layout.addWidget(self.dpi_scaling_combo, 1, 1)
        display_layout.setColumnStretch(1, 1)
        main_layout.addWidget(display_group)

        button_layout = QHBoxLayout()
        button_layout.addStretch(1)
        self.cancel_button = QPushButton("Cancel"); self.cancel_button.clicked.connect(self.reject); button_layout.addWidget(self.cancel_button)
        self.save_button = QPushButton("Save"); self.save_button.clicked.connect(self.accept); button_layout.addWidget(self.save_button)
        main_layout.addLayout(button_layout)
        
        self.adjustSize()

    def browse_vbox_path(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select VirtualBox Executable", os.path.dirname(self.vbox_path_edit.text()), "VirtualBox Executable (VirtualBox.exe);;All Executable Files (*.exe)")
        if path:
            self.vbox_path_edit.setText(path)
            self.vbox_path_edit.setCursorPosition(0)
    def get_settings(self):
        return {"auto_attach": self.auto_attach_checkbox.isChecked(), "refresh_interval": self.refresh_interval_spinbox.value(), "vbox_path": self.vbox_path_edit.text(), "theme": self.theme_combo.currentText(), "dpi_scaling": self.dpi_scaling_combo.currentText()}

class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About VBoxTabs Manager")
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)
        main_layout = QHBoxLayout(self); main_layout.setContentsMargins(15, 15, 15, 15); main_layout.setSpacing(15)
        icon_label = QLabel(); pixmap = QApplication.style().standardIcon(QStyle.SP_ComputerIcon).pixmap(64, 64)
        icon_label.setPixmap(pixmap); main_layout.addWidget(icon_label, 0, Qt.AlignTop)
        text_layout = QVBoxLayout(); text_layout.setSpacing(5)
        title_label = QLabel("VBoxTabs Manager"); title_font = self.font(); title_font.setPointSize(16); title_font.setBold(True)
        title_label.setFont(title_font); text_layout.addWidget(title_label)
        version_label = QLabel("Version 2.0-beta 1"); text_layout.addWidget(version_label)
        desc_label = QLabel("A tabbed window manager for VirtualBox and more."); text_layout.addWidget(desc_label)
        text_layout.addSpacing(10)
        author_label = QLabel("Copyright (c) 2025 Zalexanninev15"); text_layout.addWidget(author_label)
        github_label = QLabel('<a href="https://github.com/Zalexanninev15/VBoxTabs-Manager">GitHub Repository</a>')
        github_label.setOpenExternalLinks(True); text_layout.addWidget(github_label)
        license_label = QLabel("Licensed under the MIT License."); text_layout.addWidget(license_label)
        text_layout.addStretch()
        main_layout.addLayout(text_layout, 1)
        final_layout = QVBoxLayout(); final_layout.addLayout(main_layout)
        button_box = QHBoxLayout(); button_box.addStretch()
        ok_button = QPushButton("OK"); ok_button.clicked.connect(self.accept); ok_button.setDefault(True)
        button_box.addWidget(ok_button); final_layout.addLayout(button_box)
        self.setLayout(final_layout); self.setFixedSize(self.sizeHint())

class VmInfoDialog(QDialog):
    def __init__(self, registered_vm_name, settings, parent=None):
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
    def __init__(self, window_manager):
        super().__init__()
        self.window_manager = window_manager
        self.manually_detached_windows = set()
        self.is_picking_window = False
        self.picker_timer = QTimer(self); self.picker_timer.setInterval(50); self.picker_timer.timeout.connect(self._on_picker_tick)
        
        self.theme_map = {
            "Dark": "windows11",
            "Light": "windowsvista",
            "Classic": "Windows",
            "Fusion": "Fusion",
            "QDark": "qdarkstyle",
        }

        self.settings_file = self.get_settings_path(); self.settings = self.load_settings()
        self.apply_dpi_scaling(); self.setWindowTitle("VBoxTabs Manager"); self.resize(1280, 800); self.setMinimumSize(640, 480)
        style = QApplication.style(); self.setWindowIcon(style.standardIcon(QStyle.SP_ComputerIcon))
        central_widget = QWidget(); self.setCentralWidget(central_widget); main_layout = QVBoxLayout(central_widget)
        button_layout = QHBoxLayout()
        
        # Icons
        refresh_icon = style.standardIcon(QStyle.SP_BrowserReload); attach_all_icon = style.standardIcon(QStyle.SP_ArrowUp); detach_icon = style.standardIcon(QStyle.SP_DialogCancelButton);
        picker_icon = style.standardIcon(QStyle.SP_DesktopIcon); rename_icon = style.standardIcon(QStyle.SP_FileDialogNewFolder); vbox_icon = style.standardIcon(QStyle.SP_ComputerIcon);
        close_icon = style.standardIcon(QStyle.SP_DialogResetButton); settings_icon = style.standardIcon(QStyle.SP_FileDialogDetailedView); about_icon = style.standardIcon(QStyle.SP_MessageBoxInformation)
        info_icon = style.standardIcon(QStyle.SP_MessageBoxInformation); close_all_icon = style.standardIcon(QStyle.SP_DialogCloseButton);
        snapshot_icon = style.standardIcon(QStyle.SP_DialogSaveButton); cad_icon = style.standardIcon(QStyle.SP_DialogOkButton); power_icon = style.standardIcon(QStyle.SP_VistaShield)

        self.refresh_button = QToolButton(); self.refresh_button.setIcon(refresh_icon); self.refresh_button.setToolTip("Refresh VM list"); self.refresh_button.clicked.connect(self.refresh_tabs); button_layout.addWidget(self.refresh_button)
        self.attach_all_button = QToolButton(); self.attach_all_button.setIcon(attach_all_icon); self.attach_all_button.setToolTip("Attach all available VMs"); self.attach_all_button.clicked.connect(self.refresh_tabs); button_layout.addWidget(self.attach_all_button)
        self.picker_button = QToolButton(); self.picker_button.setIcon(picker_icon); self.picker_button.setToolTip("Attach any window by clicking it"); self.picker_button.clicked.connect(self.start_window_picker); button_layout.addWidget(self.picker_button)
        
        separator1 = QFrame(); separator1.setFrameShape(QFrame.VLine); separator1.setFrameShadow(QFrame.Sunken); button_layout.addWidget(separator1)

        self.detach_button = QToolButton(); self.detach_button.setIcon(detach_icon); self.detach_button.setToolTip("Detach current VM"); self.detach_button.clicked.connect(self.detach_current_tab); button_layout.addWidget(self.detach_button)
        self.close_window_button = QToolButton(); self.close_window_button.setIcon(close_icon); self.close_window_button.setToolTip("Close current VM window"); self.close_window_button.clicked.connect(self.close_current_window); button_layout.addWidget(self.close_window_button)
        self.close_all_button = QToolButton(); self.close_all_button.setIcon(close_all_icon); self.close_all_button.setToolTip("Close ALL VM windows"); self.close_all_button.clicked.connect(self.close_all_vms); button_layout.addWidget(self.close_all_button)
        
        separator2 = QFrame(); separator2.setFrameShape(QFrame.VLine); separator2.setFrameShadow(QFrame.Sunken); button_layout.addWidget(separator2)

        self.rename_button = QToolButton(); self.rename_button.setIcon(rename_icon); self.rename_button.setToolTip("Rename current tab"); self.rename_button.clicked.connect(self.rename_current_tab); button_layout.addWidget(self.rename_button)
        self.info_button = QToolButton(); self.info_button.setIcon(info_icon); self.info_button.setToolTip("Show current VM Information"); self.info_button.clicked.connect(self.show_current_vm_info); button_layout.addWidget(self.info_button)
        self.snapshot_button = QToolButton(); self.snapshot_button.setIcon(snapshot_icon); self.snapshot_button.setToolTip("Take a snapshot of the current VM"); self.snapshot_button.clicked.connect(self.take_snapshot); button_layout.addWidget(self.snapshot_button)
        self.cad_button = QToolButton(); self.cad_button.setIcon(cad_icon); self.cad_button.setToolTip("Send Ctrl+Alt+Del to current VM"); self.cad_button.clicked.connect(self.send_ctrl_alt_del); button_layout.addWidget(self.cad_button)

        self.power_button = QToolButton(); self.power_button.setIcon(power_icon); self.power_button.setToolTip("Power options for current VM");
        power_menu = QMenu(self)
        power_menu.addAction(style.standardIcon(QStyle.SP_BrowserStop), "Suspend", self.suspend_vm)
        power_menu.addAction(style.standardIcon(QStyle.SP_DialogYesButton), "Restart", self.restart_vm)
        power_menu.addAction(style.standardIcon(QStyle.SP_DialogNoButton), "Shutdown", self.shutdown_vm)
        self.power_button.setMenu(power_menu); self.power_button.setPopupMode(QToolButton.InstantPopup); button_layout.addWidget(self.power_button)
        
        separator3 = QFrame(); separator3.setFrameShape(QFrame.VLine); separator3.setFrameShadow(QFrame.Sunken); button_layout.addWidget(separator3)
        
        self.vbox_main_button = QToolButton(); self.vbox_main_button.setIcon(vbox_icon); self.vbox_main_button.setToolTip("Open VirtualBox main application"); self.vbox_main_button.clicked.connect(self.open_virtualbox_main); button_layout.addWidget(self.vbox_main_button)
        
        button_layout.addStretch(1)
        
        self.settings_button = QToolButton(); self.settings_button.setIcon(settings_icon); self.settings_button.setToolTip("Settings"); self.settings_button.clicked.connect(self.show_settings_dialog); button_layout.addWidget(self.settings_button)
        self.about_button = QToolButton(); self.about_button.setIcon(about_icon); self.about_button.setToolTip("About"); self.about_button.clicked.connect(self.show_about_dialog); button_layout.addWidget(self.about_button)
        
        main_layout.addLayout(button_layout)

        self.tab_widget = QTabWidget(); custom_tab_bar = MiddleClickCloseTabBar(self.tab_widget); custom_tab_bar.middleCloseRequested.connect(self.close_tab_by_index)
        self.tab_widget.setTabBar(custom_tab_bar); self.tab_widget.setMovable(True); self.tab_widget.setTabsClosable(False); main_layout.addWidget(self.tab_widget)
        self.tabs = {}; self.auto_refresh_timer = QTimer(self); self.auto_refresh_timer.timeout.connect(self.refresh_tabs); self.auto_refresh_timer.start(self.settings.get("refresh_interval", 5) * 1000)
        self.refresh_tabs()
        self.tab_widget.tabBar().setContextMenuPolicy(Qt.CustomContextMenu); self.tab_widget.tabBar().customContextMenuRequested.connect(self.show_tab_context_menu)
        self.tab_widget.currentChanged.connect(self.update_tool_buttons)
        self.update_tool_buttons(self.tab_widget.currentIndex())

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

    def get_settings_path(self): return os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), 'settings.json')
    def load_settings(self):
        defaults = {"auto_attach": True, "refresh_interval": 5, "vbox_path": r"C:\Program Files\Oracle\VirtualBox\VirtualBox.exe", "theme": "Fusion", "dpi_scaling": "Auto"}
        if not os.path.exists(self.settings_file): return defaults
        try:
            with open(self.settings_file, 'r') as f: settings = json.load(f)
            settings.update({k: v for k, v in defaults.items() if k not in settings}); return settings
        except Exception: return defaults
    def save_settings(self):
        try:
            with open(self.settings_file, 'w') as f: json.dump(self.settings, f, indent=4)
        except Exception as e: QMessageBox.warning(self, "Error", f"Failed to save settings: {str(e)}")
    def apply_dpi_scaling(self):
        scaling = self.settings.get("dpi_scaling", "Auto")
        os.environ['QT_AUTO_SCREEN_SCALE_FACTOR'] = '1' if scaling == "Auto" else '0'
        if scaling != "Auto":
            try: os.environ['QT_SCALE_FACTOR'] = str(float(scaling.strip('%')) / 100.0)
            except ValueError: os.environ['QT_AUTO_SCREEN_SCALE_FACTOR'] = '1'
    def show_settings_dialog(self):
        dialog = SettingsDialog(self.settings, self)
        if dialog.exec():
            new = dialog.get_settings(); old = self.settings.copy(); self.settings = new; self.save_settings()
            if new["refresh_interval"] != old.get("refresh_interval"): self.auto_refresh_timer.start(new["refresh_interval"] * 1000)
            if new["theme"] != old.get("theme"): self.change_theme(new["theme"])
            if new["dpi_scaling"] != old.get("dpi_scaling"): QMessageBox.information(self, "Restart Required", "DPI scaling changes will take effect after restarting the application.")
    
    def change_theme(self, theme_name):
        app = QApplication.instance()
        app.setStyleSheet("")
        app.setPalette(QPalette())

        theme_key = self.theme_map.get(theme_name)
        if theme_key == "qdarkstyle" and QDARKSTYLE_AVAILABLE:
            app.setStyleSheet(qdarkstyle.load_stylesheet(qt_api='pyside6'))
        elif theme_key:
            QApplication.setStyle(QStyleFactory.create(theme_key))
        else:
            QApplication.setStyle(QStyleFactory.create("Fusion"))

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
        
        is_manual_attach_all = hasattr(self, 'attach_all_button') and self.sender() == self.attach_all_button
        was_manually_detached = hwnd in self.manually_detached_windows
        auto_attach_enabled = self.settings.get("auto_attach", True)
        should_attach = force_attach or is_manual_attach_all or (auto_attach_enabled and not was_manually_detached)
        
        original_title = window_info['title']
        display_title = original_title
        win_type = window_info.get('type', 'Unknown')
        if win_type != 'VirtualBox' and len(original_title) > 20:
            display_title = original_title[:20] + '...'
        window_info['display_title'] = display_title

        tab = VBoxTab(window_info, self.window_manager)
        self.tabs[hwnd] = tab
        tab_text = display_title if win_type == 'VirtualBox' else f"[{win_type}] {display_title}"
        index = self.tab_widget.addTab(tab, tab_text)
        self.tab_widget.setTabToolTip(index, f"Type: {win_type}\nOriginal: {original_title}\nHandle: {hwnd}")
        
        if should_attach:
            if was_manually_detached: self.manually_detached_windows.remove(hwnd)
            tab.attach_window()
        
        self.tab_widget.setCurrentIndex(index)

    def refresh_tabs(self):
        is_attach_all_request = hasattr(self, 'attach_all_button') and self.sender() == self.attach_all_button
        found_windows = self.window_manager.find_embeddable_windows()
        
        for hwnd, tab in list(self.tabs.items()):
            if not self.window_manager.is_window(hwnd):
                index = self.tab_widget.indexOf(tab)
                if index != -1: self.tab_widget.removeTab(index)
                del self.tabs[hwnd]
                if hwnd in self.manually_detached_windows: self.manually_detached_windows.remove(hwnd)
        
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
        
        # Iterate backwards to safely remove tabs
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
            if hwnd in self.manually_detached_windows: self.manually_detached_windows.remove(hwnd)
    def close_current_window(self): self.close_tab_by_index(self.tab_widget.currentIndex())
    def closeEvent(self, event):
        for tab in self.tabs.values(): tab.detach_window()
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
        current_name = tab.window_info['title']
        new_name, ok = QInputDialog.getText(self, "Rename Tab", "Enter new name:", text=current_name)
        if ok and new_name:
            tab.window_info['title'] = new_name
            display_name = new_name
            if tab.window_info.get('type') != 'VirtualBox' and len(new_name) > 20:
                display_name = new_name[:20] + '...'
            tab.window_info['display_title'] = display_name
            tab_text = display_name if tab.window_info.get('type') == 'VirtualBox' else f"[{tab.window_info.get('type')}] {display_name}"
            self.tab_widget.setTabText(index, tab_text)

    def _get_current_vm_name(self):
        index = self.tab_widget.currentIndex()
        if index < 0: return None, "No tab selected."
        tab = self.tab_widget.widget(index)
        if not isinstance(tab, VBoxTab) or tab.window_info.get('type') != 'VirtualBox':
            return None, "Selected tab is not a VirtualBox VM."
        return self._get_registered_vm_name(tab.window_info['title'])

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

    def _run_vboxmanage_command(self, command_args):
        vm_name, error = self._get_current_vm_name()
        if error:
            QMessageBox.warning(self, "Command Error", error)
            return
        
        vbox_manage_path = os.path.join(os.path.dirname(self.settings.get("vbox_path", "")), "VBoxManage.exe")
        full_command = [vbox_manage_path] + ["controlvm", vm_name] + command_args
        
        try:
            subprocess.run(full_command, check=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW if IS_WINDOWS else 0)
        except subprocess.CalledProcessError as e:
            error_output = e.stderr.decode('utf-8', errors='replace') if e.stderr else str(e)
            QMessageBox.critical(self, "VBoxManage Error", f"Failed to execute command:\n{' '.join(full_command)}\n\n{error_output}")
        except Exception as e:
            QMessageBox.critical(self, "Execution Error", f"An unexpected error occurred:\n\n{str(e)}")

    def take_snapshot(self):
        vm_name, error = self._get_current_vm_name()
        if error:
            QMessageBox.warning(self, "Snapshot Error", error)
            return

        default_name = f"Snapshot_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
        snap_name, ok = QInputDialog.getText(self, "Take Snapshot", f"Enter snapshot name for '{vm_name}':", text=default_name)
        
        if ok and snap_name:
            vbox_manage_path = os.path.join(os.path.dirname(self.settings.get("vbox_path", "")), "VBoxManage.exe")
            full_command = [vbox_manage_path, "snapshot", vm_name, "take", snap_name]
            try:
                subprocess.run(full_command, check=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW if IS_WINDOWS else 0)
                QMessageBox.information(self, "Success", f"Snapshot '{snap_name}' created successfully.")
            except subprocess.CalledProcessError as e:
                error_output = e.stderr.decode('utf-8', errors='replace') if e.stderr else str(e)
                QMessageBox.critical(self, "VBoxManage Error", f"Failed to take snapshot:\n\n{error_output}")
            except Exception as e:
                 QMessageBox.critical(self, "Execution Error", f"An unexpected error occurred:\n\n{str(e)}")

    def send_ctrl_alt_del(self):
        self._run_vboxmanage_command(["keyboardputscancode", "1d", "38", "e0", "53", "e0", "d3", "e0", "b8", "e0", "9d"])
    
    def suspend_vm(self): self._run_vboxmanage_command(["savestate"])
    def restart_vm(self): self._run_vboxmanage_command(["reset"])
    def shutdown_vm(self): self._run_vboxmanage_command(["acpipowerbutton"])

    def show_current_vm_info(self):
        registered_name, error = self._get_current_vm_name()
        if registered_name:
            VmInfoDialog(registered_name, self.settings, self).exec()
        #else:
        #    QMessageBox.warning(self, "VM Info Error", f"Could not get information for this VM.\n\nReason: {error}")

    def show_tab_context_menu(self, pos):
        tabBar = self.tab_widget.tabBar(); index = tabBar.tabAt(pos)
        if index < 0: return
        self.tab_widget.setCurrentIndex(index); tab = self.tab_widget.widget(index)
        style = QApplication.style(); menu = QMenu(self)
        rename_action = menu.addAction(style.standardIcon(QStyle.SP_FileDialogNewFolder), "Rename Tab")
        detach_action = menu.addAction(style.standardIcon(QStyle.SP_DialogCancelButton), "Detach Window")
        close_action = menu.addAction(style.standardIcon(QStyle.SP_DialogResetButton), "Close Window")
        
        # Initialize action variables to None
        info_action = None
        snapshot_action = None

        if tab.window_info.get('type') == "VirtualBox":
            menu.addSeparator()
            info_action = menu.addAction(style.standardIcon(QStyle.SP_MessageBoxInformation), "Show VM Info")
            snapshot_action = menu.addAction(style.standardIcon(QStyle.SP_DialogSaveButton), "Take Snapshot")
        
        action = menu.exec(tabBar.mapToGlobal(pos))
        
        if action == rename_action: self._rename_tab_at_index(index)
        elif action == detach_action: self.detach_tab(index)
        elif action == close_action: self.close_tab_by_index(index)
        elif action == info_action: self.show_current_vm_info()
        elif action == snapshot_action: self.take_snapshot()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    if IS_WINDOWS: window_manager = WindowsWindowManager()
    elif IS_LINUX: window_manager = LinuxWindowManager()
    else: QMessageBox.critical(None, "Unsupported OS", f"This application does not support {platform.system()}."); sys.exit(1)
    if not QDARKSTYLE_AVAILABLE: print("Warning: qdarkstyle not found. The 'QDark' theme may be unavailable.")
    
    window = VirtualBoxTabs(window_manager)
    window.show()
    
    QTimer.singleShot(50, lambda: window.change_theme(window.settings.get("theme", "Fusion")))
    
    sys.exit(app.exec())