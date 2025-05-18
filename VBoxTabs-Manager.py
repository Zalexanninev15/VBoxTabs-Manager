# VBoxTabs Manager
# Copyright (c) 2025 Zalexanninev15
# https://github.com/Zalexanninev15/VBoxTabs-Manager
#
# Licensed under the MIT License. See LICENSE file in the project root for full license information.
"""
VBoxTabs Manager - Combining VirtualBox windows into a single tabbed window.
"""

import sys
import os
import subprocess
import json
import time # For debugging timestamps
import win32gui
import win32process
import win32con
import win32api
from PySide6.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget,
                               QVBoxLayout, QPushButton, QMessageBox, QLabel,
                               QDialog, QHBoxLayout, QInputDialog, QCheckBox,
                               QStyleFactory, QComboBox, QToolButton, QMenu,
                               QStyle, QTabBar, QSpinBox, QLineEdit, QGridLayout,
                               QGroupBox, QFileDialog)
from PySide6.QtCore import Qt, QTimer, Signal, QObject, QSize, QPoint, QSettings, QRect
from PySide6.QtGui import (QFont, QAction, QMouseEvent, QIcon, QPainter, QColor,
                           QScreen, QPaintEvent, QKeyEvent, QShowEvent, QPen)


# Import for Windows registry access (for theme detection)
try:
    import winreg
    WINREG_AVAILABLE = True
except ImportError:
    WINREG_AVAILABLE = False

# QDarkStyle (optional)
try:
    import qdarkstyle
    QDARKSTYLE_AVAILABLE = True
except ImportError:
    QDARKSTYLE_AVAILABLE = False

# qt-themes (optional)
try:
    import qt_themes
    QT_THEMES_AVAILABLE = True
except ImportError:
    QT_THEMES_AVAILABLE = False

# Define Win32 API constants
WS_CHILD = 0x40000000
GWL_STYLE = -16
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOZORDER = 0x0004
SWP_FRAMECHANGED = 0x0020
SWP_NOACTIVATE = 0x0010
SWP_SHOWWINDOW = 0x0040
SWP_NOSHOWWINDOW = 0x0080
WS_CAPTION = 0x00C00000
WS_THICKFRAME = 0x00040000
WS_MINIMIZEBOX = 0x00020000
WS_MAXIMIZEBOX = 0x00010000
WS_SYSMENU = 0x00080000

# For SetWindowPos Z-order
HWND_TOPMOST = -1
HWND_NOTOPMOST = -2
# For ShowWindow
SW_MINIMIZE = 6
SW_RESTORE = 9


VB_MANAGER = False

# Helper function to detect Windows light/dark theme
def is_windows_light_theme():
    """Checks if Windows is set to use a light theme for apps."""
    if not WINREG_AVAILABLE:
        return False
    try:
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        return value == 1
    except (FileNotFoundError, OSError):
        return False

class WindowFinder:
    """Class for finding VirtualBox windows"""

    def __init__(self):
        self.virtualbox_windows = []

    def enum_windows_callback(self, hwnd, _):
        """Callback for EnumWindows"""
        if win32gui.IsWindowVisible(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            if window_title and ("[Running]" in window_title or "[Работает]" in window_title) and " Oracle VirtualBox" in window_title:
                rect = win32gui.GetWindowRect(hwnd)
                width = rect[2] - rect[0]
                height = rect[3] - rect[1]
                vm_name = window_title
                if "[Running] - Oracle VirtualBox" in vm_name:
                    vm_name = vm_name.split(" [Running] - Oracle VirtualBox")[0]
                elif "[Работает] - Oracle VirtualBox" in vm_name:
                    vm_name = vm_name.split(" [Работает] - Oracle VirtualBox")[0]
                self.virtualbox_windows.append({
                    'hwnd': hwnd, 'title': vm_name, 'original_title': window_title,
                    'width': width, 'height': height
                })
            elif window_title and ("Oracle VirtualBox " in window_title): # Potential VB Manager
                if "VirtualBox Manager" in window_title or window_title.startswith("Oracle VM VirtualBox"):
                    rect = win32gui.GetWindowRect(hwnd)
                    width = rect[2] - rect[0]
                    height = rect[3] - rect[1]
                    self.virtualbox_windows.append({
                        'hwnd': hwnd, 'title': "VB Manager", 'original_title': window_title,
                        'width': width, 'height': height
                    })
        return True

    def find_virtualbox_windows(self):
        """Finds all visible VirtualBox windows"""
        self.virtualbox_windows = []
        win32gui.EnumWindows(self.enum_windows_callback, None)
        return self.virtualbox_windows


class WindowManager:
    """Class for managing windows using Win32 API"""

    @staticmethod
    def set_window_parent(hwnd, parent_hwnd):
        style = win32gui.GetWindowLong(hwnd, GWL_STYLE)
        old_styles = style & (WS_CAPTION | WS_THICKFRAME | WS_MINIMIZEBOX |
                              WS_MAXIMIZEBOX | WS_SYSMENU)
        new_style = (style & ~(WS_CAPTION | WS_THICKFRAME | WS_MINIMIZEBOX |
                               WS_MAXIMIZEBOX | WS_SYSMENU)) | WS_CHILD
        win32gui.SetWindowLong(hwnd, GWL_STYLE, new_style)
        win32gui.SetParent(hwnd, parent_hwnd)
        win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                              SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)
        return old_styles

    @staticmethod
    def restore_window_style(hwnd, old_styles):
        current_style = win32gui.GetWindowLong(hwnd, GWL_STYLE)
        new_style = (current_style & ~WS_CHILD) | old_styles
        win32gui.SetWindowLong(hwnd, GWL_STYLE, new_style)
        win32gui.SetParent(hwnd, 0)
        win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                              SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)


class SettingsDialog(QDialog):
    """Settings dialog window"""
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("Settings")
        self.setMinimumWidth(500)
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint |
                            Qt.WindowTitleHint | Qt.WindowCloseButtonHint)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        general_group = QGroupBox("General Settings")
        general_layout = QGridLayout(general_group)
        general_layout.setColumnStretch(1, 1)
        self.auto_attach_checkbox = QCheckBox("Automatically attach VirtualBox windows")
        self.auto_attach_checkbox.setChecked(self.settings.get("auto_attach", True))
        general_layout.addWidget(self.auto_attach_checkbox, 0, 0, 1, 2)
        general_layout.addWidget(QLabel("Window detection interval (seconds):"), 1, 0)
        self.refresh_interval_spinbox = QSpinBox()
        self.refresh_interval_spinbox.setRange(1, 60)
        self.refresh_interval_spinbox.setValue(self.settings.get("refresh_interval", 5))
        general_layout.addWidget(self.refresh_interval_spinbox, 1, 1)
        general_layout.addWidget(QLabel("VirtualBox executable path:"), 2, 0)
        vbox_path_layout = QHBoxLayout()
        self.vbox_path_edit = QLineEdit(self.settings.get("vbox_path", r"C:\Program Files\Oracle\VirtualBox\VirtualBox.exe"))
        vbox_path_layout.addWidget(self.vbox_path_edit)
        self.browse_button = QPushButton("Browse...")
        self.browse_button.clicked.connect(self.browse_vbox_path)
        vbox_path_layout.addWidget(self.browse_button)
        general_layout.addLayout(vbox_path_layout, 2, 1)
        main_layout.addWidget(general_group)

        display_group = QGroupBox("Display Settings")
        display_layout = QGridLayout(display_group)
        display_layout.setColumnStretch(1, 1)
        display_layout.addWidget(QLabel("Theme:"), 0, 0)
        self.theme_combo = QComboBox()
        theme_names = []
        if not is_windows_light_theme():
            theme_names.append("Dark")
        theme_names.extend(["Light", "Classic", "Fusion"])
        if QDARKSTYLE_AVAILABLE: theme_names.append("QDark")
        if QT_THEMES_AVAILABLE:
            theme_names.extend(["Atom One", "Blender", "Catppuccin Frappe", "Catppuccin Latte",
                                "Catppuccin Macchiato", "Catppuccin Mocha", "Dracula",
                                "GitHub Dark", "GitHub Light", "Modern Dark", "Modern Light",
                                "Monokai", "Nord", "One Dark Two"])
        self.theme_combo.addItems(theme_names)
        current_theme = self.settings.get("theme", "Fusion")
        if current_theme in theme_names: self.theme_combo.setCurrentText(current_theme)
        else: self.theme_combo.setCurrentText("Fusion" if "Fusion" in theme_names else (theme_names[0] if theme_names else ""))
        display_layout.addWidget(self.theme_combo, 0, 1)
        display_layout.addWidget(QLabel("DPI Scaling:"), 1, 0)
        self.dpi_scaling_combo = QComboBox()
        self.dpi_scaling_combo.addItems(["Auto", "100%", "125%", "150%", "175%", "200%"])
        self.dpi_scaling_combo.setCurrentText(self.settings.get("dpi_scaling", "Auto"))
        display_layout.addWidget(self.dpi_scaling_combo, 1, 1)
        main_layout.addWidget(display_group)

        button_layout = QHBoxLayout()
        button_layout.addStretch(1)
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_button)
        self.save_button = QPushButton("Save")
        self.save_button.clicked.connect(self.accept)
        button_layout.addWidget(self.save_button)
        main_layout.addLayout(button_layout)

    def browse_vbox_path(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select VirtualBox Executable",
            os.path.dirname(self.vbox_path_edit.text()),
            "VirtualBox Executable (VirtualBox.exe);;All Executable Files (*.exe)")
        if file_path: self.vbox_path_edit.setText(file_path)

    def get_settings(self):
        return {"auto_attach": self.auto_attach_checkbox.isChecked(),
                "refresh_interval": self.refresh_interval_spinbox.value(),
                "vbox_path": self.vbox_path_edit.text(),
                "theme": self.theme_combo.currentText(),
                "dpi_scaling": self.dpi_scaling_combo.currentText()}

class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About VBoxTabs Manager")
        self.setFixedSize(400, 250)
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(0,0,0,0); content_layout.setSpacing(8)
        title_label = QLabel("VBoxTabs Manager"); title_font=QFont(); title_font.setPointSize(16); title_font.setBold(True); title_label.setFont(title_font); title_label.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(title_label)
        desc_label = QLabel("A tool for combining VirtualBox windows into a single tabbed window."); desc_label.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(desc_label)
        content_layout.addWidget(QLabel("Version 1.4pa1", alignment=Qt.AlignCenter))
        content_layout.addWidget(QLabel("Author: Zalexanninev15", alignment=Qt.AlignCenter))
        github_label = QLabel('<a href="https://github.com/Zalexanninev15/VBoxTabs-Manager">GitHub Repository</a>'); github_label.setOpenExternalLinks(True); github_label.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(github_label)
        content_layout.addWidget(QLabel("License: MIT", alignment=Qt.AlignCenter))
        content_layout.addSpacing(10)
        info_label = QLabel("This application allows you to manage multiple VirtualBox machines "
            "in a single window with tabs. It preserves all VirtualBox functionality "
            "while providing a more convenient interface.\n\n"
            "Built with PySide6 (Qt6) and Win32 API."
        ); info_label.setWordWrap(True); info_label.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(info_label)
        main_layout.addWidget(content_widget, 1)
        button_container = QWidget(); button_layout = QHBoxLayout(button_container); button_layout.setContentsMargins(0,10,0,0)
        ok_button = QPushButton("OK"); ok_button.setFixedWidth(80); ok_button.clicked.connect(self.accept)
        button_layout.addStretch(); button_layout.addWidget(ok_button); button_layout.addStretch()
        main_layout.addWidget(button_container)

class VBoxTab(QWidget):
    """Tab widget for VirtualBox window"""
    def __init__(self, window_info, parent=None):
        super().__init__(parent)
        self.theme_map = {
            "Dark": "windows11", "Light": "windowsvista", "Classic": "Windows", "Fusion": "Fusion", "QDark": "qdarkstyle",
            "Atom One": "atom_one", "Blender": "blender", "Catppuccin Frappe": "catppuccin_frappe", "Catppuccin Latte": "catppuccin_latte",
            "Catppuccin Macchiato": "catppuccin_macchiato", "Catppuccin Mocha": "catppuccin_mocha", "Dracula": "dracula",
            "GitHub Dark": "github_dark", "GitHub Light": "github_light", "Modern Dark": "modern_dark", "Modern Light": "modern_light",
            "Monokai": "monokai", "Nord": "nord", "One Dark Two": "one_dark_two"
        }
        self.window_info = window_info
        self.hwnd = window_info['hwnd']
        self.title = window_info['title']
        self.original_title = window_info['original_title']
        self.orig_styles = None
        self.attached = False
        self.detached_manually = False
        self.is_custom = window_info.get('is_custom', False) 

        layout = QVBoxLayout(self); layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(layout)
        self.container = QWidget(self)
        layout.addWidget(self.container)

    def attach_window(self):
        if not self.attached and win32gui.IsWindow(self.hwnd): 
            self.orig_styles = WindowManager.set_window_parent(self.hwnd, int(self.container.winId()))
            win32gui.MoveWindow(self.hwnd, 0, 0, self.container.width(), self.container.height(), True)
            self.attached = True
            self.detached_manually = False 

    def detach_window(self):
        if self.attached and self.orig_styles is not None and win32gui.IsWindow(self.hwnd):
            WindowManager.restore_window_style(self.hwnd, self.orig_styles)
            win32gui.ShowWindow(self.hwnd, win32con.SW_SHOW) 
            self.attached = False

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.attached and win32gui.IsWindow(self.hwnd):
            win32gui.MoveWindow(self.hwnd, 0, 0, self.container.width(), self.container.height(), True)

class RefreshSignal(QObject):
    refreshRequested = Signal()

class MiddleClickCloseTabBar(QTabBar):
    middleCloseRequested = Signal(int)
    def __init__(self, parent=None): super().__init__(parent)
    def mouseReleaseEvent(self, event: QMouseEvent):
        if event.button() == Qt.MouseButton.MiddleButton:
            tab_index = self.tabAt(event.position().toPoint())
            if tab_index != -1:
                self.middleCloseRequested.emit(tab_index)
                event.accept(); return
        super().mouseReleaseEvent(event)

# WindowSelectorOverlay class definition
class WindowSelectorOverlay(QWidget):
    windowSelected = Signal(int, str, str, int, int)
    closed = Signal() 

    def __init__(self, main_app_hwnd, parent=None):
        super().__init__(parent)
        # print(f"{time.time():.3f} Overlay: __init__ (Main HWND: {main_app_hwnd})")
        self.main_app_hwnd = main_app_hwnd
        self.hovered_hwnd = None
        self.hovered_rect = None

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool | Qt.X11BypassWindowManagerHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_DeleteOnClose) # Ensure it's cleaned up
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.StrongFocus) # Try to ensure it can get focus

        geometry = QRect()
        for screen in QApplication.screens():
            geometry = geometry.united(screen.geometry())
        self.setGeometry(geometry)
        self.setCursor(Qt.CrossCursor)

        self.desktop_hwnd = win32gui.GetDesktopWindow()
        self.progman_hwnd = win32gui.FindWindow("Progman", "Program Manager")
        
        self.workerw_hwnds = []
        workerw_hwnd_candidate = win32gui.FindWindowEx(0, 0, "WorkerW", None)
        while workerw_hwnd_candidate:
            self.workerw_hwnds.append(workerw_hwnd_candidate)
            workerw_hwnd_candidate = win32gui.FindWindowEx(0, workerw_hwnd_candidate, "WorkerW", None)
        if self.progman_hwnd:
            shelldll_defview = win32gui.FindWindowEx(self.progman_hwnd, 0, "SHELLDLL_DefView", None)
            if shelldll_defview:
                self.workerw_hwnds.append(shelldll_defview)
        
    def mouseMoveEvent(self, event: QMouseEvent):
        global_pos = event.globalPosition().toPoint()
        screen_pos_tuple = (global_pos.x(), global_pos.y())
        
        hwnd_under_cursor = win32gui.WindowFromPoint(screen_pos_tuple)
        hwnd = win32gui.GetAncestor(hwnd_under_cursor, win32con.GA_ROOTOWNER)

        new_hovered_hwnd = None
        new_hovered_rect = None

        if hwnd and hwnd != self.winId() and \
           hwnd != self.main_app_hwnd and \
           hwnd != self.desktop_hwnd and \
           hwnd != self.progman_hwnd and \
           hwnd not in self.workerw_hwnds and \
           win32gui.IsWindowVisible(hwnd) and \
           not win32gui.IsIconic(hwnd):
            
            try:
                current_rect_test = win32gui.GetWindowRect(hwnd)
                if ((current_rect_test[2] - current_rect_test[0]) > 10 and \
                    (current_rect_test[3] - current_rect_test[1]) > 10):
                    new_hovered_hwnd = hwnd
                    new_hovered_rect = current_rect_test
            except win32gui.error:
                pass 

        if self.hovered_hwnd != new_hovered_hwnd or self.hovered_rect != new_hovered_rect:
            self.hovered_hwnd = new_hovered_hwnd
            self.hovered_rect = new_hovered_rect
            self.update()

    def paintEvent(self, event: QPaintEvent):
        painter = QPainter(self)
        painter.setCompositionMode(QPainter.CompositionMode_Clear)
        painter.fillRect(self.rect(), Qt.transparent)
        painter.setCompositionMode(QPainter.CompositionMode_SourceOver)

        if self.hovered_hwnd and self.hovered_rect:
            try:
                if not win32gui.IsWindow(self.hovered_hwnd):
                    self.hovered_hwnd = None; self.hovered_rect = None; self.update(); return

                top_left_global = QPoint(self.hovered_rect[0], self.hovered_rect[1])
                width = self.hovered_rect[2] - self.hovered_rect[0]
                height = self.hovered_rect[3] - self.hovered_rect[1]
                top_left_local = self.mapFromGlobal(top_left_global)
                draw_q_rect = QRect(top_left_local.x(), top_left_local.y(), width, height)

                pen = QPen(QColor("#8e8cd8"), 3)
                painter.setPen(pen)
                painter.drawRect(draw_q_rect)
            except Exception as e:
                self.hovered_hwnd = None; self.hovered_rect = None

    def mousePressEvent(self, event: QMouseEvent):
        # print(f"{time.time():.3f} Overlay: mousePressEvent - Button: {event.button()}, Hovered: {self.hovered_hwnd}")
        if event.button() == Qt.LeftButton and self.hovered_hwnd:
            try:
                if not win32gui.IsWindow(self.hovered_hwnd): self.custom_close(); return
                title = win32gui.GetWindowText(self.hovered_hwnd)
                if not title:
                    class_name = win32gui.GetClassName(self.hovered_hwnd)
                    title = f"Window ({class_name})" if class_name else "Untitled Window"
                
                width = self.hovered_rect[2] - self.hovered_rect[0]
                height = self.hovered_rect[3] - self.hovered_rect[1]
                self.windowSelected.emit(self.hovered_hwnd, title, title, width, height)
            except Exception as e: print(f"Error emitting windowSelected: {e}")
            finally: self.custom_close()
        elif event.button() == Qt.RightButton: self.custom_close()

    def keyPressEvent(self, event: QKeyEvent):
        # print(f"{time.time():.3f} Overlay: keyPressEvent - Key: {event.key()}")
        if event.key() == Qt.Key_Escape:
            self.custom_close()

    def showEvent(self, event: QShowEvent):
        # print(f"{time.time():.3f} Overlay: showEvent - Geometry: {self.geometry()}")
        super().showEvent(event)
        # These are crucial for making it truly active and on top
        self.raise_() 
        self.activateWindow()
        self.setFocus() # Explicitly set focus
        if self.winId():
             # print(f"{time.time():.3f} Overlay: Setting HWND_TOPMOST for {self.winId()}")
             win32gui.SetWindowPos(int(self.winId()), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOACTIVATE) # SWP_NOACTIVATE to prevent stealing focus from itself immediately? Test this.
             win32gui.SetForegroundWindow(int(self.winId())) # Try to bring it to foreground
        # print(f"{time.time():.3f} Overlay: showEvent finished.")


    def custom_close(self):
        # print(f"{time.time():.3f} Overlay: custom_close called. Visible: {self.isVisible()}")
        if not self.isVisible(): # Prevent double close or actions if already closing
            return
        self.closed.emit() 
        if self.winId():
             win32gui.SetWindowPos(int(self.winId()), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOSHOWWINDOW ) 
        self.hide() # Hide before close to make it disappear faster
        self.deleteLater() # Schedule for deletion instead of immediate self.close()
                           # This is safer if events are still pending.


class VirtualBoxTabs(QMainWindow):
    """Main application window"""
    def __init__(self):
        super().__init__()
        self.manually_detached_windows = set()
        self.theme_map = {
            "Dark": "windows11", "Light": "windowsvista", "Classic": "Windows", "Fusion": "Fusion", "QDark": "qdarkstyle",
            "Atom One": "atom_one", "Blender": "blender", "Catppuccin Frappe": "catppuccin_frappe", "Catppuccin Latte": "catppuccin_latte",
            "Catppuccin Macchiato": "catppuccin_macchiato", "Catppuccin Mocha": "catppuccin_mocha", "Dracula": "dracula",
            "GitHub Dark": "github_dark", "GitHub Light": "github_light", "Modern Dark": "modern_dark", "Modern Light": "modern_light",
            "Monokai": "monokai", "Nord": "nord", "One Dark Two": "one_dark_two"
        }
        self.settings_file = self.get_settings_path()
        self.settings = self.load_settings()
        self.apply_dpi_scaling()
        self.setWindowTitle("VBoxTabs Manager"); self.resize(1280, 800); self.setMinimumSize(420, 251)
        style = QApplication.style()
        self.setWindowIcon(style.standardIcon(QStyle.SP_ComputerIcon))
        central_widget = QWidget(); self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        button_layout = QHBoxLayout()

        refresh_icon = style.standardIcon(QStyle.SP_BrowserReload)
        detach_icon = style.standardIcon(QStyle.SP_DialogCancelButton)
        attach_icon = style.standardIcon(QStyle.SP_ArrowUp)
        about_icon = style.standardIcon(QStyle.SP_MessageBoxInformation)
        rename_icon = style.standardIcon(QStyle.SP_FileDialogNewFolder)
        vbox_icon = style.standardIcon(QStyle.SP_ComputerIcon)
        close_icon = style.standardIcon(QStyle.SP_DialogResetButton)
        settings_icon = style.standardIcon(QStyle.SP_FileDialogDetailedView)
        close_all_icon = style.standardIcon(QStyle.SP_DialogCloseButton)

        self.refresh_button = QToolButton(); self.refresh_button.setIcon(refresh_icon); self.refresh_button.setToolTip("Refresh VM list"); self.refresh_button.clicked.connect(self.refresh_tabs)
        self.attach_button = QToolButton(); self.attach_button.setIcon(attach_icon); self.attach_button.setToolTip("Attach all available VMs"); self.attach_button.clicked.connect(self.refresh_tabs)
        self.detach_button = QToolButton(); self.detach_button.setIcon(detach_icon); self.detach_button.setToolTip("Detach current VM"); self.detach_button.clicked.connect(self.detach_current_tab)
        self.close_window_button = QToolButton(); self.close_window_button.setIcon(close_icon); self.close_window_button.setToolTip("Close current VM window"); self.close_window_button.clicked.connect(self.close_current_window)
        self.close_all_button = QToolButton(); self.close_all_button.setIcon(close_all_icon); self.close_all_button.setToolTip("Close all VM windows"); self.close_all_button.clicked.connect(self.close_all_windows)
        self.rename_button = QToolButton(); self.rename_button.setIcon(rename_icon); self.rename_button.setToolTip("Rename current tab"); self.rename_button.clicked.connect(self.rename_current_tab)
        
        self.add_any_window_button = QToolButton()
        self.add_any_window_button.setText("+") 
        font = self.add_any_window_button.font(); font.setPointSize(font.pointSize() + 4); font.setBold(True)
        self.add_any_window_button.setFont(font)
        self.add_any_window_button.setToolTip("Add any window as a tab")
        self.add_any_window_button.clicked.connect(self.select_window_to_add_handler)
        self.overlay = None 

        self.vbox_main_button = QToolButton(); self.vbox_main_button.setIcon(vbox_icon); self.vbox_main_button.setToolTip("Open VirtualBox main application"); self.vbox_main_button.clicked.connect(self.open_virtualbox_main)
        self.settings_button = QToolButton(); self.settings_button.setIcon(settings_icon); self.settings_button.setToolTip("Settings"); self.settings_button.clicked.connect(self.show_settings_dialog)
        self.about_button = QToolButton(); self.about_button.setIcon(about_icon); self.about_button.setToolTip("About"); self.about_button.clicked.connect(self.show_about_dialog)

        button_layout.addWidget(self.refresh_button)
        button_layout.addWidget(self.attach_button)
        button_layout.addWidget(self.detach_button)
        button_layout.addWidget(self.close_window_button)
        button_layout.addWidget(self.close_all_button)
        button_layout.addWidget(self.rename_button)
        button_layout.addWidget(self.add_any_window_button) 
        button_layout.addWidget(self.vbox_main_button)
        button_layout.addStretch(1)
        button_layout.addWidget(self.settings_button)
        button_layout.addWidget(self.about_button)
        main_layout.addLayout(button_layout)

        self.tab_widget = QTabWidget()
        custom_tab_bar = MiddleClickCloseTabBar(); custom_tab_bar.middleCloseRequested.connect(self.close_tab_by_middle_click)
        self.tab_widget.setTabBar(custom_tab_bar)
        self.tab_widget.setTabsClosable(False); self.tab_widget.setMovable(True)
        main_layout.addWidget(self.tab_widget)

        self.window_finder = WindowFinder(); self.tabs = {}
        self.refresh_signal = RefreshSignal(); self.refresh_signal.refreshRequested.connect(self.refresh_tabs)
        self.auto_refresh_timer = QTimer()
        self.auto_refresh_timer.timeout.connect(lambda: self.refresh_signal.refreshRequested.emit())
        self.auto_refresh_timer.start(self.settings.get("refresh_interval", 5) * 1000)
        self.change_theme(self.settings.get("theme", "Fusion"))
        self.refresh_tabs()
        self.setAcceptDrops(True)
        self.tab_widget.tabBar().setContextMenuPolicy(Qt.CustomContextMenu)
        self.tab_widget.tabBar().customContextMenuRequested.connect(self.show_tab_context_menu)
        self.is_selecting_window = False 
        self.main_window_was_minimized = False # Track if we minimized it

    def get_settings_path(self):
        base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, 'settings.json')

    def load_settings(self):
        default_settings = {"auto_attach": True, "refresh_interval": 5, "vbox_path": r"C:\Program Files\Oracle\VirtualBox\VirtualBox.exe", "theme": "Fusion", "dpi_scaling": "Auto"}
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f: settings = json.load(f)
                for key, value in default_settings.items():
                    if key not in settings: settings[key] = value
                return settings
        except Exception as e: print(f"Error loading settings: {e}")
        return default_settings

    def save_settings(self):
        try:
            with open(self.settings_file, 'w') as f: json.dump(self.settings, f, indent=4)
        except Exception as e: print(f"Error saving settings: {e}"); QMessageBox.warning(self, "Error", f"Failed to save settings: {str(e)}")

    def apply_dpi_scaling(self):
        scaling = self.settings.get("dpi_scaling", "Auto")
        if scaling == "Auto":
            os.environ.pop('QT_SCALE_FACTOR', None) 
            os.environ['QT_AUTO_SCREEN_SCALE_FACTOR'] = '1'
        else:
            try: 
                scale_val = float(scaling.strip('%')) / 100.0
                os.environ['QT_SCALE_FACTOR'] = str(scale_val)
                os.environ.pop('QT_AUTO_SCREEN_SCALE_FACTOR', None) 
            except ValueError: 
                os.environ.pop('QT_SCALE_FACTOR', None)
                os.environ['QT_AUTO_SCREEN_SCALE_FACTOR'] = '1'
    
    def select_window_to_add_handler(self):
        # print(f"{time.time():.3f} Main: select_window_to_add_handler. is_selecting_window: {self.is_selecting_window}")
        if self.is_selecting_window:
            if self.overlay and self.overlay.isVisible():
                self.overlay.activateWindow()
                self.overlay.raise_()
            return

        if not self.winId(): 
            QMessageBox.warning(self, "Error", "Main window handle not yet available. Please try again.")
            return
        
        self.is_selecting_window = True
        self.main_window_was_minimized = False
        if self.winId():
            # print(f"{time.time():.3f} Main: Minimizing main window {self.winId()}")
            if not self.isMinimized():
                win32gui.ShowWindow(int(self.winId()), SW_MINIMIZE)
                self.main_window_was_minimized = True # We minimized it
            else:
                pass # It was already minimized, so we don't need to restore it differently

        # Use a short timer to allow the minimize animation/action to complete
        # before showing the overlay.
        QTimer.singleShot(250, self._show_overlay) # Increased delay for minimize

    def _show_overlay(self):
        # print(f"{time.time():.3f} Main: _show_overlay called")
        if not self.is_selecting_window: # Check if selection was cancelled before timer fired
            # print(f"{time.time():.3f} Main: Selection cancelled before overlay shown.")
            self._restore_main_window_state()
            return

        main_hwnd = int(self.winId()) if self.winId() else 0 
        self.overlay = WindowSelectorOverlay(main_app_hwnd=main_hwnd) 
        self.overlay.windowSelected.connect(self.add_selected_custom_window_handler)
        self.overlay.closed.connect(self._handle_overlay_closed) 
        self.overlay.show()
        # print(f"{time.time():.3f} Main: Overlay show() called for overlay {self.overlay}")


    def _handle_overlay_closed(self):
        # print(f"{time.time():.3f} Main: _handle_overlay_closed. is_selecting_window: {self.is_selecting_window}")
        if not self.is_selecting_window and not self.overlay : # Already handled or cancelled
             # print(f"{time.time():.3f} Main: Overlay already handled or selection cancelled.")
             return

        self.is_selecting_window = False
        self._restore_main_window_state()
        
        if self.overlay: # Overlay uses WA_DeleteOnClose now via custom_close -> deleteLater
            self.overlay = None 
            # print(f"{time.time():.3f} Main: Overlay reference cleared.")

    def _restore_main_window_state(self):
        # print(f"{time.time():.3f} Main: Restoring main window state.")
        if self.winId():
            if self.main_window_was_minimized: # Only restore if we minimized it
                # print(f"{time.time():.3f} Main: Restoring minimized main window {self.winId()}")
                win32gui.ShowWindow(int(self.winId()), SW_RESTORE)
            
            # Try to bring it to foreground regardless of minimized state
            # print(f"{time.time():.3f} Main: Setting main window foreground {self.winId()}")
            try:
                win32gui.SetForegroundWindow(int(self.winId()))
            except Exception as e:
                pass # print(f"Error setting foreground window: {e}") # Can fail if other window is more aggressive
            self.activateWindow() # Qt's way
            self.raise_()

    def add_selected_custom_window_handler(self, hwnd, title, original_title, width, height):
        # print(f"{time.time():.3f} Main: add_selected_custom_window_handler - HWND: {hwnd}, Title: {title}")
        # _handle_overlay_closed will be called due to overlay's custom_close
        
        window_info = {
            'hwnd': hwnd, 'title': title, 'original_title': original_title,
            'width': width, 'height': height, 'is_custom': True
        }

        if hwnd in self.tabs: 
            for i in range(self.tab_widget.count()):
                tab_at_i = self.tab_widget.widget(i)
                if isinstance(tab_at_i, VBoxTab) and tab_at_i.hwnd == hwnd:
                    self.tab_widget.setCurrentIndex(i)
                    # QMessageBox.information(self, "Window Already Tabbed", "This window is already in a tab.") # Noisy
                    return
            # QMessageBox.information(self, "Window Already Tracked", "This window is already tracked.") # Noisy
            return

        tab = VBoxTab(window_info)
        self.tabs[hwnd] = tab
        
        idx = self.tab_widget.addTab(tab, tab.title)
        self.tab_widget.setCurrentIndex(idx)
        
        QTimer.singleShot(100, tab.attach_window)

    def show_settings_dialog(self):
        dialog = SettingsDialog(self.settings, self)
        if dialog.exec():
            new_settings = dialog.get_settings(); old_settings = self.settings.copy()
            self.settings = new_settings; self.save_settings()
            if new_settings["refresh_interval"] != old_settings.get("refresh_interval", 5):
                self.auto_refresh_timer.stop(); self.auto_refresh_timer.start(new_settings["refresh_interval"] * 1000)
            if new_settings["theme"] != old_settings.get("theme", "Fusion"): self.change_theme(new_settings["theme"])
            if new_settings["auto_attach"] != old_settings.get("auto_attach", True): self.refresh_tabs()
            if new_settings["dpi_scaling"] != old_settings.get("dpi_scaling", "Auto"):
                QMessageBox.information(self, "Restart Required", "DPI scaling changes will take effect after restarting the application.")

    def close_all_windows(self):
        if not self.tabs: return
        if QMessageBox.question(self, "Close All VMs/Windows", "Forcefully close all tabbed windows?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes: return
        
        tabs_to_close = list(self.tabs.items()); closed_count = 0
        try: subprocess.run(['tskill', 'VBoxSVC'], capture_output=True, text=True, check=False) 
        except FileNotFoundError: pass 

        for hwnd, tab in tabs_to_close:
            try:
                _, process_id = win32process.GetWindowThreadProcessId(hwnd)
                process_handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, process_id)
                win32api.TerminateProcess(process_handle, 0)
                win32api.CloseHandle(process_handle)
                closed_count += 1
            except Exception: pass 
        
        self.tabs.clear()
        while self.tab_widget.count() > 0: self.tab_widget.removeTab(0)
        QMessageBox.information(self, "Windows Closed", f"{closed_count} windows forcefully closed.")
        self.refresh_tabs() 

    def close_tab_by_middle_click(self, index):
        if not (0 <= index < self.tab_widget.count()): return
        tab = self.tab_widget.widget(index)
        if not isinstance(tab, VBoxTab): 
            try: self.tab_widget.removeTab(index); tab.deleteLater()
            except: pass
            return

        hwnd = tab.hwnd; process_id = None
        try:
            if win32gui.IsWindow(hwnd): _, process_id = win32process.GetWindowThreadProcessId(hwnd)
        except: pass
        
        if process_id:
            try:
                process_handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, process_id)
                if process_handle: win32api.TerminateProcess(process_handle, 0); win32api.CloseHandle(process_handle)
            except: pass 
        
        if tab.attached: tab.detach_window() 
        self.tab_widget.removeTab(index)
        if hwnd in self.tabs: del self.tabs[hwnd]
        tab.deleteLater()

    def change_theme(self, theme_name):
        app_instance = QApplication.instance()
        if theme_name == "QDark" and QDARKSTYLE_AVAILABLE: app_instance.setStyleSheet(qdarkstyle.load_stylesheet())
        elif theme_name in self.theme_map and self.theme_map[theme_name] not in ["windows11", "windowsvista", "Windows", "Fusion"] and QT_THEMES_AVAILABLE:
            app_instance.setStyleSheet("") 
            qt_themes.set_theme(self.theme_map[theme_name])
        elif theme_name in self.theme_map: 
            app_instance.setStyleSheet("") 
            QApplication.setStyle(QStyleFactory.create(self.theme_map[theme_name]))
        else: 
            app_instance.setStyleSheet("")
            QApplication.setStyle(QStyleFactory.create("Fusion"))

    def show_about_dialog(self): AboutDialog(self).exec()

    def rename_current_tab(self):
        current_index = self.tab_widget.currentIndex()
        if current_index >= 0:
            tab = self.tab_widget.widget(current_index)
            new_name, ok = QInputDialog.getText(self, "Rename Tab", "Enter new name:", text=tab.title)
            if ok and new_name: tab.title = new_name; self.tab_widget.setTabText(current_index, new_name)

    def detach_current_tab(self):
        current_index = self.tab_widget.currentIndex()
        if current_index >= 0: self.detach_tab(current_index)

    def refresh_tabs(self):
        vbox_windows = self.window_finder.find_virtualbox_windows()
        is_manual_attach_all = (self.sender() == self.attach_button)

        for window in vbox_windows:
            hwnd = window['hwnd']
            
            if hwnd in self.tabs and self.tabs[hwnd].is_custom:
                continue

            was_manually_detached = hwnd in self.manually_detached_windows
            if is_manual_attach_all and was_manually_detached:
                self.manually_detached_windows.remove(hwnd)
                was_manually_detached = False 

            existing_tab = self.tabs.get(hwnd)
            if existing_tab and existing_tab.detached_manually:
                 was_manually_detached = True

            should_attach_vb_window = is_manual_attach_all or \
                                     (self.settings.get("auto_attach", True) and not was_manually_detached)

            if hwnd not in self.tabs: 
                if should_attach_vb_window:
                    tab = VBoxTab(window) 
                    self.tabs[hwnd] = tab
                    idx = self.tab_widget.addTab(tab, tab.title)
                    self.tab_widget.setCurrentIndex(idx) 
                    tab.attach_window()
            elif existing_tab and not existing_tab.attached and should_attach_vb_window: 
                existing_tab.detached_manually = False 
                existing_tab.attach_window()
        
        all_tab_hwnds = list(self.tabs.keys())

        for hwnd in all_tab_hwnds:
            tab = self.tabs.get(hwnd) 
            if not tab: continue

            if not win32gui.IsWindow(hwnd): 
                for i in range(self.tab_widget.count()):
                    if self.tab_widget.widget(i) == tab:
                        self.tab_widget.removeTab(i)
                        break
                del self.tabs[hwnd]
                if hwnd in self.manually_detached_windows:
                    self.manually_detached_windows.remove(hwnd)
                tab.deleteLater() 

    def detach_tab(self, index):
        tab = self.tab_widget.widget(index)
        if not isinstance(tab, VBoxTab): return 

        hwnd = tab.hwnd
        
        if not tab.is_custom and self.settings.get("auto_attach", True):
            self.manually_detached_windows.add(hwnd)
        tab.detached_manually = True 

        if tab.attached: tab.detach_window()
        
        self.tab_widget.removeTab(index)
        if hwnd in self.tabs: del self.tabs[hwnd]
        tab.deleteLater() 

    def closeEvent(self, event):
        # print(f"{time.time():.3f} Main: closeEvent")
        if self.overlay and self.overlay.isVisible(): 
            # print(f"{time.time():.3f} Main: Closing visible overlay during app close.")
            self.overlay.custom_close() # Ensure overlay is properly disposed

        for hwnd, tab in list(self.tabs.items()): 
            if tab.attached: tab.detach_window()
        self.tabs.clear() 
        event.accept()

    def dragEnterEvent(self, event): event.acceptProposedAction()
    def dropEvent(self, event): self.refresh_tabs(); event.acceptProposedAction()

    def close_current_window(self):
        current_index = self.tab_widget.currentIndex()
        if current_index >= 0:
            tab = self.tab_widget.widget(current_index)
            if not isinstance(tab, VBoxTab): return

            hwnd = tab.hwnd
            try:
                _, process_id = win32process.GetWindowThreadProcessId(hwnd)
                process_handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, process_id)
                win32api.TerminateProcess(process_handle, 0)
                win32api.CloseHandle(process_handle)
                
                if tab.attached: tab.detach_window() 
                self.tab_widget.removeTab(current_index)
                if hwnd in self.tabs: del self.tabs[hwnd]
                tab.deleteLater()
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to close window: {str(e)}")


    def open_virtualbox_main(self):
        vbox_path = self.settings["vbox_path"]
        if os.path.exists(vbox_path):
            try: subprocess.Popen([vbox_path]); global VB_MANAGER; VB_MANAGER = True
            except Exception as e: QMessageBox.warning(self, "Error", f"Failed to open VirtualBox: {str(e)}")
        else: QMessageBox.warning(self, "Error", f"VirtualBox executable not found at: {vbox_path}")

    def show_tab_context_menu(self, pos):
        tabBar = self.tab_widget.tabBar()
        index = tabBar.tabAt(pos)
        if index < 0: return
        if not (QApplication.keyboardModifiers() & Qt.ControlModifier): self.tab_widget.setCurrentIndex(index)
        
        tab = self.tab_widget.widget(index)
        if not isinstance(tab, VBoxTab): return

        style = QApplication.style()
        menu = QMenu(self)
        rename_action = QAction(style.standardIcon(QStyle.SP_FileDialogNewFolder), "Rename", self)
        detach_action = QAction(style.standardIcon(QStyle.SP_DialogCancelButton), "Detach", self)
        close_action = QAction(style.standardIcon(QStyle.SP_DialogCloseButton), "Close Window (Force)", self)
        
        menu.addAction(rename_action)
        menu.addAction(detach_action)
        menu.addAction(close_action)
        
        action = menu.exec(tabBar.mapToGlobal(pos))
        if action == rename_action:
            new_name, ok = QInputDialog.getText(self, "Rename Tab", "Enter new name:", text=tab.title)
            if ok and new_name: tab.title = new_name; self.tab_widget.setTabText(index, new_name)
        elif action == detach_action:
            self.detach_tab(index)
        elif action == close_action:
            self.tab_widget.setCurrentIndex(index) 
            self.close_current_window()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    if not QDARKSTYLE_AVAILABLE: print("QDarkStyle not found. Theme 'QDark' unavailable.")
    if not QT_THEMES_AVAILABLE: print("qt-themes not found. Some themes unavailable.")
    window = VirtualBoxTabs()
    window.show()
    sys.exit(app.exec())