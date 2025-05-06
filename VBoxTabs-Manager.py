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
from PySide6.QtCore import Qt, QTimer, Signal, QObject, QSize, QPoint, QSettings
from PySide6.QtGui import QFont, QAction, QMouseEvent, QIcon

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
WS_CAPTION = 0x00C00000
WS_THICKFRAME = 0x00040000
WS_MINIMIZEBOX = 0x00020000
WS_MAXIMIZEBOX = 0x00010000
WS_SYSMENU = 0x00080000

VB_MANAGER = False


class WindowFinder:
    """Class for finding VirtualBox windows"""

    def __init__(self):
        self.virtualbox_windows = []

    def enum_windows_callback(self, hwnd, _):
        """Callback for EnumWindows"""
        if win32gui.IsWindowVisible(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            if window_title and ("[Running]" in window_title or "[Работает]" in window_title) and " Oracle VirtualBox" in window_title:
                # Get window size
                rect = win32gui.GetWindowRect(hwnd)
                width = rect[2] - rect[0]
                height = rect[3] - rect[1]

                # Extract VM name
                vm_name = window_title
                if "[Running] - Oracle VirtualBox" in vm_name:
                    vm_name = vm_name.split(
                        " [Running] - Oracle VirtualBox")[0]
                elif "[Работает] - Oracle VirtualBox" in vm_name:
                    vm_name = vm_name.split(
                        " [Работает] - Oracle VirtualBox")[0]

                # Add window to list
                self.virtualbox_windows.append({
                    'hwnd': hwnd,
                    'title': vm_name,
                    'original_title': window_title,
                    'width': width,
                    'height': height
                })

            elif window_title and ("Oracle VirtualBox " in window_title or "Oracle VirtualBox " in window_title):
                # Get window size
                rect = win32gui.GetWindowRect(hwnd)
                width = rect[2] - rect[0]
                height = rect[3] - rect[1]

                vm_name = "VB Manager"

                # Add window to list
                self.virtualbox_windows.append({
                    'hwnd': hwnd,
                    'title': vm_name,
                    'original_title': window_title,
                    'width': width,
                    'height': height
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
        """Sets parent window for hwnd"""
        # Get current window style
        style = win32gui.GetWindowLong(hwnd, GWL_STYLE)

        # Save current styles
        old_styles = style & (WS_CAPTION | WS_THICKFRAME | WS_MINIMIZEBOX |
                              WS_MAXIMIZEBOX | WS_SYSMENU)

        # Change window style, removing title and frame
        new_style = (style & ~(WS_CAPTION | WS_THICKFRAME | WS_MINIMIZEBOX |
                               WS_MAXIMIZEBOX | WS_SYSMENU)) | WS_CHILD

        win32gui.SetWindowLong(hwnd, GWL_STYLE, new_style)

        # Set new parent window
        win32gui.SetParent(hwnd, parent_hwnd)

        # Update window
        win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                              SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)

        return old_styles

    @staticmethod
    def restore_window_style(hwnd, old_styles):
        """Restores original window style"""
        current_style = win32gui.GetWindowLong(hwnd, GWL_STYLE)

        # Remove WS_CHILD and restore original styles
        new_style = (current_style & ~WS_CHILD) | old_styles

        win32gui.SetWindowLong(hwnd, GWL_STYLE, new_style)
        win32gui.SetParent(hwnd, 0)  # Set parent to None (0)

        # Update window
        win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                              SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)


class SettingsDialog(QDialog):
    """Settings dialog window"""

    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("Settings")
        self.setMinimumWidth(500)
        # Set proper popup dialog flags - make it modal and non-resizable
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint |
                            Qt.WindowTitleHint | Qt.WindowCloseButtonHint)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # General settings group
        general_group = QGroupBox("General Settings")
        general_layout = QGridLayout(general_group)
        general_layout.setColumnStretch(1, 1)

        # Auto-attach windows option
        self.auto_attach_checkbox = QCheckBox(
            "Automatically attach VirtualBox windows")
        self.auto_attach_checkbox.setChecked(
            self.settings.get("auto_attach", True))
        general_layout.addWidget(self.auto_attach_checkbox, 0, 0, 1, 2)

        # Refresh interval
        general_layout.addWidget(
            QLabel("Window detection interval (seconds):"), 1, 0)
        self.refresh_interval_spinbox = QSpinBox()
        self.refresh_interval_spinbox.setRange(1, 60)
        self.refresh_interval_spinbox.setValue(
            self.settings.get("refresh_interval", 5))
        general_layout.addWidget(self.refresh_interval_spinbox, 1, 1)

        # VirtualBox path
        general_layout.addWidget(QLabel("VirtualBox executable path:"), 2, 0)
        vbox_path_layout = QHBoxLayout()
        self.vbox_path_edit = QLineEdit(self.settings.get(
            "vbox_path", r"C:\Program Files\Oracle\VirtualBox\VirtualBox.exe"))
        vbox_path_layout.addWidget(self.vbox_path_edit)
        self.browse_button = QPushButton("Browse...")
        self.browse_button.clicked.connect(self.browse_vbox_path)
        vbox_path_layout.addWidget(self.browse_button)
        general_layout.addLayout(vbox_path_layout, 2, 1)

        main_layout.addWidget(general_group)

        # Display settings group
        display_group = QGroupBox("Display Settings")
        display_layout = QGridLayout(display_group)
        display_layout.setColumnStretch(1, 1)

        # Theme selection
        display_layout.addWidget(QLabel("Theme:"), 0, 0)
        self.theme_combo = QComboBox()

        # Add available themes
        theme_names = ["Dark", "Light", "Classic", "Fusion"]
        if QDARKSTYLE_AVAILABLE:
            theme_names.append("QDark")
        if QT_THEMES_AVAILABLE:
            theme_names.extend(["Atom One", "Blender", "Catppuccin Frappe", "Catppuccin Latte",
                                "Catppuccin Macchiato", "Catppuccin Mocha", "Dracula",
                                "GitHub Dark", "GitHub Light", "Modern Dark", "Modern Light",
                                "Monokai", "Nord", "One Dark Two"])

        self.theme_combo.addItems(theme_names)
        current_theme = self.settings.get("theme", "Fusion")
        if current_theme in theme_names:
            self.theme_combo.setCurrentText(current_theme)
        display_layout.addWidget(self.theme_combo, 0, 1)

        # DPI scaling
        display_layout.addWidget(QLabel("DPI Scaling:"), 1, 0)
        self.dpi_scaling_combo = QComboBox()
        self.dpi_scaling_combo.addItems(
            ["Auto", "100%", "125%", "150%", "175%", "200%"])
        current_scaling = self.settings.get("dpi_scaling", "Auto")
        self.dpi_scaling_combo.setCurrentText(current_scaling)
        display_layout.addWidget(self.dpi_scaling_combo, 1, 1)

        main_layout.addWidget(display_group)

        # Buttons
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
        """Opens file dialog to select VirtualBox executable"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select VirtualBox Executable",
            os.path.dirname(self.vbox_path_edit.text()),
            "VirtualBox Executable (VirtualBox.exe);;All Executable Files (*.exe)"
        )
        if file_path:
            self.vbox_path_edit.setText(file_path)

    def get_settings(self):
        """Returns the settings from the dialog"""
        return {
            "auto_attach": self.auto_attach_checkbox.isChecked(),
            "refresh_interval": self.refresh_interval_spinbox.value(),
            "vbox_path": self.vbox_path_edit.text(),
            "theme": self.theme_combo.currentText(),
            "dpi_scaling": self.dpi_scaling_combo.currentText()
        }


class AboutDialog(QDialog):
    """About dialog window"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About VBoxTabs Manager")
        self.setFixedSize(400, 250)
        # Set proper popup dialog flags - make it modal and non-resizable
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint |
                            Qt.WindowTitleHint | Qt.WindowCloseButtonHint)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Using a scroll area to prevent content from moving during resize
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(8)

        # About information
        title_label = QLabel("VBoxTabs Manager")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(title_label)

        desc_label = QLabel(
            "A tool for combining VirtualBox windows into a single tabbed window.")
        desc_label.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(desc_label)

        version_label = QLabel("Version 1.3")
        version_label.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(version_label)

        author_label = QLabel("Author: Zalexanninev15")
        author_label.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(author_label)

        github_label = QLabel(
            '<a href="https://github.com/Zalexanninev15/VBoxTabs-Manager">GitHub Repository</a>')
        github_label.setOpenExternalLinks(True)
        github_label.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(github_label)

        license_label = QLabel("License: MIT")
        license_label.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(license_label)

        content_layout.addSpacing(10)

        info_label = QLabel(
            "This application allows you to manage multiple VirtualBox machines "
            "in a single window with tabs. It preserves all VirtualBox functionality "
            "while providing a more convenient interface.\n\n"
            "Built with PySide6 (Qt6) and Win32 API."
        )
        info_label.setWordWrap(True)
        info_label.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(info_label)

        main_layout.addWidget(content_widget, 1)

        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(0, 10, 0, 0)

        ok_button = QPushButton("OK")
        ok_button.setFixedWidth(80)
        ok_button.clicked.connect(self.accept)
        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        button_layout.addStretch()

        main_layout.addWidget(button_container)


class VBoxTab(QWidget):
    """Tab widget for VirtualBox window"""

    def __init__(self, window_info, parent=None):
        super().__init__(parent)

        self.theme_map = {
            # Standard Qt themes
            "Dark": "windows11",
            "Light": "windowsvista",
            "Classic": "Windows",
            "Fusion": "Fusion",
            "QDark": "qdarkstyle",

            # qt-themes themes
            "Atom One": "atom_one",
            "Blender": "blender",
            "Catppuccin Frappe": "catppuccin_frappe",
            "Catppuccin Latte": "catppuccin_latte",
            "Catppuccin Macchiato": "catppuccin_macchiato",
            "Catppuccin Mocha": "catppuccin_mocha",
            "Dracula": "dracula",
            "GitHub Dark": "github_dark",
            "GitHub Light": "github_light",
            "Modern Dark": "modern_dark",
            "Modern Light": "modern_light",
            "Monokai": "monokai",
            "Nord": "nord",
            "One Dark Two": "one_dark_two"
        }

        self.window_info = window_info
        self.hwnd = window_info['hwnd']
        self.title = window_info['title']
        self.original_title = window_info['original_title']
        self.orig_styles = None
        self.attached = False
        self.detached_manually = False

        # Create Layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(layout)

        # Container for virtual machine
        self.container = QWidget(self)
        layout.addWidget(self.container)

    def attach_window(self):
        """Attaches VirtualBox window to tab"""
        if not self.attached:
            self.orig_styles = WindowManager.set_window_parent(
                self.hwnd, int(self.container.winId()))
            win32gui.MoveWindow(
                self.hwnd, 0, 0, self.container.width(), self.container.height(), True)
            self.attached = True
            self.detached_manually = False

    def detach_window(self):
        """Detaches VirtualBox window from tab"""
        if self.attached and self.orig_styles is not None:
            WindowManager.restore_window_style(self.hwnd, self.orig_styles)
            win32gui.ShowWindow(self.hwnd, win32con.SW_SHOW)
            self.attached = False
            self.detached_manually = True

    def resizeEvent(self, event):
        """Handles tab resize event"""
        super().resizeEvent(event)
        if self.attached and win32gui.IsWindow(self.hwnd):
            # Resize inner VirtualBox window
            win32gui.MoveWindow(
                self.hwnd, 0, 0, self.container.width(), self.container.height(), True)


class RefreshSignal(QObject):
    """Signal class for tab refresh"""
    refreshRequested = Signal()


class MiddleClickCloseTabBar(QTabBar):
    # Define a custom signal that will emit the index of the tab to close
    middleCloseRequested = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)

    def mouseReleaseEvent(self, event: QMouseEvent):
        # Check if the middle mouse button was released
        if event.button() == Qt.MouseButton.MiddleButton:
            # Find the tab index at the click position
            tab_index = self.tabAt(event.position().toPoint())
            if tab_index != -1:
                # Emit the custom signal with the tab index
                self.middleCloseRequested.emit(tab_index)
                event.accept()  # Indicate we handled this event
                return

        # Call the base class implementation for other mouse events
        super().mouseReleaseEvent(event)


class VirtualBoxTabs(QMainWindow):
    """Main application window"""

    def __init__(self):
        super().__init__()

        # Set to track manually detached windows by hwnd
        self.manually_detached_windows = set()

        # Define theme mapping dictionary
        self.theme_map = {
            # Standard Qt themes
            "Dark": "windows11",
            "Light": "windowsvista",
            "Classic": "Windows",
            "Fusion": "Fusion",
            "QDark": "qdarkstyle",

            # qt-themes themes
            "Atom One": "atom_one",
            "Blender": "blender",
            "Catppuccin Frappe": "catppuccin_frappe",
            "Catppuccin Latte": "catppuccin_latte",
            "Catppuccin Macchiato": "catppuccin_macchiato",
            "Catppuccin Mocha": "catppuccin_mocha",
            "Dracula": "dracula",
            "GitHub Dark": "github_dark",
            "GitHub Light": "github_light",
            "Modern Dark": "modern_dark",
            "Modern Light": "modern_light",
            "Monokai": "monokai",
            "Nord": "nord",
            "One Dark Two": "one_dark_two"
        }

        # Load settings
        self.settings_file = self.get_settings_path()
        self.settings = self.load_settings()

        # Apply DPI scaling settings
        self.apply_dpi_scaling()

        self.setWindowTitle("VBoxTabs Manager")
        self.resize(1280, 800)
        self.setMinimumSize(420, 251)

        # Main icon
        style = QApplication.style()
        self.setWindowIcon(style.standardIcon(QStyle.SP_ComputerIcon))

        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Create main layout
        main_layout = QVBoxLayout(central_widget)

        # Create button panel
        button_layout = QHBoxLayout()

        # Use standard Qt icons
        refresh_icon = style.standardIcon(QStyle.SP_BrowserReload)
        detach_icon = style.standardIcon(QStyle.SP_DialogCancelButton)
        attach_icon = style.standardIcon(QStyle.SP_ArrowUp)
        about_icon = style.standardIcon(QStyle.SP_MessageBoxInformation)
        rename_icon = style.standardIcon(QStyle.SP_FileDialogNewFolder)
        vbox_icon = style.standardIcon(QStyle.SP_ComputerIcon)
        close_icon = style.standardIcon(QStyle.SP_DialogResetButton)
        settings_icon = style.standardIcon(QStyle.SP_FileDialogDetailedView)
        close_all_icon = style.standardIcon(QStyle.SP_DialogCloseButton)

        # Refresh button
        self.refresh_button = QToolButton()
        self.refresh_button.setIcon(refresh_icon)
        self.refresh_button.setToolTip("Refresh VM list")
        self.refresh_button.clicked.connect(self.refresh_tabs)
        button_layout.addWidget(self.refresh_button)

        # Attach all button
        self.attach_button = QToolButton()
        self.attach_button.setIcon(attach_icon)
        self.attach_button.setToolTip("Attach all available VMs")
        self.attach_button.clicked.connect(self.refresh_tabs)
        button_layout.addWidget(self.attach_button)

        # Detach button
        self.detach_button = QToolButton()
        self.detach_button.setIcon(detach_icon)
        self.detach_button.setToolTip("Detach current VM")
        self.detach_button.clicked.connect(self.detach_current_tab)
        button_layout.addWidget(self.detach_button)

        # Close window button
        self.close_window_button = QToolButton()
        self.close_window_button.setIcon(close_icon)
        self.close_window_button.setToolTip("Close current VM window")
        self.close_window_button.clicked.connect(self.close_current_window)
        button_layout.addWidget(self.close_window_button)

        # Close all windows button
        self.close_all_button = QToolButton()
        self.close_all_button.setIcon(close_all_icon)
        self.close_all_button.setToolTip("Close all VM windows")
        self.close_all_button.clicked.connect(self.close_all_windows)
        button_layout.addWidget(self.close_all_button)

        # Rename button
        self.rename_button = QToolButton()
        self.rename_button.setIcon(rename_icon)
        self.rename_button.setToolTip("Rename current tab")
        self.rename_button.clicked.connect(self.rename_current_tab)
        button_layout.addWidget(self.rename_button)

        # VirtualBox main application button
        self.vbox_main_button = QToolButton()
        self.vbox_main_button.setIcon(vbox_icon)
        self.vbox_main_button.setToolTip("Open VirtualBox main application")
        self.vbox_main_button.clicked.connect(self.open_virtualbox_main)
        button_layout.addWidget(self.vbox_main_button)

        # Spacer
        button_layout.addStretch(1)

        # Settings button
        self.settings_button = QToolButton()
        self.settings_button.setIcon(settings_icon)
        self.settings_button.setToolTip("Settings")
        self.settings_button.clicked.connect(self.show_settings_dialog)
        button_layout.addWidget(self.settings_button)

        # About button
        self.about_button = QToolButton()
        self.about_button.setIcon(about_icon)
        self.about_button.setToolTip("About")
        self.about_button.clicked.connect(self.show_about_dialog)
        button_layout.addWidget(self.about_button)

        main_layout.addLayout(button_layout)

        # Create TabWidget
        self.tab_widget = QTabWidget()
        custom_tab_bar = MiddleClickCloseTabBar()
        custom_tab_bar.middleCloseRequested.connect(
            self.close_tab_by_middle_click)
        self.tab_widget.setTabBar(custom_tab_bar)
        self.tab_widget.setTabsClosable(False)
        self.tab_widget.setMovable(True)
        self.tab_widget.tabCloseRequested.connect(self.detach_tab)
        main_layout.addWidget(self.tab_widget)

        # Window finder object
        self.window_finder = WindowFinder()

        # Dictionary for storing tabs
        self.tabs = {}

        # Signal for tab refresh
        self.refresh_signal = RefreshSignal()
        self.refresh_signal.refreshRequested.connect(self.refresh_tabs)

        # Timer for automatic refresh
        self.auto_refresh_timer = QTimer()
        self.auto_refresh_timer.timeout.connect(
            lambda: self.refresh_signal.refreshRequested.emit())
        refresh_interval = self.settings.get("refresh_interval", 5) * 1000
        self.auto_refresh_timer.start(refresh_interval)

        # Apply theme from settings
        theme = self.settings.get("theme", "Fusion")
        self.change_theme(theme)

        # Initialize tabs
        self.refresh_tabs()

        self.setAcceptDrops(True)

        # Context menu for tabs
        self.tab_widget.tabBar().setContextMenuPolicy(Qt.CustomContextMenu)
        self.tab_widget.tabBar().customContextMenuRequested.connect(
            self.show_tab_context_menu)

    def get_settings_path(self):  # Add self parameter here
        # If running as executable (frozen)
        if getattr(sys, 'frozen', False):
            # Get the directory of the executable
            base_dir = os.path.dirname(sys.executable)
        else:
            # Get the directory of the script
            base_dir = os.path.dirname(os.path.abspath(__file__))

        return os.path.join(base_dir, 'settings.json')

    def load_settings(self):
        """Load settings from file"""
        default_settings = {
            "auto_attach": True,
            "refresh_interval": 5,
            "vbox_path": r"C:\Program Files\Oracle\VirtualBox\VirtualBox.exe",
            "theme": "Fusion",
            "dpi_scaling": "Auto"
        }

        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
                    # Merge with defaults to ensure all keys exist
                    for key, value in default_settings.items():
                        if key not in settings:
                            settings[key] = value
                    return settings
        except Exception as e:
            print(f"Error loading settings: {e}")

        return default_settings

    def save_settings(self):
        """Save settings to file"""
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")
            QMessageBox.warning(
                self, "Error", f"Failed to save settings: {str(e)}")

    def apply_dpi_scaling(self):
        """Apply DPI scaling settings"""
        scaling = self.settings.get("dpi_scaling", "Auto")
        if scaling == "Auto":
            os.environ['QT_AUTO_SCREEN_SCALE_FACTOR'] = '1'
        else:
            try:
                scale_factor = float(scaling.strip('%')) / 100.0
                os.environ['QT_SCALE_FACTOR'] = str(scale_factor)
            except ValueError:
                os.environ['QT_AUTO_SCREEN_SCALE_FACTOR'] = '1'

    def show_settings_dialog(self):
        """Show settings dialog"""
        dialog = SettingsDialog(self.settings, self)
        if dialog.exec():
            # Get updated settings
            new_settings = dialog.get_settings()
            old_settings = self.settings.copy()

            # Update settings
            self.settings = new_settings
            self.save_settings()

            # Check if refresh interval changed
            if new_settings["refresh_interval"] != old_settings.get("refresh_interval", 5):
                self.auto_refresh_timer.stop()
                self.auto_refresh_timer.start(
                    new_settings["refresh_interval"] * 1000)

            # Apply theme immediately if changed
            if new_settings["theme"] != old_settings.get("theme", "Fusion"):
                self.change_theme(new_settings["theme"])

            # Apply auto_attach setting immediately
            if new_settings["auto_attach"] != old_settings.get("auto_attach", True):
                # Refresh tabs to apply the new auto_attach setting
                self.refresh_tabs()

            # Show message about DPI scaling changes requiring restart
            if new_settings["dpi_scaling"] != old_settings.get("dpi_scaling", "Auto"):
                QMessageBox.information(self, "Restart Required",
                                        "DPI scaling changes will take effect after restarting the application.")

    def close_all_windows(self):
        """Forcefully close all VM windows"""
        if not self.tabs:
            return

        reply = QMessageBox.question(
            self, "Close All VMs",
            "Are you sure you want to forcefully close all VM windows?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply != QMessageBox.Yes:
            return

        # Store tabs in a list to avoid modification during iteration
        tabs_to_close = list(self.tabs.items())
        closed_count = 0

        # Terminate VBoxSVC process
        try:
            subprocess.run(['tskill', 'VBoxSVC'],
                           capture_output=True, text=True)
        except:
            pass

        for hwnd, tab in tabs_to_close:
            try:
                # Get the process ID associated with the window
                _, process_id = win32process.GetWindowThreadProcessId(hwnd)

                # Open the process with termination rights
                process_handle = win32api.OpenProcess(
                    win32con.PROCESS_TERMINATE, False, process_id)

                # Terminate the process
                win32api.TerminateProcess(process_handle, 0)
                win32api.CloseHandle(process_handle)
                closed_count += 1

            except:
                pass

        # Clear tabs dictionary and remove all tabs from UI
        self.tabs.clear()
        while self.tab_widget.count() > 0:
            self.tab_widget.removeTab(0)

        QMessageBox.information(
            self, "VMs Closed", f"{closed_count} VM windows have been forcefully closed.")

        # Refresh tabs to ensure UI is updated
        self.refresh_tabs()

    def close_tab_by_middle_click(self, index):
        """Forcefully closes the VM window associated with the tab at the given index."""
        if not (0 <= index < self.tab_widget.count()):
            return

        tab = self.tab_widget.widget(index)
        if not isinstance(tab, VBoxTab):
            try:
                widget_to_close = self.tab_widget.widget(index)
                if widget_to_close:
                    self.tab_widget.removeTab(index)
                    widget_to_close.deleteLater()
            except Exception:
                pass
            return

        hwnd = tab.hwnd
        process_id = None

        # Get process ID
        try:
            if win32gui.IsWindow(hwnd):
                _, process_id = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:
            pass
        # Terminate the process if we have an ID
        terminated_ok = False
        if process_id:
            try:
                process_handle = win32api.OpenProcess(
                    win32con.PROCESS_TERMINATE, False, process_id)
                if process_handle:
                    win32api.TerminateProcess(process_handle, 0)
                    win32api.CloseHandle(process_handle)
                    terminated_ok = True
            except Exception as e:
                # Could fail if process already exited or due to permissions
                # Check if window still exists after failed termination attempt
                if not win32gui.IsWindow(hwnd):
                    terminated_ok = True

        # Remove the tab from the UI and internal tracking
        try:
            # Ensure the tab widget itself exists and index is still valid
            # (could change during async operations, though unlikely here)
            if self.tab_widget and 0 <= index < self.tab_widget.count() and self.tab_widget.widget(index) == tab:
                self.tab_widget.removeTab(index)
                if hwnd in self.tabs:
                    del self.tabs[hwnd]
                tab.deleteLater()  # Schedule the QWidget for deletion
            else:
                # Clean up self.tabs just in case
                if hwnd in self.tabs:
                    del self.tabs[hwnd]

            QMessageBox.information(
                self, "VMs Closed", "VM window have been forcefully closed.")
        except Exception as e:
            # Attempt cleanup of internal state even if UI removal failed
            if hwnd in self.tabs:
                del self.tabs[hwnd]

    def change_theme(self, theme_name):
        """Changes application theme"""
        if theme_name == "QDark" and QDARKSTYLE_AVAILABLE:
            # Apply QDarkStyle
            QApplication.instance().setStyleSheet(qdarkstyle.load_stylesheet())
        elif theme_name in ["Atom One", "Blender", "Catppuccin Frappe", "Catppuccin Latte",
                            "Catppuccin Macchiato", "Catppuccin Mocha", "Dracula",
                            "GitHub Dark", "GitHub Light", "Modern Dark", "Modern Light",
                            "Monokai", "Nord", "One Dark Two"] and QT_THEMES_AVAILABLE:
            # Apply qt-themes theme
            QApplication.instance().setStyleSheet("")
            qt_themes.set_theme(self.theme_map[theme_name])
        else:
            # For other themes, first clear any stylesheet
            QApplication.instance().setStyleSheet("")
            # Then apply the Qt style
            QApplication.setStyle(
                QStyleFactory.create(self.theme_map[theme_name]))
            # Reset any qt-themes theme by setting empty stylesheet
            if QT_THEMES_AVAILABLE:
                # qt_themes doesn't have reset_theme method, use set_theme with empty string
                QApplication.instance().setStyleSheet("")

    def show_about_dialog(self):
        """Shows about dialog"""
        about_dialog = AboutDialog(self)
        about_dialog.exec()

    def rename_current_tab(self):
        """Renames the current tab"""
        current_index = self.tab_widget.currentIndex()
        if current_index >= 0:
            tab = self.tab_widget.widget(current_index)
            current_name = tab.title

            # Show input dialog for new name
            new_name, ok = QInputDialog.getText(
                self, "Rename Tab", "Enter new name:", text=current_name
            )

            if ok and new_name:
                tab.title = new_name
                self.tab_widget.setTabText(current_index, new_name)

    def detach_current_tab(self):
        """Detaches the current tab"""
        current_index = self.tab_widget.currentIndex()
        if current_index >= 0:
            self.detach_tab(current_index)

    def refresh_tabs(self):
        """Refreshes tabs with VirtualBox windows"""
        vbox_windows = self.window_finder.find_virtualbox_windows()

        # Determine if this refresh was triggered by the attach button
        is_manual_attach = (self.sender() == self.attach_button)

        for window in vbox_windows:
            hwnd = window['hwnd']

            # Check if this window was manually detached previously
            was_manually_detached = hwnd in self.manually_detached_windows

            # If this is a manual attach action, remove from detached list
            if is_manual_attach and was_manually_detached:
                self.manually_detached_windows.remove(hwnd)
                was_manually_detached = False

            # Check if this window already exists and was manually detached
            existing_tab = self.tabs.get(hwnd, None)
            if existing_tab and existing_tab.detached_manually:
                was_manually_detached = True

            # Determine if this window should be attached
            # Only attach if: manual attach OR (auto-attach is enabled AND not manually detached before)
            should_attach = is_manual_attach or (self.settings.get(
                "auto_attach", True) and not was_manually_detached)

            # Only add new tabs if auto-attach is enabled or if manually attaching
            if hwnd not in self.tabs and (should_attach or is_manual_attach):
                tab = VBoxTab(window)
                self.tabs[hwnd] = tab
                self.tab_widget.addTab(tab, tab.title)
                if should_attach:
                    tab.attach_window()
            # If tab exists but isn't attached and should be attached now
            elif hwnd in self.tabs and is_manual_attach and not self.tabs[hwnd].attached:
                # Reset the detached_manually flag when manually attaching
                self.tabs[hwnd].detached_manually = False
                self.tabs[hwnd].attach_window()

    def detach_tab(self, index):
        """Detaches tab and closes it"""
        tab = self.tab_widget.widget(index)
        hwnd = tab.hwnd

        # Set the detached_manually flag before detaching
        tab.detached_manually = True
        # Add to manually detached windows set to remember it
        self.manually_detached_windows.add(hwnd)
        # Detach VirtualBox window
        tab.detach_window()

        # Remove tab
        self.tab_widget.removeTab(index)
        del self.tabs[hwnd]

    def closeEvent(self, event):
        """Handles application window close event"""
        # Always deattach all windows without a dialog box
        for hwnd, tab in list(self.tabs.items()):
            tab.detach_window()
        event.accept()

    # --- Drag & Drop to attach VirtualBox window ---
    def dragEnterEvent(self, event):
        # Allow Drag & Drop always
        event.acceptProposedAction()

    def dropEvent(self, event):
        # Here you can implement attachment of a window by hwnd if you pass hwnd through mimeData
        self.refresh_tabs()
        event.acceptProposedAction()

    def close_current_window(self):
        """Forcefully closes the current VM window, terminates VBoxSVC, and removes its tab"""
        current_index = self.tab_widget.currentIndex()
        if current_index >= 0:
            tab = self.tab_widget.widget(current_index)
            hwnd = tab.hwnd

            # Get the process ID associated with the window
            _, process_id = win32process.GetWindowThreadProcessId(hwnd)

            try:
                # Open the process with termination rights
                process_handle = win32api.OpenProcess(
                    win32con.PROCESS_TERMINATE, False, process_id)
                # Terminate the process
                win32api.TerminateProcess(process_handle, 0)
                win32api.CloseHandle(process_handle)

                # Remove the tab
                self.tab_widget.removeTab(current_index)
                del self.tabs[hwnd]

                QMessageBox.information(
                    self, "VMs Closed", "VM window have been forcefully closed.")
            except Exception as e:
                QMessageBox.warning(
                    self, "Error", f"Failed to close VM window: {str(e)}")

    def open_virtualbox_main(self):
        """Opens the main VirtualBox application"""
        vbox_path = self.settings["vbox_path"]
        if os.path.exists(vbox_path):
            try:
                subprocess.Popen([vbox_path])
                VB_MANAGER = True
            except Exception as e:
                QMessageBox.warning(
                    self, "Error", f"Failed to open VirtualBox: {str(e)}")
        else:
            QMessageBox.warning(
                self, "Error", f"VirtualBox executable not found at the specified path.\nPlease check the path in Settings:\n{vbox_path}")

    # --- Context menu for tabs ---
    def show_tab_context_menu(self, pos):
        tabBar = self.tab_widget.tabBar()
        index = tabBar.tabAt(pos)
        if index < 0:
            return

        # Check if Ctrl is pressed
        ctrl_pressed = QApplication.keyboardModifiers() & Qt.ControlModifier

        # If Ctrl is NOT pressed, make the clicked tab active
        if not ctrl_pressed:
            self.tab_widget.setCurrentIndex(index)

        style = QApplication.style()
        rename_icon = style.standardIcon(QStyle.SP_FileDialogNewFolder)
        detach_icon = style.standardIcon(QStyle.SP_DialogCancelButton)
        close_icon = style.standardIcon(QStyle.SP_DialogResetButton)

        menu = QMenu(self)
        rename_action = QAction(rename_icon, "Rename", self)
        detach_action = QAction(detach_icon, "Detach", self)
        close_window_action = QAction(close_icon, "Close window", self)
        menu.addAction(rename_action)
        menu.addAction(detach_action)
        menu.addAction(close_window_action)
        action = menu.exec(tabBar.mapToGlobal(pos))
        if action == rename_action:
            # Rename the tab at 'index', not necessarily the current one
            tab = self.tab_widget.widget(index)
            current_name = tab.title
            new_name, ok = QInputDialog.getText(
                self, "Rename Tab", "Enter new name:", text=current_name
            )
            if ok and new_name:
                tab.title = new_name
                self.tab_widget.setTabText(index, new_name)
        elif action == detach_action:
            self.detach_tab(index)
        elif action == close_window_action:
            # Set the current index to the tab that was right-clicked
            self.tab_widget.setCurrentIndex(index)
            # Close the window
            self.close_current_window()

    # --- Detach when dragging a tab beyond --- Roadmap: Quickly deattach windows...
    # def mouseReleaseEvent(self, event):
        # Check if the mouse is released outside the tab area and there is a draggable tab
        # if event.button() == Qt.LeftButton:
            # tabBar = self.tab_widget.tabBar()
            # global_pos = event.globalPosition().toPoint() if hasattr(
            # event, "globalPosition") else event.globalPos()
            # tabBar_rect = tabBar.rect()
            # tabBar_global = tabBar.mapToGlobal(
            # tabBar_rect.topLeft()), tabBar.mapToGlobal(tabBar_rect.bottomRight())
            # if not (tabBar_global[0].x() <= global_pos.x() <= tabBar_global[1].x() and
            # tabBar_global[0].y() <= global_pos.y() <= tabBar_global[1].y()):
            # The tab has been dragged outside the tab bar
            # index = self.tab_widget.currentIndex()
            # if index >= 0:
            # self.detach_tab(index)
        # super().mouseReleaseEvent(event)


if __name__ == "__main__":
    # Create application first to be able to access QApplication.instance() later
    app = QApplication(sys.argv)

    if not QDARKSTYLE_AVAILABLE:
        print("QDarkStyle not found. The theme 'QDark' is unavailable!")
    if not QT_THEMES_AVAILABLE:
        print("qt-themes not found. Some themes are unavailable!")

    window = VirtualBoxTabs()
    window.show()

    sys.exit(app.exec())
