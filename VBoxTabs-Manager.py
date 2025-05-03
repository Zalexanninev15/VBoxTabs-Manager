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
import time
import win32gui
import win32process
import win32con
import win32api
import ctypes
from ctypes import wintypes
from PySide6.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, 
                              QVBoxLayout, QPushButton, QMessageBox, QLabel,
                              QDialog, QCheckBox, QHBoxLayout, QInputDialog,
                              QStyleFactory, QComboBox, QToolButton, QMenu,
                              QToolTip, QStyle)
from PySide6.QtCore import Qt, QTimer, Signal, QObject, QSize, QPoint
from PySide6.QtGui import QIcon, QFont, QAction  # QPixmap, base64 удалены

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


class WindowFinder:
    """Class for finding VirtualBox windows"""
    
    def __init__(self):
        self.virtualbox_windows = []
    
    def enum_windows_callback(self, hwnd, _):
        """Callback for EnumWindows"""
        if win32gui.IsWindowVisible(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            # Support both English and Russian status indicators
            if window_title and ("[Running]" in window_title or "[Работает]" in window_title) and "Oracle VirtualBox" in window_title:
                # Get window size
                rect = win32gui.GetWindowRect(hwnd)
                width = rect[2] - rect[0]
                height = rect[3] - rect[1]
                
                # Extract VM name (remove " - Oracle VirtualBox")
                vm_name = window_title
                if " - Oracle VirtualBox" in vm_name:
                    vm_name = vm_name.split(" - Oracle VirtualBox")[0]
                
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


class AboutDialog(QDialog):
    """About dialog window"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About VBoxTabs Manager")
        self.setFixedSize(400, 250)

        layout = QVBoxLayout(self)
        
        # Add about information
        title_label = QLabel("VBoxTabs Manager")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        desc_label = QLabel("A tool for combining VirtualBox windows into a single tabbed window.")
        desc_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(desc_label)
        
        version_label = QLabel("Version 1.01")
        version_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(version_label)
        
        author_label = QLabel("Author: Zalexanninev15")
        author_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(author_label)

        github_label = QLabel('<a href="https://github.com/Zalexanninev15/VBoxTabs-Manager">GitHub Repository</a>')
        github_label.setOpenExternalLinks(True)
        github_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(github_label)

        license_label = QLabel("License: MIT")
        license_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(license_label)
        
        layout.addSpacing(10)
        
        info_label = QLabel(
            "This application allows you to manage multiple VirtualBox machines "
            "in a single window with tabs. It preserves all VirtualBox functionality "
            "while providing a more convenient interface.\n\n"
            "Built with PySide6 (Qt6) and Win32 API."
        )
        info_label.setWordWrap(True)
        info_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(info_label)
        
        # Add OK button
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        layout.addWidget(ok_button)

class VBoxTab(QWidget):
    """Tab widget for VirtualBox window"""
    
    def __init__(self, window_info, parent=None):
        super().__init__(parent)
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
        
        # Add container for virtual machine
        self.container = QWidget(self)
        layout.addWidget(self.container)
    
    def attach_window(self):
        """Attaches VirtualBox window to tab"""
        if not self.attached:
            self.orig_styles = WindowManager.set_window_parent(self.hwnd, int(self.container.winId()))
            win32gui.MoveWindow(self.hwnd, 0, 0, self.container.width(), self.container.height(), True)
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
            win32gui.MoveWindow(self.hwnd, 0, 0, self.container.width(), self.container.height(), True)


class RefreshSignal(QObject):
    """Signal class for tab refresh"""
    refreshRequested = Signal()


class VirtualBoxTabs(QMainWindow):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("VBoxTabs Manager")
        self.resize(1280, 800)
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
        attach_icon = style.standardIcon(QStyle.SP_DialogOpenButton)
        about_icon = style.standardIcon(QStyle.SP_MessageBoxInformation)
        rename_icon = style.standardIcon(QStyle.SP_FileDialogNewFolder)
        
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
        
        # Rename button
        self.rename_button = QToolButton()
        self.rename_button.setIcon(rename_icon)
        self.rename_button.setToolTip("Rename current tab")
        self.rename_button.clicked.connect(self.rename_current_tab)
        button_layout.addWidget(self.rename_button)
        
        # Add spacer
        button_layout.addStretch(1)
        
        # Theme selector
        theme_label = QLabel("Theme:")
        button_layout.addWidget(theme_label)
        
        self.theme_selector = QComboBox()
        self.theme_selector.addItems(QStyleFactory.keys())
        self.theme_selector.currentTextChanged.connect(self.change_theme)
        button_layout.addWidget(self.theme_selector)
        
        # About button
        self.about_button = QToolButton()
        self.about_button.setIcon(about_icon)
        self.about_button.setToolTip("About")
        self.about_button.clicked.connect(self.show_about_dialog)
        button_layout.addWidget(self.about_button)
        
        # Add button panel to main layout
        main_layout.addLayout(button_layout)
        
        # Create TabWidget
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)
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
        self.auto_refresh_timer.timeout.connect(lambda: self.refresh_signal.refreshRequested.emit())
        self.auto_refresh_timer.start(5000)  # Refresh every 5 seconds
        
        # Initialize tabs
        self.refresh_tabs()

        self.setAcceptDrops(True)

        # Context menu for tabs
        self.tab_widget.tabBar().setContextMenuPolicy(Qt.CustomContextMenu)
        self.tab_widget.tabBar().customContextMenuRequested.connect(self.show_tab_context_menu)
    
    def change_theme(self, theme_name):
        """Changes application theme"""
        QApplication.setStyle(QStyleFactory.create(theme_name))
    
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
        for window in vbox_windows:
            hwnd = window['hwnd']
            if hwnd not in self.tabs:
                tab = VBoxTab(window)
                self.tabs[hwnd] = tab
                self.tab_widget.addTab(tab, tab.title)
                # Attach only if not manually detached
                if not getattr(tab, "detached_manually", False):
                    tab.attach_window()
    
    def detach_tab(self, index):
        """Detaches tab and closes it"""
        tab = self.tab_widget.widget(index)
        hwnd = tab.hwnd
        
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
        close_icon = style.standardIcon(QStyle.SP_TitleBarCloseButton)

        menu = QMenu(self)
        rename_action = QAction(rename_icon, "Rename", self)
        detach_action = QAction(detach_icon, "Detach", self)
        close_action = QAction(close_icon, "Close tab", self)
        menu.addAction(rename_action)
        menu.addAction(detach_action)
        menu.addAction(close_action)
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
        elif action == close_action:
            self.detach_tab(index)

    # --- Detach when dragging a tab beyond ---
    def mouseReleaseEvent(self, event):
        # Check if the mouse is released outside the tab area and there is a draggable tab
        if event.button() == Qt.LeftButton:
            tabBar = self.tab_widget.tabBar()
            global_pos = event.globalPosition().toPoint() if hasattr(event, "globalPosition") else event.globalPos()
            tabBar_rect = tabBar.rect()
            tabBar_global = tabBar.mapToGlobal(tabBar_rect.topLeft()), tabBar.mapToGlobal(tabBar_rect.bottomRight())
            if not (tabBar_global[0].x() <= global_pos.x() <= tabBar_global[1].x() and
                    tabBar_global[0].y() <= global_pos.y() <= tabBar_global[1].y()):
                # The tab has been dragged outside the tab bar
                index = self.tab_widget.currentIndex()
                if index >= 0:
                    self.detach_tab(index)
        super().mouseReleaseEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VirtualBoxTabs()
    window.show()
    sys.exit(app.exec())
