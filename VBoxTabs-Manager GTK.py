# VBoxTabs-Manager-gtk.py
# VBoxTabs Manager (GTK4/Adwaita Version)
# Original Copyright (c) 2025 Zalexanninev15
# https://github.com/Zalexanninev15/VBoxTabs-Manager
#
# Licensed under the MIT License.
"""
VBoxTabs Manager - Combining VirtualBox windows into a single tabbed window, using GTK4/Adwaita.
"""

import sys
import os
import subprocess
import json
import time # For ensuring widget is realized

# Win32 imports
import win32gui
import win32process
import win32con
import win32api

# GTK imports
import gi
gi.require_version('Gtk', '4.0')
gi.require_version('Adw', '1')
gi.require_version('GdkWin32', '4.0')
from gi.repository import Gtk, Adw, Gio, GLib, Gdk, GdkWin32

# --- Start of Win32 Utilities (adapted from original) ---
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

APP_ID = "com.zalexanninev15.vboxtabsmanager"
APP_TITLE = "VBoxTabs Manager"

class WindowFinder:
    def __init__(self):
        self.virtualbox_windows = []

    def enum_windows_callback(self, hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == os.getpid():
                    return True
            except: 
                if APP_TITLE in window_title: 
                    return True

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
                    'hwnd': hwnd, 'title': vm_name.strip(), 'original_title': window_title,
                    'width': width, 'height': height
                })
        return True

    def find_virtualbox_windows(self):
        self.virtualbox_windows = []
        win32gui.EnumWindows(self.enum_windows_callback, None)
        return self.virtualbox_windows

class WindowManager:
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
# --- End of Win32 Utilities ---

def apply_global_font_style(font_family="Segoe UI", font_size="9pt"):
    """Applies a global font style using GTK CSS."""
    css = f"* {{ font-family: '{font_family}'; font-size: {font_size}; }}" # MODIFIED
    
    provider = Gtk.CssProvider()
    provider.load_from_data(css.encode('utf-8'))

    display = Gdk.Display.get_default()
    if display:
        Gtk.StyleContext.add_provider_for_display(
            display,
            provider,
            Gtk.STYLE_PROVIDER_PRIORITY_APPLICATION 
        )
        print(f"Applied global font style: {font_family} {font_size}")
    else:
        print("Warning: Could not get default GDK display to apply font style.")


class SettingsDialog(Adw.PreferencesDialog):
    def __init__(self, parent_window, settings_obj):
        super().__init__() 
        
        if parent_window and isinstance(parent_window, Gtk.Window):
             print(f"Info: SettingsDialog created with parent {parent_window}. Transience not explicitly set due to API issues.")
        
        self.settings_obj = settings_obj 
        self.set_title("Settings")
        self.set_size_request(500, -1) # MODIFIED: Use Gtk.Widget's method for basic sizing hint

        page = Adw.PreferencesPage()
        self.add(page)

        general_group = Adw.PreferencesGroup(title="General Settings")
        page.add(general_group)

        self.auto_attach_switch = Gtk.Switch()
        self.auto_attach_switch.set_valign(Gtk.Align.CENTER)
        self.auto_attach_switch.set_active(self.settings_obj.get("auto_attach", True))
        row_auto_attach = Adw.ActionRow(title="Automatically attach VirtualBox windows")
        row_auto_attach.add_suffix(self.auto_attach_switch)
        row_auto_attach.set_activatable_widget(self.auto_attach_switch)
        general_group.add(row_auto_attach)

        self.refresh_interval_spin = Gtk.SpinButton()
        self.refresh_interval_spin.set_range(1, 60)
        self.refresh_interval_spin.set_increments(1, 5)
        self.refresh_interval_spin.set_value(self.settings_obj.get("refresh_interval", 5))
        self.refresh_interval_spin.set_valign(Gtk.Align.CENTER)
        row_refresh = Adw.ActionRow(title="Window detection interval (seconds)")
        row_refresh.add_suffix(self.refresh_interval_spin)
        row_refresh.set_activatable_widget(self.refresh_interval_spin)
        general_group.add(row_refresh)

        self.vbox_path_entry = Gtk.Entry()
        self.vbox_path_entry.set_text(self.settings_obj.get("vbox_path", r"C:\Program Files\Oracle\VirtualBox\VirtualBox.exe"))
        self.vbox_path_entry.set_hexpand(True)
        browse_button = Gtk.Button(label="Browse...")
        browse_button.connect("clicked", self.on_browse_vbox_path)
        
        path_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        path_box.append(self.vbox_path_entry)
        path_box.append(browse_button)

        row_vbox_path = Adw.ActionRow(title="VirtualBox executable path")
        row_vbox_path.add_suffix(path_box)
        general_group.add(row_vbox_path)

        display_group = Adw.PreferencesGroup(title="Display Settings")
        page.add(display_group)

        self.color_scheme_combo = Gtk.ComboBoxText()
        self.schemes_with_ids = { 
            "system_default": "System Default",
            "force_light": "Prefer Light",
            "force_dark": "Prefer Dark"
        }
        
        for id_str, name_str in self.schemes_with_ids.items():
            self.color_scheme_combo.append(id_str, name_str) 
        
        current_scheme_name = self.settings_obj.get("color_scheme", "System Default")
        active_id_to_set = "system_default" 
        for id_val, name_val in self.schemes_with_ids.items():
            if name_val == current_scheme_name:
                active_id_to_set = id_val
                break
        
        self.color_scheme_combo.set_active_id(active_id_to_set) 

        row_color_scheme = Adw.ActionRow(title="Color Scheme")
        row_color_scheme.add_suffix(self.color_scheme_combo)
        row_color_scheme.set_activatable_widget(self.color_scheme_combo)
        display_group.add(row_color_scheme)

    def on_browse_vbox_path(self, button):
        dialog = Gtk.FileChooserDialog(
            title="Select VirtualBox Executable",
            transient_for=self, 
            action=Gtk.FileChooserAction.OPEN
        )
        dialog.add_buttons(
            "_Cancel", Gtk.ResponseType.CANCEL,
            "_Open", Gtk.ResponseType.ACCEPT
        )
        
        filter_exe = Gtk.FileFilter()
        filter_exe.set_name("Executable files (*.exe)")
        filter_exe.add_pattern("*.exe")
        dialog.add_filter(filter_exe)

        current_path = self.vbox_path_entry.get_text()
        if current_path and os.path.exists(os.path.dirname(current_path)):
            try:
                dialog.set_current_folder(Gio.File.new_for_path(os.path.dirname(current_path)))
            except GLib.Error as e:
                print(f"Error setting current folder for file chooser: {e}")
        
        def on_response(d, response_id):
            if response_id == Gtk.ResponseType.ACCEPT:
                file = d.get_file()
                if file:
                    self.vbox_path_entry.set_text(file.get_path())
            d.destroy()
        
        dialog.connect("response", on_response)
        dialog.show()

    def get_settings(self):
         active_id = self.color_scheme_combo.get_active_id()
         color_scheme_name = self.schemes_with_ids.get(active_id, "System Default") 

         return {
            "auto_attach": self.auto_attach_switch.get_active(),
            "refresh_interval": int(self.refresh_interval_spin.get_value()),
            "vbox_path": self.vbox_path_entry.get_text(),
            "color_scheme": color_scheme_name
        }

class VBoxTabGtk(Gtk.Box):
    def __init__(self, window_info):
        super().__init__(orientation=Gtk.Orientation.VERTICAL)
        self.window_info = window_info
        self.hwnd = window_info['hwnd']
        self.vm_title = window_info['title'] 
        self.original_title = window_info['original_title']
        self.orig_styles = None
        self.attached = False
        self.detached_manually = False

        self.vm_container = Gtk.Box() 
        self.vm_container.set_hexpand(True)
        self.vm_container.set_vexpand(True)
        self.vm_container.set_can_focus(True) 
        self.append(self.vm_container)

        self.vm_container.connect("size-allocate", self.on_vm_container_resize)
        self.vm_container.connect("realize", self.do_attach_if_pending)
        self.pending_attach = False

    def do_attach_if_pending(self, widget):
        if self.pending_attach:
            GLib.timeout_add(50, self._attempt_attach_window) 
            self.pending_attach = False
        
    def _attempt_attach_window(self):
        self.attach_window()
        return GLib.SOURCE_REMOVE

    def attach_window(self):
        if self.attached or not self.hwnd or not win32gui.IsWindow(self.hwnd):
            return

        if not self.vm_container.get_realized():
            self.pending_attach = True
            print(f"Deferring attach for '{self.vm_title}': container not realized.")
            return

        gdk_surface = self.vm_container.get_surface()
        if not gdk_surface:
            print(f"Error: VM container for '{self.vm_title}' has no GDK surface.")
            return

        try:
            parent_hwnd = GdkWin32.Win32Surface.get_handle(gdk_surface)
        except Exception as e:
            print(f"Error getting parent HWND for '{self.vm_title}': {e}")
            return

        if parent_hwnd == 0:
            print(f"Error: Parent HWND is 0 for '{self.vm_title}'.")
            return

        try:
            self.orig_styles = WindowManager.set_window_parent(self.hwnd, parent_hwnd)
            allocation = self.vm_container.get_allocation()
            win32gui.MoveWindow(self.hwnd, 0, 0, allocation.width, allocation.height, True)
            win32gui.ShowWindow(self.hwnd, win32con.SW_SHOW)
            self.attached = True
            self.detached_manually = False
            print(f"Attached '{self.vm_title}' (HWND: {self.hwnd}) to GTK parent (HWND: {parent_hwnd})")
        except Exception as e:
            print(f"Error during final attach/move of '{self.vm_title}': {e}")
            if self.orig_styles is not None and win32gui.IsWindow(self.hwnd):
                 WindowManager.restore_window_style(self.hwnd, self.orig_styles)
            self.attached = False


    def detach_window(self, manual=True):
        if not self.attached or self.orig_styles is None or not win32gui.IsWindow(self.hwnd):
            return
        
        try:
            WindowManager.restore_window_style(self.hwnd, self.orig_styles)
            win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE)
        except Exception as e:
            print(f"Error detaching '{self.vm_title}' (HWND: {self.hwnd}): {e}")

        self.attached = False
        if manual:
            self.detached_manually = True
        print(f"Detached '{self.vm_title}' (HWND: {self.hwnd})")

    def on_vm_container_resize(self, widget, allocation):
        if self.attached and win32gui.IsWindow(self.hwnd):
            try:
                win32gui.MoveWindow(self.hwnd, 0, 0, allocation.width, allocation.height, True)
            except Exception as e:
                if win32gui.IsWindow(self.hwnd):
                    print(f"Error resizing HWND {self.hwnd} ('{self.vm_title}'): {e}")
                if not win32gui.IsWindow(self.hwnd):
                    self.detach_window(manual=False) 

    def get_vm_title(self):
        return self.vm_title

    def set_vm_title(self, new_title):
        self.vm_title = new_title


class VBoxTabsWindow(Adw.ApplicationWindow):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.settings_file = self.get_settings_path()
        self.settings = self.load_settings()
        self.apply_color_scheme()

        self.set_title(APP_TITLE) 
        self.set_default_size(1280, 800)

        self.window_finder = WindowFinder()
        self.tabs_by_hwnd = {} 
        self.manually_detached_windows = set() 

        self.main_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        
        self.header_bar = Adw.HeaderBar()
        self.main_box.append(self.header_bar)

        button_content_refresh = Adw.ButtonContent(icon_name="view-refresh-symbolic", label="Refresh")
        self.refresh_button = Gtk.Button()
        self.refresh_button.set_child(button_content_refresh)
        self.refresh_button.set_tooltip_text("Refresh VM list and attach new VMs")
        self.refresh_button.connect("clicked", lambda _: self.refresh_tabs_ui(manual_attach_all=False))
        self.header_bar.pack_start(self.refresh_button)
        
        button_content_attach_all = Adw.ButtonContent(icon_name="list-add-symbolic", label="Attach All")
        self.attach_all_button = Gtk.Button()
        self.attach_all_button.set_child(button_content_attach_all)
        self.attach_all_button.set_tooltip_text("Attempt to attach all found VirtualBox windows")
        self.attach_all_button.connect("clicked", lambda _: self.refresh_tabs_ui(manual_attach_all=True))
        self.header_bar.pack_start(self.attach_all_button)

        button_content_detach = Adw.ButtonContent(icon_name="window-close-symbolic", label="Detach")
        self.detach_button = Gtk.Button()
        self.detach_button.set_child(button_content_detach)
        self.detach_button.set_tooltip_text("Detach current VM")
        self.detach_button.connect("clicked", self.detach_current_tab_ui)
        self.header_bar.pack_start(self.detach_button)
        
        button_content_close_vm = Adw.ButtonContent(icon_name="process-stop-symbolic", label="Close VM")
        self.close_vm_button = Gtk.Button()
        self.close_vm_button.set_child(button_content_close_vm)
        self.close_vm_button.set_tooltip_text("Force close current VM")
        self.close_vm_button.connect("clicked", self.close_current_vm_window)
        self.header_bar.pack_start(self.close_vm_button)

        button_content_vbox = Adw.ButtonContent(icon_name="utilities-terminal-symbolic", label="VBox")
        self.vbox_main_button = Gtk.Button()
        self.vbox_main_button.set_child(button_content_vbox)
        self.vbox_main_button.set_tooltip_text("Open VirtualBox main application")
        self.vbox_main_button.connect("clicked", self.open_virtualbox_main)
        self.header_bar.pack_start(self.vbox_main_button)

        menu_model = Gio.Menu()
        menu_model.append("Rename Tab", "win.rename_current_tab")
        menu_model.append("Close All VMs", "win.close_all_vms")
        
        section_menu = Gio.Menu()
        section_menu.append("Settings", "win.show_settings")
        section_menu.append("About", "win.show_about")
        menu_model.append_section(None, section_menu)
        
        self.menu_button = Gtk.MenuButton(icon_name="open-menu-symbolic", menu_model=menu_model)
        self.header_bar.pack_end(self.menu_button)

        self.notebook = Gtk.Notebook()
        self.notebook.set_scrollable(True)
        self.notebook.set_hexpand(True)
        self.notebook.set_vexpand(True)
        self.main_box.append(self.notebook)

        self.set_content(self.main_box)

        self.create_action("rename_current_tab", self.rename_current_tab_ui)
        self.create_action("close_all_vms", self.close_all_windows_ui)
        self.create_action("show_settings", self.show_settings_dialog_action_cb)
        self.create_action("show_about", self.show_about_dialog_action_cb)
        
        refresh_interval_ms = self.settings.get("refresh_interval", 5) * 1000
        self.auto_refresh_timer_id = GLib.timeout_add(refresh_interval_ms, self.on_auto_refresh_timer)

        self.connect("close-request", self.on_window_close_request)
        GLib.idle_add(lambda: self.refresh_tabs_ui() or GLib.SOURCE_REMOVE)

        self.notebook.connect("page-added", self.on_notebook_page_added)
        
        drop_target = Gtk.DropTarget.new(str, Gdk.DragAction.COPY)
        drop_target.connect("drop", lambda s, v, x, y: self.refresh_tabs_ui() or True)
        self.add_controller(drop_target)

    def create_action(self, name, callback):
        action = Gio.SimpleAction.new(name, None)
        action.connect("activate", lambda a, p: callback())
        self.add_action(action)

    def get_settings_path(self):
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, 'settings.json')

    def load_settings(self):
        default_settings = {
            "auto_attach": True,
            "refresh_interval": 5,
            "vbox_path": r"C:\Program Files\Oracle\VirtualBox\VirtualBox.exe",
            "color_scheme": "System Default"
        }
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    for key, value in default_settings.items():
                        if key not in settings:
                            settings[key] = value
                    return settings
        except Exception as e:
            print(f"Error loading settings: {e}")
        return default_settings

    def save_settings(self):
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")
            self.show_message_dialog("Error", f"Failed to save settings: {str(e)}", Gtk.MessageType.ERROR)

    def apply_color_scheme(self):
        scheme_name = self.settings.get("color_scheme", "System Default")
        style_manager = Adw.StyleManager.get_default()
        if scheme_name == "Prefer Light":
            style_manager.set_color_scheme(Adw.ColorScheme.FORCE_LIGHT)
        elif scheme_name == "Prefer Dark":
            style_manager.set_color_scheme(Adw.ColorScheme.FORCE_DARK)
        else: 
            style_manager.set_color_scheme(Adw.ColorScheme.DEFAULT)

    def show_settings_dialog_action_cb(self): 
        dialog = SettingsDialog(self, self.settings)
        
        def on_dialog_hidden(d): 
            print("SettingsDialog hidden, fetching settings.")
            new_settings = d.get_settings()
            
            old_refresh_interval = self.settings.get("refresh_interval")
            old_color_scheme = self.settings.get("color_scheme")
            old_auto_attach = self.settings.get("auto_attach")

            self.settings.update(new_settings) 
            self.save_settings()

            if new_settings["refresh_interval"] != old_refresh_interval:
                if self.auto_refresh_timer_id:
                    GLib.source_remove(self.auto_refresh_timer_id)
                refresh_ms = new_settings["refresh_interval"] * 1000
                self.auto_refresh_timer_id = GLib.timeout_add(refresh_ms, self.on_auto_refresh_timer)
            
            if new_settings["color_scheme"] != old_color_scheme:
                self.apply_color_scheme()

            if new_settings["auto_attach"] != old_auto_attach: 
                 self.refresh_tabs_ui()
            
            # Adw.PreferencesDialog might need to be explicitly destroyed if it doesn't auto-destroy
            # when it's hidden and no longer referenced.
            # However, if it's truly non-modal, just hiding might be fine if it's reused.
            # For now, let's assume it can be destroyed after hiding if we don't plan to reuse the instance.
            # d.destroy() # Be cautious with this if it's meant to be reused or auto-destroys.

        dialog.connect("hide", on_dialog_hidden) 
        dialog.present()

    def show_about_dialog_action_cb(self): 
        about_dialog = Adw.AboutWindow(
            transient_for=self,
            modal=True,
            application_name=APP_TITLE,
            application_icon=APP_ID, 
            version="1.3pa1 (GTK4 Port) / Alpha version",
            developer_name="Zalexanninev15", 
            website="https://github.com/Zalexanninev15/VBoxTabs-Manager",
            issue_url="https://github.com/Zalexanninev15/VBoxTabs-Manager/issues",
            copyright="© 2025 Zalexanninev15", 
            license_type=Gtk.License.MIT_X11,
            comments="Combining VirtualBox windows into a single tabbed window.",
            developers=["Zalexanninev15"],
        )
        
        # add_credit_section is deprecated but should still work with a warning.
        # The modern way involves structured lists for `designers`, `artists`, etc.
        # or potentially using Adw.AboutWindowGroup for custom sections if that API evolves.
        about_dialog.add_credit_section("Ported and Enhanced With AI", ["Generative AI Contributor"])
        about_dialog.add_credit_section("Built With Technologies", ["Python", "GTK4", "LibAdwaita", "PyGObject", "PyWin32"])
        
        about_dialog.present()
        
    def on_auto_refresh_timer(self):
        self.refresh_tabs_ui()
        return GLib.SOURCE_CONTINUE 

    def refresh_tabs_ui(self, widget=None, manual_attach_all=False):
        vbox_windows = self.window_finder.find_virtualbox_windows()
        found_hwnds = set()

        for window_data in vbox_windows:
            hwnd = window_data['hwnd']
            found_hwnds.add(hwnd)

            if hwnd not in self.tabs_by_hwnd:
                should_attach_this_one = manual_attach_all or \
                    (self.settings.get("auto_attach", True) and hwnd not in self.manually_detached_windows)

                if should_attach_this_one:
                    print(f"Found new VM: {window_data['title']}. Attaching.")
                    tab_content = VBoxTabGtk(window_data)
                    self.tabs_by_hwnd[hwnd] = tab_content
                    
                    tab_label_box = Gtk.Box(spacing=6, orientation=Gtk.Orientation.HORIZONTAL)
                    tab_label_text = Gtk.Label(label=tab_content.get_vm_title(), xalign=0.0)
                    tab_label_box.append(tab_label_text)

                    page_num = self.notebook.append_page(tab_content, tab_label_box)
                    self.notebook.set_tab_reorderable(tab_content, True)
                    
                    GLib.idle_add(tab_content.attach_window)

                    self.notebook.set_current_page(page_num)
                    if hwnd in self.manually_detached_windows:
                        self.manually_detached_windows.remove(hwnd)
            else:
                tab_content = self.tabs_by_hwnd[hwnd]
                if manual_attach_all and not tab_content.attached:
                    if hwnd in self.manually_detached_windows:
                        self.manually_detached_windows.remove(hwnd)
                    GLib.idle_add(tab_content.attach_window) 
                    page_num = self.get_page_num_for_widget(tab_content)
                    if page_num != -1: self.notebook.set_current_page(page_num)
                
                if tab_content.get_vm_title() != window_data['title'] or \
                   tab_content.original_title != window_data['original_title']:
                    tab_content.set_vm_title(window_data['title'])
                    tab_content.original_title = window_data['original_title']
                    self.update_tab_label_text(tab_content, window_data['title'])


        hwnds_to_remove = set(self.tabs_by_hwnd.keys()) - found_hwnds
        for hwnd in hwnds_to_remove:
            self.remove_tab_by_hwnd(hwnd, detach_first=False)

    def remove_tab_by_hwnd(self, hwnd, detach_first=True):
        if hwnd in self.tabs_by_hwnd:
            tab_content = self.tabs_by_hwnd[hwnd]
            if detach_first and tab_content.attached:
                tab_content.detach_window(manual=False) 

            page_num = self.get_page_num_for_widget(tab_content)
            if page_num != -1:
                self.notebook.remove_page(page_num)
            
            del self.tabs_by_hwnd[hwnd]
            if hwnd in self.manually_detached_windows: 
                self.manually_detached_windows.remove(hwnd)
            print(f"Removed tab for HWND {hwnd} ('{tab_content.get_vm_title()}')")

    def get_page_num_for_widget(self, widget_to_find):
        for i in range(self.notebook.get_n_pages()):
            if self.notebook.get_nth_page(i) == widget_to_find:
                return i
        return -1
        
    def get_current_tab_content(self):
        current_page_num = self.notebook.get_current_page()
        if current_page_num >= 0:
            return self.notebook.get_nth_page(current_page_num)
        return None

    def detach_current_tab_ui(self, widget=None):
        tab_content = self.get_current_tab_content()
        if tab_content and isinstance(tab_content, VBoxTabGtk):
            self.detach_specific_tab(tab_content) 

    def rename_current_tab_ui(self, action=None, param=None):
        tab_content = self.get_current_tab_content()
        if tab_content and isinstance(tab_content, VBoxTabGtk):
            self.rename_specific_tab(tab_content) 

    def update_tab_label_text(self, tab_content_widget, new_text):
        page_num = self.get_page_num_for_widget(tab_content_widget)
        if page_num != -1:
            tab_label_widget = self.notebook.get_tab_label(tab_content_widget)
            if tab_label_widget and isinstance(tab_label_widget, Gtk.Box): 
                child = tab_label_widget.get_first_child()
                while child:
                    if isinstance(child, Gtk.Label):
                        child.set_text(new_text)
                        break
                    child = child.get_next_sibling()

    def on_window_close_request(self, window):
        print("Main window close requested. Detaching all tabs.")
        for hwnd, tab_content in list(self.tabs_by_hwnd.items()):
            if tab_content.attached:
                tab_content.detach_window(manual=False)
        
        if self.auto_refresh_timer_id:
            GLib.source_remove(self.auto_refresh_timer_id)
            self.auto_refresh_timer_id = None
        
        self.save_settings()
        return False 

    def open_virtualbox_main(self, widget=None):
        vbox_path = self.settings.get("vbox_path")
        if vbox_path and os.path.exists(vbox_path):
            try:
                subprocess.Popen([vbox_path])
            except Exception as e:
                self.show_message_dialog("Error", f"Failed to open VirtualBox: {e}", Gtk.MessageType.ERROR)
        else:
            self.show_message_dialog("Error", f"VirtualBox executable not found at:\n{vbox_path or 'Not specified'}\nPlease check Settings.", Gtk.MessageType.WARNING)

    def _terminate_vm_process(self, hwnd):
        if not win32gui.IsWindow(hwnd): 
            return True
        try:
            _, process_id = win32process.GetWindowThreadProcessId(hwnd)
            process_handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, process_id)
            win32api.TerminateProcess(process_handle, 0)
            win32api.CloseHandle(process_handle)
            return True
        except Exception as e:
            print(f"Failed to terminate process for HWND {hwnd}: {e}")
            if not win32gui.IsWindow(hwnd): return True 
            self.show_message_dialog("Error", f"Failed to close VM window (HWND: {hwnd}): {e}", Gtk.MessageType.ERROR)
            return False

    def close_current_vm_window(self, widget=None):
        tab_content = self.get_current_tab_content()
        if tab_content and isinstance(tab_content, VBoxTabGtk):
            self.close_specific_vm_window(tab_content) 

    def close_all_windows_ui(self, action=None, param=None):
        if not self.tabs_by_hwnd:
            self.show_message_dialog("Information", "No VMs are currently tabbed.", Gtk.MessageType.INFO)
            return

        dialog = Gtk.MessageDialog(transient_for=self, modal=True,
                                   message_type=Gtk.MessageType.QUESTION,
                                   buttons=Gtk.ButtonsType.YES_NO,
                                   text="Close All VMs?")
        dialog.format_secondary_text("Are you sure you want to forcefully close all tabbed VM windows?")
        
        def on_response(d, response_id):
            if response_id == Gtk.ResponseType.YES:
                closed_count = 0
                try: subprocess.run(['taskkill', '/F', '/IM', 'VBoxSVC.exe'], check=False, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
                except Exception: pass

                for hwnd, tab_content in list(self.tabs_by_hwnd.items()): 
                    if self._terminate_vm_process(hwnd):
                        closed_count += 1
                
                self.tabs_by_hwnd.clear() 
                self.manually_detached_windows.clear()
                
                while self.notebook.get_n_pages() > 0:
                    self.notebook.remove_page(0)

                self.refresh_tabs_ui() 
                self.show_message_dialog("VMs Closed", f"{closed_count} VM windows forcefully closed.", Gtk.MessageType.INFO)
            d.destroy()

        dialog.connect("response", on_response)
        dialog.show()

    def show_message_dialog(self, title, message, msg_type=Gtk.MessageType.INFO):
        dialog = Gtk.MessageDialog(transient_for=self, modal=True,
                                   message_type=msg_type,
                                   buttons=Gtk.ButtonsType.OK,
                                   text=title)
        dialog.format_secondary_text(message)
        dialog.connect("response", lambda d, r: d.destroy())
        dialog.show()

    def on_notebook_page_added(self, notebook, child, page_num):
        tab_label_widget = notebook.get_tab_label(child)
        if tab_label_widget and isinstance(child, VBoxTabGtk): 
            middle_click_gesture = Gtk.GestureClick.new()
            middle_click_gesture.set_button(Gdk.BUTTON_MIDDLE)
            middle_click_gesture.connect("pressed", self.on_tab_middle_click, child)
            tab_label_widget.add_controller(middle_click_gesture)

            right_click_gesture = Gtk.GestureClick.new()
            right_click_gesture.set_button(Gdk.BUTTON_RIGHT)
            right_click_gesture.connect("pressed", self.on_tab_right_click, child)
            tab_label_widget.add_controller(right_click_gesture)

    def on_tab_middle_click(self, gesture, n_press, x, y, tab_content_widget):
        if isinstance(tab_content_widget, VBoxTabGtk):
            self.close_specific_vm_window(tab_content_widget, from_middle_click=True)

    def on_tab_right_click(self, gesture, n_press, x, y, tab_content_widget):
        if not isinstance(tab_content_widget, VBoxTabGtk): return

        page_num = self.get_page_num_for_widget(tab_content_widget)
        if page_num != -1: self.notebook.set_current_page(page_num)
        
        menu = Gio.Menu()
        menu.append("Rename", "tab.rename")
        menu.append("Detach", "tab.detach")
        menu.append("Close VM Window", "tab.close_vm")

        popover = Gtk.PopoverMenu.new_from_model(menu)
        tab_label_widget = self.notebook.get_tab_label(tab_content_widget)
        popover.set_parent(tab_label_widget if tab_label_widget else tab_content_widget)
        
        action_group = Gio.SimpleActionGroup()
        
        action_rename = Gio.SimpleAction.new("rename", None)
        action_rename.connect("activate", lambda a, p, tc=tab_content_widget: self.rename_specific_tab(tc))
        action_group.add_action(action_rename)

        action_detach = Gio.SimpleAction.new("detach", None)
        action_detach.connect("activate", lambda a, p, tc=tab_content_widget: self.detach_specific_tab(tc))
        action_group.add_action(action_detach)

        action_close_vm = Gio.SimpleAction.new("close_vm", None)
        action_close_vm.connect("activate", lambda a, p, tc=tab_content_widget: self.close_specific_vm_window(tc))
        action_group.add_action(action_close_vm)
        
        if not tab_content_widget.get_action_group("tab"):
            tab_content_widget.insert_action_group("tab", action_group)
        else: 
            existing_group = tab_content_widget.get_action_group("tab")
            if isinstance(existing_group, Gio.SimpleActionGroup):
                for act_name in ["rename", "detach", "close_vm"]: 
                    if existing_group.has_action(act_name): existing_group.remove_action(act_name)
                existing_group.add_action(action_rename)
                existing_group.add_action(action_detach)
                existing_group.add_action(action_close_vm)
            else: 
                 tab_content_widget.insert_action_group("tab", action_group)

        popover.popup()

    def rename_specific_tab(self, tab_content):
        dialog = Gtk.Dialog(title="Rename Tab", transient_for=self, modal=True)
        dialog.add_buttons("_Cancel", Gtk.ResponseType.CANCEL, "_Rename", Gtk.ResponseType.OK)
        
        content_area = dialog.get_content_area() 
        entry_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        entry_box.set_margin_top(12)
        entry_box.set_margin_bottom(12)
        entry_box.set_margin_start(12)
        entry_box.set_margin_end(12)
        
        entry_box.append(Gtk.Label(label="Enter new name:", xalign=0.0)) 
        
        entry = Gtk.Entry(text=tab_content.get_vm_title())
        entry.set_activates_default(True) 
        entry_box.append(entry)
        
        content_area.append(entry_box)
        dialog.set_default_response(Gtk.ResponseType.OK) 
        
        def on_response(d, response_id):
            if response_id == Gtk.ResponseType.OK:
                new_name = entry.get_text()
                if new_name:
                    tab_content.set_vm_title(new_name)
                    self.update_tab_label_text(tab_content, new_name)
            d.destroy()

        dialog.connect("response", on_response)
        dialog.show() 

    def detach_specific_tab(self, tab_content):
        hwnd = tab_content.hwnd
        tab_content.detach_window(manual=True)
        self.manually_detached_windows.add(hwnd)
        
        self.remove_tab_by_hwnd(hwnd, detach_first=False)


    def close_specific_vm_window(self, tab_content, from_middle_click=False):
        hwnd = tab_content.hwnd
        vm_title = tab_content.get_vm_title()
        if self._terminate_vm_process(hwnd):
            self.remove_tab_by_hwnd(hwnd, detach_first=False)
            msg_suffix = " by middle-click" if from_middle_click else " via context menu"
            self.show_message_dialog("VM Closed", f"VM window for '{vm_title}' forcefully closed{msg_suffix}.", Gtk.MessageType.INFO)


class VBoxTabsManagerApp(Adw.Application):
    def __init__(self, **kwargs):
        super().__init__(application_id=APP_ID, flags=Gio.ApplicationFlags.FLAGS_NONE, **kwargs)
        self.win = None

    def do_activate(self):
        if not self.win:
            self.win = VBoxTabsWindow(application=self)
        self.win.present()

    def do_startup(self):
        Adw.Application.do_startup(self)
        apply_global_font_style(font_family="Segoe UI", font_size="9pt") # MODIFIED


if __name__ == "__main__":
    if os.name == 'nt':
        msys_prefix_env = os.environ.get('MSYSTEM_PREFIX') 
        msys_prefix_path = None

        if msys_prefix_env: 
            msys_prefix_path = msys_prefix_env
        elif 'msys64' in sys.executable.lower(): 
            if 'ucrt64' in sys.executable.lower(): msys_prefix_path = '/ucrt64'
            elif 'mingw64' in sys.executable.lower(): msys_prefix_path = '/mingw64'
            elif 'clang64' in sys.executable.lower(): msys_prefix_path = '/clang64'
        
        if msys_prefix_path:
            if msys_prefix_path.startswith(('C:', 'D:', 'E:', 'F:')): 
                parts = msys_prefix_path.replace('\\', '/').split('/')
                if len(parts) >= 2 and parts[-2].lower() == 'msys64':
                    msys_prefix_path = '/' + parts[-1] 

            data_dirs_str = os.environ.get('XDG_DATA_DIRS', '')
            data_dirs = [d.replace('\\', '/') for d in data_dirs_str.split(os.pathsep) if d]
            
            msys_data_dir_candidate = f"{msys_prefix_path}/share" 

            current_xdg_data_dirs = os.environ.get('XDG_DATA_DIRS', '').replace('\\', '/')
            if msys_data_dir_candidate not in current_xdg_data_dirs:
                new_data_dirs_list = [msys_data_dir_candidate]
                
                base_msys_install_path = None
                if sys.platform == "win32":
                    try:
                        py_path_parts = sys.executable.replace('\\', '/').lower().split('/')
                        if 'msys64' in py_path_parts:
                            msys64_index = py_path_parts.index('msys64')
                            base_msys_install_path = "/".join(py_path_parts[:msys64_index+1]) 
                    except ValueError:
                        pass 
                
                potential_msys_share_paths = []
                if base_msys_install_path:
                    for subdir in ["ucrt64", "mingw64", "clang64", "usr"]:
                        potential_msys_share_paths.append(f"{base_msys_install_path}/{subdir}/share")
                    potential_msys_share_paths.append(f"{base_msys_install_path}/usr/local/share")


                for p in potential_msys_share_paths:
                    p_normalized = p.replace('\\', '/')
                    if p_normalized not in new_data_dirs_list and p_normalized not in data_dirs:
                         new_data_dirs_list.append(p_normalized)
                
                new_data_dirs_list.extend(data_dirs) 
                
                os.environ['XDG_DATA_DIRS'] = os.pathsep.join(list(dict.fromkeys(filter(None, new_data_dirs_list))))
                print(f"Updated XDG_DATA_DIRS: {os.environ['XDG_DATA_DIRS']}")
            else:
                print(f"XDG_DATA_DIRS ('{current_xdg_data_dirs}') appears to include {msys_data_dir_candidate}.")
        else:
            print("MSYS_PREFIX not found or derived, XDG_DATA_DIRS not modified by script.")

    app = VBoxTabsManagerApp()
    exit_status = app.run(sys.argv)
    sys.exit(exit_status)