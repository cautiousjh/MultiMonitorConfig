"""System tray application with settings window."""
import os
import sys
import threading
from tkinter import filedialog
from typing import Optional, List, Dict

import pystray
from PIL import Image, ImageDraw
import customtkinter as ctk

from monitor_api import get_monitors, MonitorInfo
from profile_manager import ProfileManager

# Set appearance
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Fix customtkinter threading bug (RuntimeError: dictionary changed size during iteration)
# Patch the ScalingTracker to handle concurrent modifications safely
try:
    from customtkinter.windows.widgets.scaling import ScalingTracker
    _original_check_dpi = ScalingTracker.check_dpi_scaling.__func__

    @classmethod
    def _safe_check_dpi_scaling(cls, *args, **kwargs):
        try:
            return _original_check_dpi(cls, *args, **kwargs)
        except RuntimeError:
            pass  # Ignore "dictionary changed size during iteration"

    ScalingTracker.check_dpi_scaling = _safe_check_dpi_scaling
except Exception:
    pass


def create_icon_image(size=64):
    """Create a simple monitor icon with white outline for system tray."""
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    # Colors
    outline_color = (255, 255, 255)  # White outline
    fill_color = (100, 180, 255)     # Light blue fill

    margin = size // 8
    # Monitor frame - white outline
    draw.rectangle(
        [margin, margin, size - margin, size - margin - size // 6],
        outline=outline_color,
        width=3
    )
    # Monitor frame - inner fill
    draw.rectangle(
        [margin + 3, margin + 3, size - margin - 3, size - margin - size // 6 - 3],
        fill=(40, 40, 40)
    )
    # Stand
    stand_width = size // 4
    stand_x = (size - stand_width) // 2
    draw.rectangle(
        [stand_x, size - margin - size // 6, stand_x + stand_width, size - margin],
        fill=outline_color
    )
    # Screen content (small squares)
    inner_margin = margin + 8
    sq_size = (size - 2 * inner_margin) // 3
    for i in range(2):
        for j in range(2):
            x = inner_margin + i * (sq_size + 2)
            y = inner_margin + j * (sq_size + 2)
            draw.rectangle([x, y, x + sq_size - 2, y + sq_size - 2], fill=fill_color)

    return image


class SettingsWindow:
    """Settings window for profile management."""

    def __init__(self, profile_manager: ProfileManager, on_close=None):
        self.profile_manager = profile_manager
        self.on_close = on_close
        self.window: Optional[ctk.CTk] = None
        self.monitor_vars: Dict[str, ctk.BooleanVar] = {}
        self.disable_extra_var: Optional[ctk.BooleanVar] = None

    def show(self):
        """Show the settings window."""
        if self.window is not None:
            try:
                self.window.lift()
                self.window.focus_force()
                return
            except:
                self.window = None

        self.window = ctk.CTk()
        self.window.title("DisplaySnap")
        self.window.geometry("620x600")
        self.window.minsize(500, 500)
        self.window.resizable(True, True)
        self.window.attributes('-topmost', True)  # Always on top

        # Set window icon (for taskbar) - delayed to ensure window is ready
        def set_icon():
            try:
                script_dir = os.path.dirname(os.path.abspath(__file__))
                ico_path = os.path.join(script_dir, "icon.ico")
                if os.path.exists(ico_path):
                    self.window.iconbitmap(default=ico_path)
                    self.window.iconbitmap(ico_path)
            except Exception:
                pass
        self.window.after(100, set_icon)

        # Main frame
        main_frame = ctk.CTkFrame(self.window)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)

        # === BOTTOM SECTION (pack first with side="bottom" to ensure always visible) ===

        # Buttons - pack at very bottom
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(side="bottom", fill="x")

        # Apply options - above buttons
        options_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        options_frame.pack(side="bottom", fill="x", pady=(0, 10))

        self.disable_extra_var = ctk.BooleanVar(value=True)  # Default ON for single monitor setups
        disable_extra_cb = ctk.CTkCheckBox(
            options_frame,
            text="Disable monitors not in profile when applying",
            variable=self.disable_extra_var
        )
        disable_extra_cb.pack(anchor="w")

        # Profile detail - above options
        self.detail_text = ctk.CTkTextbox(main_frame, height=80)
        self.detail_text.pack(side="bottom", fill="x", pady=(0, 10))
        self.detail_text.configure(state="disabled")

        detail_label = ctk.CTkLabel(main_frame, text="Profile Detail", font=ctk.CTkFont(size=14, weight="bold"))
        detail_label.pack(side="bottom", anchor="w", pady=(10, 5))

        # === TOP SECTION (pack normally from top) ===

        # Current monitors section
        monitors_label = ctk.CTkLabel(main_frame, text="Current Monitors", font=ctk.CTkFont(size=14, weight="bold"))
        monitors_label.pack(anchor="w", pady=(0, 0))

        monitors_hint = ctk.CTkLabel(main_frame, text="(Uncheck to disable when saving profile)", font=ctk.CTkFont(size=11), text_color="gray")
        monitors_hint.pack(anchor="w", pady=(0, 5))

        self.monitors_frame = ctk.CTkFrame(main_frame)
        self.monitors_frame.pack(fill="x", pady=(0, 15))

        # === MIDDLE SECTION (expandable - fills remaining space) ===

        # Profiles section - this is the only expandable section
        profiles_label = ctk.CTkLabel(main_frame, text="Saved Profiles", font=ctk.CTkFont(size=14, weight="bold"))
        profiles_label.pack(anchor="w", pady=(0, 5))

        # Profile list frame - expands to fill available space
        list_frame = ctk.CTkFrame(main_frame)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))

        self.profile_listbox = ctk.CTkScrollableFrame(list_frame, height=100)
        self.profile_listbox.pack(fill="both", expand=True, padx=5, pady=5)
        self.profile_buttons: List[ctk.CTkButton] = []

        ctk.CTkButton(btn_frame, text="Save", command=self._save_profile, width=55).pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="Apply", command=self._apply_selected, width=55).pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="Rename", command=self._rename_profile, width=65).pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="Delete", command=self._delete_profile, width=55, fg_color="#c42b1c", hover_color="#a32419").pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="▲", command=self._move_up, width=28).pack(side="left", padx=1)
        ctk.CTkButton(btn_frame, text="▼", command=self._move_down, width=28).pack(side="left", padx=1)
        ctk.CTkButton(btn_frame, text="↻", command=self._refresh, width=28).pack(side="left", padx=1)
        # Export/Import - subtle style
        ctk.CTkButton(btn_frame, text="Import", command=self._import_profiles, width=55, fg_color="gray30", hover_color="gray40").pack(side="right", padx=2)
        ctk.CTkButton(btn_frame, text="Export", command=self._export_profiles, width=55, fg_color="gray30", hover_color="gray40").pack(side="right", padx=2)

        self.selected_profile: Optional[str] = None
        self._refresh()

        self.window.protocol("WM_DELETE_WINDOW", self._on_close)
        try:
            self.window.mainloop()
        except Exception as e:
            print(f"Settings window error: {e}")
            self._on_close()

    def _refresh(self):
        """Refresh the display."""
        # Clear and rebuild monitor checkboxes
        for widget in self.monitors_frame.winfo_children():
            widget.destroy()
        self.monitor_vars.clear()

        monitors = get_monitors()
        for m in monitors:
            var = ctk.BooleanVar(value=True)
            self.monitor_vars[m.device_name] = var

            frame = ctk.CTkFrame(self.monitors_frame, fg_color="transparent")
            frame.pack(fill="x", pady=2, padx=5)

            cb = ctk.CTkCheckBox(frame, text="", variable=var, width=20)
            cb.pack(side="left")

            primary = " [Primary]" if m.is_primary else ""
            text = f"{m.device_name}{primary}: {m.width}x{m.height} @ {m.refresh_rate}Hz"
            label = ctk.CTkLabel(frame, text=text, anchor="w")
            label.pack(side="left", padx=5)

        # Update profile list
        for btn in self.profile_buttons:
            btn.destroy()
        self.profile_buttons.clear()

        for name in self.profile_manager.get_profile_names():
            btn = ctk.CTkButton(
                self.profile_listbox,
                text=name,
                anchor="w",
                fg_color="transparent" if name != self.selected_profile else None,
                command=lambda n=name: self._select_profile(n)
            )
            btn.pack(fill="x", pady=2)
            btn.bind("<Double-Button-1>", lambda e, n=name: self._apply_profile(n))
            self.profile_buttons.append(btn)

    def _select_profile(self, name: str):
        """Select a profile."""
        self.selected_profile = name

        # Update button styles
        for btn in self.profile_buttons:
            if btn.cget("text") == name:
                btn.configure(fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"])
            else:
                btn.configure(fg_color="transparent")

        # Show detail
        profile = self.profile_manager.get_profile(name)
        if profile:
            self.detail_text.configure(state="normal")
            self.detail_text.delete("1.0", "end")
            for m in profile.monitors:
                status = "DISABLED" if not m.enabled else "enabled"
                self.detail_text.insert("end", f"[{status}] {m.device_name}: {m.width}x{m.height} @ {m.refresh_rate}Hz\n")
            self.detail_text.configure(state="disabled")

    def _save_profile(self):
        """Save current configuration as a new profile."""
        dialog = ctk.CTkInputDialog(text="Enter profile name:", title="Save Profile")
        name = dialog.get_input()
        if not name:
            return
        name = name.strip()
        if not name:
            return

        monitors = get_monitors()
        for m in monitors:
            if m.device_name in self.monitor_vars:
                m.enabled = self.monitor_vars[m.device_name].get()

        self.profile_manager.save_current_as_with_states(name, monitors)
        self.selected_profile = name
        self._refresh()

    def _apply_selected(self):
        """Apply selected profile."""
        if not self.selected_profile:
            return
        self._apply_profile(self.selected_profile)

    def _apply_profile(self, name: str):
        """Apply a profile."""
        disable_extra = self.disable_extra_var.get() if self.disable_extra_var else False
        result = self.profile_manager.apply_profile(name, disable_extra=disable_extra)
        if result.success:
            self._refresh()

    def _rename_profile(self):
        """Rename selected profile."""
        if not self.selected_profile:
            return

        dialog = ctk.CTkInputDialog(text=f"New name for '{self.selected_profile}':", title="Rename Profile")
        new_name = dialog.get_input()
        if new_name:
            new_name = new_name.strip()
            if new_name and new_name != self.selected_profile:
                if self.profile_manager.rename_profile(self.selected_profile, new_name):
                    self.selected_profile = new_name
                    self._refresh()

    def _move_up(self):
        """Move selected profile up."""
        if not self.selected_profile:
            return
        names = self.profile_manager.get_profile_names()
        try:
            idx = names.index(self.selected_profile)
            if idx > 0:
                self.profile_manager.move_profile(idx, idx - 1)
                self._refresh()
        except ValueError:
            pass

    def _move_down(self):
        """Move selected profile down."""
        if not self.selected_profile:
            return
        names = self.profile_manager.get_profile_names()
        try:
            idx = names.index(self.selected_profile)
            if idx < len(names) - 1:
                self.profile_manager.move_profile(idx, idx + 1)
                self._refresh()
        except ValueError:
            pass

    def _delete_profile(self):
        """Delete selected profile."""
        if not self.selected_profile:
            return

        self.profile_manager.delete_profile(self.selected_profile)
        self.selected_profile = None
        self._refresh()

    def _export_profiles(self):
        """Export all profiles to a JSON file."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Export Profiles"
        )
        if file_path:
            if self.profile_manager.export_profiles(file_path):
                pass  # Success, no message needed

    def _import_profiles(self):
        """Import profiles from a JSON file."""
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Import Profiles"
        )
        if file_path:
            if self.profile_manager.import_profiles(file_path):
                self._refresh()

    def _on_close(self):
        """Handle window close."""
        if self.window:
            self.window.destroy()
            self.window = None
        if self.on_close:
            self.on_close()


class TrayApp:
    """System tray application."""

    def __init__(self):
        self.profile_manager = ProfileManager()
        self.settings_window: Optional[SettingsWindow] = None
        self.icon: Optional[pystray.Icon] = None

    def _build_menu(self):
        """Build the tray menu."""
        items = []

        items.append(pystray.MenuItem("Open Settings...", self._show_settings, default=True))
        items.append(pystray.Menu.SEPARATOR)

        profile_names = self.profile_manager.get_profile_names()
        if profile_names:
            profile_items = [
                pystray.MenuItem(name, lambda _, n=name: self._apply_profile(n))
                for name in profile_names
            ]
            items.append(pystray.MenuItem("Apply Profile", pystray.Menu(*profile_items)))
        else:
            items.append(pystray.MenuItem("No saved profiles", None, enabled=False))

        items.append(pystray.Menu.SEPARATOR)
        items.append(pystray.MenuItem("Save Current Config...", self._quick_save))
        items.append(pystray.Menu.SEPARATOR)
        items.append(pystray.MenuItem("Exit", self._exit))

        return pystray.Menu(*items)

    def _show_settings(self):
        """Show the settings window."""
        # Check if window already exists and is open
        if self.settings_window and self.settings_window.window:
            try:
                self.settings_window.window.lift()
                self.settings_window.window.focus_force()
                return
            except:
                pass  # Window was destroyed, create new one

        def run_settings():
            self.settings_window = SettingsWindow(
                self.profile_manager,
                on_close=lambda: self._update_menu()
            )
            self.settings_window.show()

        threading.Thread(target=run_settings, daemon=True).start()

    def _apply_profile(self, name: str):
        """Apply a profile from the tray menu."""
        result = self.profile_manager.apply_profile(name)
        if result.success:
            msg = f"Applied: {name}"
            if result.skipped:
                msg += f" (skipped: {len(result.skipped)})"
            if result.disabled:
                msg += f" (disabled: {len(result.disabled)})"
            self.icon.notify(msg, "DisplaySnap")
        else:
            self.icon.notify(f"Failed to apply: {name}", "DisplaySnap")

    def _quick_save(self):
        """Quick save current configuration."""
        def save():
            try:
                app = ctk.CTk()
                app.withdraw()
                dialog = ctk.CTkInputDialog(text="Enter profile name:", title="Save Profile")
                name = dialog.get_input()
                app.destroy()

                if name:
                    name = name.strip()
                    if name:
                        self.profile_manager.save_current_as(name)
                        self._update_menu()
                        self.icon.notify(f"Saved: {name}", "DisplaySnap")
            except Exception as e:
                print(f"Quick save error: {e}")

        threading.Thread(target=save, daemon=True).start()

    def _update_menu(self):
        """Update the tray menu."""
        if self.icon:
            self.icon.menu = self._build_menu()

    def _exit(self):
        """Exit the application."""
        if self.icon:
            self.icon.stop()

    def run(self):
        """Run the tray application."""
        while True:
            try:
                image = create_icon_image()
                self.icon = pystray.Icon(
                    "DisplaySnap",
                    image,
                    "DisplaySnap",
                    menu=self._build_menu()
                )

                # Open settings window after a short delay (gives tray icon time to initialize)
                def delayed_open_settings():
                    import time
                    time.sleep(0.5)  # Small delay to ensure tray icon is ready
                    self._show_settings()

                threading.Thread(target=delayed_open_settings, daemon=True).start()

                # Run the icon (blocking call - keeps app alive)
                self.icon.run()
                break
            except Exception as e:
                print(f"Tray app error: {e}, restarting...")
                import time
                time.sleep(1)


def main():
    try:
        app = TrayApp()
        app.run()
    except Exception as e:
        print(f"Fatal error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
