"""Profile management for monitor configurations."""
import json
import os
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from monitor_api import MonitorInfo, ApplyResult, get_monitors, apply_monitor_settings
import window_manager


def get_config_dir() -> Path:
    """Get the configuration directory."""
    appdata = os.environ.get("APPDATA", os.path.expanduser("~"))
    config_dir = Path(appdata) / "DisplaySnap"
    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir


def get_profiles_path() -> Path:
    """Get the profiles file path."""
    return get_config_dir() / "profiles.json"


@dataclass
class Profile:
    """A saved monitor configuration profile."""
    name: str
    monitors: List[MonitorInfo]
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())
    updated_at: str = field(default_factory=lambda: datetime.now().isoformat())

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "monitors": [m.to_dict() for m in self.monitors],
            "created_at": self.created_at,
            "updated_at": self.updated_at,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "Profile":
        return cls(
            name=data["name"],
            monitors=[MonitorInfo.from_dict(m) for m in data["monitors"]],
            created_at=data.get("created_at", datetime.now().isoformat()),
            updated_at=data.get("updated_at", datetime.now().isoformat()),
        )


class ProfileManager:
    """Manages monitor configuration profiles."""

    def __init__(self):
        self.profiles: Dict[str, Profile] = {}
        self.load_profiles()

    def load_profiles(self) -> None:
        """Load profiles from disk."""
        path = get_profiles_path()
        if path.exists():
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.profiles = {
                        name: Profile.from_dict(p)
                        for name, p in data.get("profiles", {}).items()
                    }
            except (json.JSONDecodeError, KeyError):
                self.profiles = {}

    def save_profiles(self) -> None:
        """Save profiles to disk."""
        path = get_profiles_path()
        data = {
            "profiles": {name: p.to_dict() for name, p in self.profiles.items()}
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def get_profile_names(self) -> List[str]:
        """Get list of profile names."""
        return list(self.profiles.keys())

    def get_profile(self, name: str) -> Optional[Profile]:
        """Get a profile by name."""
        return self.profiles.get(name)

    def save_current_as(self, name: str) -> Profile:
        """Save current monitor configuration as a profile (all enabled)."""
        monitors = get_monitors()
        return self.save_current_as_with_states(name, monitors)

    def save_current_as_with_states(self, name: str, monitors: List[MonitorInfo]) -> Profile:
        """Save monitor configuration with custom enabled states."""
        now = datetime.now().isoformat()

        if name in self.profiles:
            # Update existing
            profile = self.profiles[name]
            profile.monitors = monitors
            profile.updated_at = now
        else:
            # Create new
            profile = Profile(name=name, monitors=monitors, created_at=now, updated_at=now)
            self.profiles[name] = profile

        self.save_profiles()
        return profile

    def delete_profile(self, name: str) -> bool:
        """Delete a profile."""
        if name in self.profiles:
            del self.profiles[name]
            self.save_profiles()
            return True
        return False

    def apply_profile(self, name: str, disable_extra: bool = False,
                       manage_windows: bool = True) -> ApplyResult:
        """Apply a saved profile. Returns detailed result.

        Args:
            name: Profile name
            disable_extra: If True, disable monitors not in the profile
            manage_windows: If True, save/restore window positions automatically
        """
        profile = self.profiles.get(name)
        if not profile:
            return ApplyResult(success=False, failed=["Profile not found"])

        # Get current monitors before applying
        current_monitors = get_monitors()
        current_monitor_positions = {(m.position_x, m.position_y) for m in current_monitors}

        # Get monitors that will be enabled in the new profile
        profile_monitor_positions = {
            (m.position_x, m.position_y) for m in profile.monitors if m.enabled
        }

        # Determine which monitors will be disabled
        monitors_to_disable = current_monitor_positions - profile_monitor_positions

        if manage_windows and monitors_to_disable:
            # Only save positions when REDUCING monitors (not when expanding)
            # This preserves the original multi-monitor layout for restoration
            try:
                # 1. Save current window positions (for restoration when expanding later)
                saved_positions = window_manager.get_window_positions()
                window_manager.save_positions_cache(saved_positions)

                # 2. Move windows from monitors that will be disabled to primary
                for mon_x, mon_y in monitors_to_disable:
                    window_manager.move_windows_from_monitor(mon_x, mon_y)
            except Exception:
                pass  # Don't fail profile application if window management fails

        # 3. Apply monitor settings
        result = apply_monitor_settings(profile.monitors, disable_extra=disable_extra)

        if manage_windows and result.success:
            try:
                # 4. Check if we enabled any new monitors and restore windows
                new_monitors = get_monitors()
                new_monitor_positions = {(m.position_x, m.position_y) for m in new_monitors}

                # Monitors that were just enabled
                newly_enabled = new_monitor_positions - current_monitor_positions

                if newly_enabled:
                    # Load cached positions and restore windows that belong to newly enabled monitors
                    cached = window_manager.load_positions_cache()
                    for pos in cached:
                        if (pos.monitor_x, pos.monitor_y) in newly_enabled:
                            window_manager.restore_window_position(pos)
            except Exception:
                pass  # Don't fail if window restoration fails

        return result

    def rename_profile(self, old_name: str, new_name: str) -> bool:
        """Rename a profile."""
        if old_name not in self.profiles or new_name in self.profiles:
            return False

        # Preserve order
        new_profiles = {}
        for name, p in self.profiles.items():
            if name == old_name:
                p.name = new_name
                p.updated_at = datetime.now().isoformat()
                new_profiles[new_name] = p
            else:
                new_profiles[name] = p
        self.profiles = new_profiles
        self.save_profiles()
        return True

    def move_profile(self, from_idx: int, to_idx: int) -> bool:
        """Move a profile from one position to another."""
        names = list(self.profiles.keys())
        if from_idx < 0 or from_idx >= len(names) or to_idx < 0 or to_idx >= len(names):
            return False

        # Swap positions
        names[from_idx], names[to_idx] = names[to_idx], names[from_idx]

        # Rebuild dict in new order
        new_profiles = {name: self.profiles[name] for name in names}
        self.profiles = new_profiles
        self.save_profiles()
        return True

    def export_profiles(self, file_path: str) -> bool:
        """Export all profiles to a JSON file."""
        try:
            data = {
                "profiles": {name: p.to_dict() for name, p in self.profiles.items()}
            }
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            return True
        except Exception:
            return False

    def import_profiles(self, file_path: str) -> bool:
        """Import profiles from a JSON file (merges with existing)."""
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                for name, p in data.get("profiles", {}).items():
                    self.profiles[name] = Profile.from_dict(p)
            self.save_profiles()
            return True
        except Exception:
            return False
