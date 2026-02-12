"""Windows Monitor API wrapper using pywin32."""
import ctypes
from ctypes import wintypes
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

# Windows API constants
ENUM_CURRENT_SETTINGS = -1
ENUM_REGISTRY_SETTINGS = -2
CDS_UPDATEREGISTRY = 0x00000001
CDS_TEST = 0x00000002
CDS_SET_PRIMARY = 0x00000010
CDS_NORESET = 0x10000000
CDS_RESET = 0x40000000

# SetDisplayConfig flags
SDC_APPLY = 0x00000080
SDC_TOPOLOGY_EXTEND = 0x00000004

DISP_CHANGE_SUCCESSFUL = 0
DISP_CHANGE_RESTART = 1
DISP_CHANGE_FAILED = -1
DISP_CHANGE_BADMODE = -2

DISPLAY_DEVICE_ATTACHED_TO_DESKTOP = 0x00000001
DISPLAY_DEVICE_PRIMARY_DEVICE = 0x00000004
DISPLAY_DEVICE_ACTIVE = 0x00000001

DM_PELSWIDTH = 0x00080000
DM_PELSHEIGHT = 0x00100000
DM_BITSPERPEL = 0x00040000
DM_DISPLAYFREQUENCY = 0x00400000
DM_POSITION = 0x00000020
DM_DISPLAYORIENTATION = 0x00000080


class DEVMODE(ctypes.Structure):
    _fields_ = [
        ("dmDeviceName", wintypes.WCHAR * 32),
        ("dmSpecVersion", wintypes.WORD),
        ("dmDriverVersion", wintypes.WORD),
        ("dmSize", wintypes.WORD),
        ("dmDriverExtra", wintypes.WORD),
        ("dmFields", wintypes.DWORD),
        ("dmPositionX", wintypes.LONG),
        ("dmPositionY", wintypes.LONG),
        ("dmDisplayOrientation", wintypes.DWORD),
        ("dmDisplayFixedOutput", wintypes.DWORD),
        ("dmColor", wintypes.SHORT),
        ("dmDuplex", wintypes.SHORT),
        ("dmYResolution", wintypes.SHORT),
        ("dmTTOption", wintypes.SHORT),
        ("dmCollate", wintypes.SHORT),
        ("dmFormName", wintypes.WCHAR * 32),
        ("dmLogPixels", wintypes.WORD),
        ("dmBitsPerPel", wintypes.DWORD),
        ("dmPelsWidth", wintypes.DWORD),
        ("dmPelsHeight", wintypes.DWORD),
        ("dmDisplayFlags", wintypes.DWORD),
        ("dmDisplayFrequency", wintypes.DWORD),
        ("dmICMMethod", wintypes.DWORD),
        ("dmICMIntent", wintypes.DWORD),
        ("dmMediaType", wintypes.DWORD),
        ("dmDitherType", wintypes.DWORD),
        ("dmReserved1", wintypes.DWORD),
        ("dmReserved2", wintypes.DWORD),
        ("dmPanningWidth", wintypes.DWORD),
        ("dmPanningHeight", wintypes.DWORD),
    ]


class DISPLAY_DEVICE(ctypes.Structure):
    _fields_ = [
        ("cb", wintypes.DWORD),
        ("DeviceName", wintypes.WCHAR * 32),
        ("DeviceString", wintypes.WCHAR * 128),
        ("StateFlags", wintypes.DWORD),
        ("DeviceID", wintypes.WCHAR * 128),
        ("DeviceKey", wintypes.WCHAR * 128),
    ]


user32 = ctypes.windll.user32


def detect_displays() -> bool:
    """Detect and attach any physically connected displays.

    This triggers Windows to scan for connected monitors and extend
    the desktop to them. Useful when a monitor was disabled and needs
    to be re-enabled.

    Returns:
        True if successful, False otherwise
    """
    result = user32.SetDisplayConfig(0, None, 0, None, SDC_APPLY | SDC_TOPOLOGY_EXTEND)
    return result == 0


@dataclass
class MonitorInfo:
    """Monitor information."""
    device_name: str
    device_string: str
    width: int
    height: int
    position_x: int
    position_y: int
    refresh_rate: int
    orientation: int
    bits_per_pixel: int
    is_primary: bool
    is_active: bool
    enabled: bool = True  # False = disable this monitor

    def to_dict(self) -> dict:
        return {
            "device_name": self.device_name,
            "device_string": self.device_string,
            "width": self.width,
            "height": self.height,
            "position_x": self.position_x,
            "position_y": self.position_y,
            "refresh_rate": self.refresh_rate,
            "orientation": self.orientation,
            "bits_per_pixel": self.bits_per_pixel,
            "is_primary": self.is_primary,
            "is_active": self.is_active,
            "enabled": self.enabled,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "MonitorInfo":
        # Handle old profiles without 'enabled' field
        if "enabled" not in data:
            data = {**data, "enabled": True}
        return cls(**data)

    def __str__(self) -> str:
        primary = " [Primary]" if self.is_primary else ""
        disabled = " [DISABLED]" if not self.enabled else ""
        return f"{self.device_name}{primary}{disabled}: {self.width}x{self.height} @ {self.refresh_rate}Hz, pos({self.position_x}, {self.position_y})"


@dataclass
class ApplyResult:
    """Result of applying monitor settings."""
    success: bool
    applied: List[str] = field(default_factory=list)      # Successfully applied
    skipped: List[str] = field(default_factory=list)      # Not found (disconnected)
    failed: List[str] = field(default_factory=list)       # Failed to apply
    disabled: List[str] = field(default_factory=list)     # Disabled monitors

    def get_message(self) -> str:
        parts = []
        if self.applied:
            parts.append(f"Applied: {', '.join(self.applied)}")
        if self.skipped:
            parts.append(f"Skipped (not connected): {', '.join(self.skipped)}")
        if self.disabled:
            parts.append(f"Disabled: {', '.join(self.disabled)}")
        if self.failed:
            parts.append(f"Failed: {', '.join(self.failed)}")
        return "\n".join(parts) if parts else "No changes"


def get_monitors() -> List[MonitorInfo]:
    """Get all connected monitors."""
    monitors = []
    device = DISPLAY_DEVICE()
    device.cb = ctypes.sizeof(device)

    i = 0
    while user32.EnumDisplayDevicesW(None, i, ctypes.byref(device), 0):
        if device.StateFlags & DISPLAY_DEVICE_ATTACHED_TO_DESKTOP:
            devmode = DEVMODE()
            devmode.dmSize = ctypes.sizeof(devmode)

            if user32.EnumDisplaySettingsW(device.DeviceName, ENUM_CURRENT_SETTINGS, ctypes.byref(devmode)):
                monitor = MonitorInfo(
                    device_name=device.DeviceName,
                    device_string=device.DeviceString,
                    width=devmode.dmPelsWidth,
                    height=devmode.dmPelsHeight,
                    position_x=devmode.dmPositionX,
                    position_y=devmode.dmPositionY,
                    refresh_rate=devmode.dmDisplayFrequency,
                    orientation=devmode.dmDisplayOrientation,
                    bits_per_pixel=devmode.dmBitsPerPel,
                    is_primary=bool(device.StateFlags & DISPLAY_DEVICE_PRIMARY_DEVICE),
                    is_active=bool(device.StateFlags & DISPLAY_DEVICE_ACTIVE),
                )
                monitors.append(monitor)
        i += 1

    return monitors


def get_connected_device_names() -> set:
    """Get set of currently connected monitor device names (active only)."""
    names = set()
    device = DISPLAY_DEVICE()
    device.cb = ctypes.sizeof(device)
    i = 0
    while user32.EnumDisplayDevicesW(None, i, ctypes.byref(device), 0):
        if device.StateFlags & DISPLAY_DEVICE_ATTACHED_TO_DESKTOP:
            names.add(device.DeviceName)
        i += 1
    return names


def get_all_device_names() -> set:
    """Get set of ALL display device names (including disabled ones)."""
    names = set()
    device = DISPLAY_DEVICE()
    device.cb = ctypes.sizeof(device)
    i = 0
    while user32.EnumDisplayDevicesW(None, i, ctypes.byref(device), 0):
        # Include all display devices, not just attached ones
        names.add(device.DeviceName)
        i += 1
    return names


def get_best_display_mode(device_name: str, target_width: int, target_height: int,
                          target_refresh: int, is_rotated: bool = False,
                          fallback_to_any: bool = True) -> Optional[DEVMODE]:
    """Find the best matching display mode for a device.

    Args:
        device_name: The device name
        target_width: Desired width (as stored in profile, may be rotated)
        target_height: Desired height (as stored in profile, may be rotated)
        target_refresh: Desired refresh rate
        is_rotated: If True, width/height are swapped (portrait mode)
        fallback_to_any: If True, return highest available mode if exact not found

    Returns:
        DEVMODE with matching mode, or best available if fallback enabled
    """
    best_mode = None
    best_score = -1
    highest_mode = None
    highest_pixels = 0

    # For rotated displays, we need to search for the native (unrotated) resolution
    search_width = target_width
    search_height = target_height
    if is_rotated or target_height > target_width:
        search_width = target_height
        search_height = target_width

    devmode = DEVMODE()
    devmode.dmSize = ctypes.sizeof(devmode)

    i = 0
    while user32.EnumDisplaySettingsW(device_name, i, ctypes.byref(devmode)):
        pixels = devmode.dmPelsWidth * devmode.dmPelsHeight

        # Track highest resolution mode for fallback
        if pixels > highest_pixels or (pixels == highest_pixels and
                                        devmode.dmDisplayFrequency > (highest_mode.dmDisplayFrequency if highest_mode else 0)):
            highest_pixels = pixels
            highest_mode = DEVMODE()
            highest_mode.dmSize = ctypes.sizeof(highest_mode)
            ctypes.memmove(ctypes.byref(highest_mode), ctypes.byref(devmode), ctypes.sizeof(devmode))

        # Check for exact match or rotated match
        exact_match = (devmode.dmPelsWidth == target_width and devmode.dmPelsHeight == target_height)
        native_match = (devmode.dmPelsWidth == search_width and devmode.dmPelsHeight == search_height)

        if exact_match or native_match:
            score = 1000 if exact_match else 900
            if devmode.dmDisplayFrequency == target_refresh:
                score += 100
            else:
                score += max(0, 50 - abs(devmode.dmDisplayFrequency - target_refresh))

            if score > best_score:
                best_score = score
                best_mode = DEVMODE()
                best_mode.dmSize = ctypes.sizeof(best_mode)
                ctypes.memmove(ctypes.byref(best_mode), ctypes.byref(devmode), ctypes.sizeof(devmode))

        i += 1
        devmode = DEVMODE()
        devmode.dmSize = ctypes.sizeof(devmode)

    # Return best exact match, or fallback to highest available
    if best_mode:
        return best_mode
    elif fallback_to_any and highest_mode:
        print(f"  Warning: {target_width}x{target_height} not available, using {highest_mode.dmPelsWidth}x{highest_mode.dmPelsHeight}")
        return highest_mode
    return None


def enable_monitor(device_name: str, monitor: 'MonitorInfo', use_noreset: bool = False) -> bool:
    """Enable a disabled monitor with the specified settings (Extend desktop).

    Args:
        device_name: The device name (e.g., \\\\.\\DISPLAY1)
        monitor: MonitorInfo with desired settings
        use_noreset: If True, don't apply immediately (for batch operations)
    """
    # Check if monitor is rotated (portrait mode - height > width)
    is_rotated = monitor.height > monitor.width or monitor.orientation in (1, 3)

    # First, try to find a valid display mode for this monitor
    best_mode = get_best_display_mode(device_name, monitor.width, monitor.height,
                                       monitor.refresh_rate, is_rotated=is_rotated)

    if best_mode:
        devmode = best_mode
        # If we found a native mode but monitor is rotated, use native dimensions
        if is_rotated and devmode.dmPelsWidth > devmode.dmPelsHeight:
            # Mode is in landscape, we need portrait - keep the native dimensions
            # The orientation flag will handle the rotation
            pass
    else:
        # Fallback: create DEVMODE from scratch with native resolution
        devmode = DEVMODE()
        devmode.dmSize = ctypes.sizeof(devmode)
        if is_rotated:
            # Use native (unrotated) dimensions
            devmode.dmPelsWidth = monitor.height
            devmode.dmPelsHeight = monitor.width
        else:
            devmode.dmPelsWidth = monitor.width
            devmode.dmPelsHeight = monitor.height
        devmode.dmDisplayFrequency = monitor.refresh_rate
        devmode.dmBitsPerPel = monitor.bits_per_pixel

    # Set position (critical for extending desktop)
    devmode.dmPositionX = monitor.position_x
    devmode.dmPositionY = monitor.position_y
    devmode.dmDisplayOrientation = monitor.orientation

    devmode.dmFields = (DM_PELSWIDTH | DM_PELSHEIGHT | DM_POSITION |
                       DM_DISPLAYFREQUENCY | DM_DISPLAYORIENTATION | DM_BITSPERPEL)

    flags = CDS_UPDATEREGISTRY
    if use_noreset:
        flags |= CDS_NORESET
    if monitor.is_primary:
        flags |= CDS_SET_PRIMARY

    result = user32.ChangeDisplaySettingsExW(
        device_name,
        ctypes.byref(devmode),
        None,
        flags,
        None
    )

    if result != DISP_CHANGE_SUCCESSFUL:
        error_msgs = {
            DISP_CHANGE_RESTART: "RESTART required",
            DISP_CHANGE_FAILED: "FAILED",
            DISP_CHANGE_BADMODE: "BADMODE - invalid mode",
        }
        print(f"enable_monitor({device_name}): {error_msgs.get(result, f'error {result}')}")
        print(f"  Attempted: {devmode.dmPelsWidth}x{devmode.dmPelsHeight} @ {devmode.dmDisplayFrequency}Hz")
        print(f"  Position: ({devmode.dmPositionX}, {devmode.dmPositionY})")

    # If not using noreset, apply immediately
    if not use_noreset and result == DISP_CHANGE_SUCCESSFUL:
        user32.ChangeDisplaySettingsExW(None, None, None, 0, None)

    return result == DISP_CHANGE_SUCCESSFUL


def disable_monitor(device_name: str, use_noreset: bool = False) -> bool:
    """Disable a monitor by detaching it from desktop.

    Args:
        device_name: The device name (e.g., \\\\.\\DISPLAY1)
        use_noreset: If True, don't apply immediately (for batch operations)

    WARNING: Cannot disable the primary monitor.
    """
    devmode = DEVMODE()
    devmode.dmSize = ctypes.sizeof(devmode)

    # Get current settings first
    if not user32.EnumDisplaySettingsW(device_name, ENUM_CURRENT_SETTINGS, ctypes.byref(devmode)):
        return False

    # Set position to detach from desktop (key: set width/height to 0)
    devmode.dmPelsWidth = 0
    devmode.dmPelsHeight = 0
    devmode.dmPositionX = 0
    devmode.dmPositionY = 0
    devmode.dmFields = DM_PELSWIDTH | DM_PELSHEIGHT | DM_POSITION

    flags = CDS_UPDATEREGISTRY
    if use_noreset:
        flags |= CDS_NORESET

    result = user32.ChangeDisplaySettingsExW(
        device_name,
        ctypes.byref(devmode),
        None,
        flags,
        None
    )

    # If not using noreset, apply immediately
    if not use_noreset and result == DISP_CHANGE_SUCCESSFUL:
        user32.ChangeDisplaySettingsExW(None, None, None, 0, None)

    return result == DISP_CHANGE_SUCCESSFUL


def apply_monitor_settings(monitors: List[MonitorInfo], disable_extra: bool = False) -> ApplyResult:
    """Apply monitor settings with detailed result reporting.

    Handles cases where:
    - Monitor count differs from saved profile
    - Some monitors are disconnected
    - Some monitors should be disabled
    - Some monitors need to be re-enabled (were disabled)

    This function first calls detect_displays() to find any physically
    connected but disabled monitors before applying settings.

    Args:
        monitors: List of monitor configurations to apply
        disable_extra: If True, disable monitors not in the profile
    """
    result = ApplyResult(success=True)

    # Get current state BEFORE detection (fast - just device enumeration)
    connected = get_connected_device_names()  # Currently active monitors
    all_devices = get_all_device_names()       # All monitors including disabled
    profile_device_names = {m.device_name for m in monitors}

    # Only call detect_displays() when we need to re-enable a disabled monitor.
    # This saves 1-3 seconds on every apply that doesn't need re-enabling.
    needs_reenable = any(
        m.enabled and m.device_name not in connected and m.device_name in all_devices
        for m in monitors
    )
    if needs_reenable:
        detect_displays()
        # Re-query after detection since new monitors may have appeared
        connected = get_connected_device_names()
        all_devices = get_all_device_names()

    # Determine primary from connected set (avoids full get_monitors() call)
    # We only need to know which device is primary to prevent disabling it.
    _primary_devices = set()
    device = DISPLAY_DEVICE()
    device.cb = ctypes.sizeof(device)
    _i = 0
    while user32.EnumDisplayDevicesW(None, _i, ctypes.byref(device), 0):
        if device.StateFlags & DISPLAY_DEVICE_PRIMARY_DEVICE:
            _primary_devices.add(device.DeviceName)
        _i += 1
    primary_devices = _primary_devices

    # First: disable monitors marked as disabled in the profile
    for monitor in monitors:
        if monitor.device_name not in all_devices:
            result.skipped.append(monitor.device_name)
            continue

        if not monitor.enabled:
            if monitor.device_name not in connected:
                # Already disabled
                result.disabled.append(monitor.device_name)
                continue

            # Check if it's the primary monitor - can't disable primary
            if monitor.device_name in primary_devices:
                result.failed.append(f"{monitor.device_name} (primary cannot be disabled)")
                continue

            if disable_monitor(monitor.device_name, use_noreset=True):
                result.disabled.append(monitor.device_name)
            else:
                result.failed.append(monitor.device_name)
                result.success = False

    # Second: disable extra monitors not in the profile (if option enabled)
    if disable_extra:
        for device_name in connected:
            if device_name not in profile_device_names:
                if device_name in primary_devices:
                    # Can't disable primary monitor
                    continue
                if disable_monitor(device_name, use_noreset=True):
                    result.disabled.append(device_name)

    # Third: configure enabled monitors (including re-enabling disabled ones)
    for monitor in monitors:
        if monitor.device_name not in all_devices:
            continue  # Already marked as skipped above

        if not monitor.enabled:
            continue  # Already handled above

        # Check if monitor is currently disabled (not in connected but in all_devices)
        is_currently_disabled = monitor.device_name not in connected

        if is_currently_disabled:
            # Re-enable the monitor with the profile settings
            if enable_monitor(monitor.device_name, monitor, use_noreset=True):
                result.applied.append(f"{monitor.device_name} (re-enabled)")
            else:
                result.failed.append(f"{monitor.device_name} (failed to enable)")
                result.success = False
        else:
            # Normal case: configure active monitor
            devmode = DEVMODE()
            devmode.dmSize = ctypes.sizeof(devmode)

            # Get current settings as base
            if not user32.EnumDisplaySettingsW(monitor.device_name, ENUM_CURRENT_SETTINGS, ctypes.byref(devmode)):
                if monitor.device_name not in result.skipped:
                    result.skipped.append(monitor.device_name)
                continue

            # Set new values
            devmode.dmPelsWidth = monitor.width
            devmode.dmPelsHeight = monitor.height
            devmode.dmPositionX = monitor.position_x
            devmode.dmPositionY = monitor.position_y
            devmode.dmDisplayFrequency = monitor.refresh_rate
            devmode.dmDisplayOrientation = monitor.orientation
            devmode.dmBitsPerPel = monitor.bits_per_pixel

            devmode.dmFields = (DM_PELSWIDTH | DM_PELSHEIGHT | DM_POSITION |
                              DM_DISPLAYFREQUENCY | DM_DISPLAYORIENTATION | DM_BITSPERPEL)

            # Apply with NORESET flag (don't apply yet)
            flags = CDS_UPDATEREGISTRY | CDS_NORESET
            if monitor.is_primary:
                flags |= CDS_SET_PRIMARY

            change_result = user32.ChangeDisplaySettingsExW(
                monitor.device_name,
                ctypes.byref(devmode),
                None,
                flags,
                None
            )

            if change_result == DISP_CHANGE_SUCCESSFUL:
                result.applied.append(monitor.device_name)
            else:
                result.failed.append(monitor.device_name)
                result.success = False

    # Final pass: apply all pending changes
    if result.applied or result.disabled:
        final_result = user32.ChangeDisplaySettingsExW(None, None, None, 0, None)
        if final_result != DISP_CHANGE_SUCCESSFUL:
            result.success = False

    return result


def apply_monitor_settings_simple(monitors: List[MonitorInfo]) -> bool:
    """Simple apply that returns bool for backward compatibility."""
    return apply_monitor_settings(monitors).success


def get_all_display_devices() -> List[dict]:
    """Get all display devices with their status (for debugging)."""
    devices = []
    device = DISPLAY_DEVICE()
    device.cb = ctypes.sizeof(device)

    i = 0
    while user32.EnumDisplayDevicesW(None, i, ctypes.byref(device), 0):
        attached = bool(device.StateFlags & DISPLAY_DEVICE_ATTACHED_TO_DESKTOP)
        primary = bool(device.StateFlags & DISPLAY_DEVICE_PRIMARY_DEVICE)

        # Check if there are valid display modes
        devmode = DEVMODE()
        devmode.dmSize = ctypes.sizeof(devmode)
        has_modes = user32.EnumDisplaySettingsW(device.DeviceName, 0, ctypes.byref(devmode))

        devices.append({
            "name": device.DeviceName,
            "string": device.DeviceString,
            "attached": attached,
            "primary": primary,
            "has_modes": bool(has_modes),
            "state_flags": device.StateFlags,
        })
        i += 1

    return devices


if __name__ == "__main__":
    # Comprehensive test
    print("=" * 60)
    print("ALL DISPLAY DEVICES (including disabled):")
    print("=" * 60)
    for dev in get_all_display_devices():
        status = []
        if dev["attached"]:
            status.append("ATTACHED")
        else:
            status.append("DETACHED")
        if dev["primary"]:
            status.append("PRIMARY")
        if dev["has_modes"]:
            status.append("has_modes")

        print(f"  {dev['name']}: [{', '.join(status)}]")
        print(f"    Device: {dev['string']}")
        print(f"    Flags: 0x{dev['state_flags']:08X}")
        print()

    print("=" * 60)
    print("ACTIVE MONITORS:")
    print("=" * 60)
    for m in get_monitors():
        print(f"  {m}")

    print()
    print("Connected (active):", get_connected_device_names())
    print("All devices:", get_all_device_names())
