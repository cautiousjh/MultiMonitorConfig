"""Window position management for DisplaySnap."""
import ctypes
from ctypes import wintypes
from dataclasses import dataclass
from typing import List, Optional, Dict
import json
import os
from pathlib import Path

# Windows API constants
SW_HIDE = 0
SW_SHOWNORMAL = 1
SW_SHOWMINIMIZED = 2
SW_SHOWMAXIMIZED = 3
SW_SHOW = 5
SW_MINIMIZE = 6
SW_RESTORE = 9

MONITOR_DEFAULTTONULL = 0
MONITOR_DEFAULTTOPRIMARY = 1
MONITOR_DEFAULTTONEAREST = 2

GWL_STYLE = -16
GWL_EXSTYLE = -20
WS_VISIBLE = 0x10000000
WS_MINIMIZE = 0x20000000
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000
WS_EX_NOACTIVATE = 0x08000000

SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOZORDER = 0x0004
SWP_NOACTIVATE = 0x0010
SWP_SHOWWINDOW = 0x0040


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", wintypes.LONG),
        ("top", wintypes.LONG),
        ("right", wintypes.LONG),
        ("bottom", wintypes.LONG),
    ]


class POINT(ctypes.Structure):
    _fields_ = [
        ("x", wintypes.LONG),
        ("y", wintypes.LONG),
    ]


class WINDOWPLACEMENT(ctypes.Structure):
    _fields_ = [
        ("length", wintypes.UINT),
        ("flags", wintypes.UINT),
        ("showCmd", wintypes.UINT),
        ("ptMinPosition", POINT),
        ("ptMaxPosition", POINT),
        ("rcNormalPosition", RECT),
    ]


class MONITORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("rcMonitor", RECT),
        ("rcWork", RECT),
        ("dwFlags", wintypes.DWORD),
    ]


user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
psapi = ctypes.windll.psapi

# Function prototypes
EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


@dataclass
class WindowPosition:
    """Saved window position data."""
    hwnd: int
    title: str
    process_name: str
    x: int
    y: int
    width: int
    height: int
    state: int  # SW_SHOWNORMAL, SW_SHOWMINIMIZED, SW_SHOWMAXIMIZED
    monitor_x: int  # Monitor's left position (to identify which monitor)
    monitor_y: int  # Monitor's top position

    def to_dict(self) -> dict:
        return {
            "hwnd": self.hwnd,
            "title": self.title,
            "process_name": self.process_name,
            "x": self.x,
            "y": self.y,
            "width": self.width,
            "height": self.height,
            "state": self.state,
            "monitor_x": self.monitor_x,
            "monitor_y": self.monitor_y,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "WindowPosition":
        return cls(**data)


def get_cache_path() -> Path:
    """Get the window positions cache file path."""
    appdata = os.environ.get("APPDATA", os.path.expanduser("~"))
    cache_dir = Path(appdata) / "DisplaySnap"
    cache_dir.mkdir(parents=True, exist_ok=True)
    return cache_dir / "window_cache.json"


def get_process_name(hwnd: int) -> str:
    """Get the process name for a window handle."""
    try:
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

        handle = kernel32.OpenProcess(0x0400 | 0x0010, False, pid.value)  # PROCESS_QUERY_INFORMATION | PROCESS_VM_READ
        if handle:
            try:
                buffer = ctypes.create_unicode_buffer(260)
                if psapi.GetModuleBaseNameW(handle, None, buffer, 260):
                    return buffer.value
            finally:
                kernel32.CloseHandle(handle)
    except Exception:
        pass
    return ""


def get_window_title(hwnd: int) -> str:
    """Get the window title."""
    try:
        length = user32.GetWindowTextLengthW(hwnd)
        if length > 0:
            buffer = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buffer, length + 1)
            return buffer.value
    except Exception:
        pass
    return ""


def is_real_window(hwnd: int) -> bool:
    """Check if this is a real user window (not system/hidden)."""
    if not user32.IsWindowVisible(hwnd):
        return False

    # Get window styles
    style = user32.GetWindowLongW(hwnd, GWL_STYLE)
    ex_style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)

    # Skip tool windows (floating toolbars, etc.)
    if ex_style & WS_EX_TOOLWINDOW and not (ex_style & WS_EX_APPWINDOW):
        return False

    # Skip windows with no title (often system windows)
    title = get_window_title(hwnd)
    if not title:
        return False

    # Skip certain system windows
    skip_titles = [
        "Program Manager",
        "Windows Input Experience",
        "Microsoft Text Input Application",
        "Settings",
    ]
    if title in skip_titles:
        return False

    # Skip windows belonging to certain processes
    process = get_process_name(hwnd)
    skip_processes = [
        "TextInputHost.exe",
        "ShellExperienceHost.exe",
        "SearchHost.exe",
        "StartMenuExperienceHost.exe",
    ]
    if process in skip_processes:
        return False

    return True


def get_window_monitor_pos(hwnd: int) -> tuple:
    """Get the monitor position (left, top) for a window."""
    hmonitor = user32.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
    if hmonitor:
        mi = MONITORINFO()
        mi.cbSize = ctypes.sizeof(mi)
        if user32.GetMonitorInfoW(hmonitor, ctypes.byref(mi)):
            return (mi.rcMonitor.left, mi.rcMonitor.top)
    return (0, 0)


def get_primary_monitor_rect() -> tuple:
    """Get the primary monitor's work area (x, y, width, height)."""
    # Get primary monitor handle (NULL point = primary)
    pt = POINT(0, 0)
    hmonitor = user32.MonitorFromPoint(pt, MONITOR_DEFAULTTOPRIMARY)
    if hmonitor:
        mi = MONITORINFO()
        mi.cbSize = ctypes.sizeof(mi)
        if user32.GetMonitorInfoW(hmonitor, ctypes.byref(mi)):
            return (
                mi.rcWork.left,
                mi.rcWork.top,
                mi.rcWork.right - mi.rcWork.left,
                mi.rcWork.bottom - mi.rcWork.top
            )
    return (0, 0, 1920, 1080)  # Fallback


def get_all_windows() -> List[int]:
    """Get all visible user windows."""
    windows = []

    def enum_callback(hwnd, lParam):
        if is_real_window(hwnd):
            windows.append(hwnd)
        return True

    user32.EnumWindows(EnumWindowsProc(enum_callback), 0)
    return windows


def get_window_positions() -> List[WindowPosition]:
    """Get positions of all visible windows."""
    positions = []

    for hwnd in get_all_windows():
        try:
            wp = WINDOWPLACEMENT()
            wp.length = ctypes.sizeof(wp)

            if not user32.GetWindowPlacement(hwnd, ctypes.byref(wp)):
                continue

            rect = wp.rcNormalPosition
            monitor_x, monitor_y = get_window_monitor_pos(hwnd)

            pos = WindowPosition(
                hwnd=hwnd,
                title=get_window_title(hwnd),
                process_name=get_process_name(hwnd),
                x=rect.left,
                y=rect.top,
                width=rect.right - rect.left,
                height=rect.bottom - rect.top,
                state=wp.showCmd,
                monitor_x=monitor_x,
                monitor_y=monitor_y,
            )
            positions.append(pos)
        except Exception:
            continue

    return positions


def save_positions_cache(positions: List[WindowPosition]) -> bool:
    """Save window positions to cache file."""
    try:
        data = {
            "positions": [p.to_dict() for p in positions]
        }
        with open(get_cache_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False


def load_positions_cache() -> List[WindowPosition]:
    """Load window positions from cache file."""
    try:
        path = get_cache_path()
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return [WindowPosition.from_dict(p) for p in data.get("positions", [])]
    except Exception:
        pass
    return []


def is_position_on_monitor(x: int, y: int, monitor_x: int, monitor_y: int) -> bool:
    """Check if a position belongs to a specific monitor (by monitor origin)."""
    # A window is considered "on" a monitor if its top-left corner
    # is within that monitor's bounds. This is a simple heuristic.
    # We just check if the saved monitor origin matches.
    return True  # Will be refined when we have the actual monitor list


def move_window_to_primary(hwnd: int) -> bool:
    """Move a window to the primary monitor."""
    try:
        # Get current window placement
        wp = WINDOWPLACEMENT()
        wp.length = ctypes.sizeof(wp)
        if not user32.GetWindowPlacement(hwnd, ctypes.byref(wp)):
            return False

        rect = wp.rcNormalPosition
        width = rect.right - rect.left
        height = rect.bottom - rect.top

        # Get primary monitor work area
        pri_x, pri_y, pri_w, pri_h = get_primary_monitor_rect()

        # Calculate new position (centered or at offset)
        # Keep some offset from corner to avoid stacking exactly
        new_x = pri_x + 50 + (hwnd % 10) * 30  # Slight stagger based on hwnd
        new_y = pri_y + 50 + (hwnd % 10) * 30

        # Ensure window fits within primary monitor
        if new_x + width > pri_x + pri_w:
            new_x = pri_x + pri_w - width - 10
        if new_y + height > pri_y + pri_h:
            new_y = pri_y + pri_h - height - 10

        # Update placement
        wp.rcNormalPosition.left = new_x
        wp.rcNormalPosition.top = new_y
        wp.rcNormalPosition.right = new_x + width
        wp.rcNormalPosition.bottom = new_y + height

        return bool(user32.SetWindowPlacement(hwnd, ctypes.byref(wp)))
    except Exception:
        return False


def move_windows_from_monitor(monitor_x: int, monitor_y: int) -> int:
    """Move all windows from a specific monitor to primary.

    Args:
        monitor_x: The monitor's left position
        monitor_y: The monitor's top position

    Returns:
        Number of windows moved
    """
    moved = 0
    for hwnd in get_all_windows():
        try:
            win_mon_x, win_mon_y = get_window_monitor_pos(hwnd)
            if win_mon_x == monitor_x and win_mon_y == monitor_y:
                if move_window_to_primary(hwnd):
                    moved += 1
        except Exception:
            continue
    return moved


def find_window_by_title_and_process(title: str, process_name: str) -> Optional[int]:
    """Find a window by title and process name."""
    for hwnd in get_all_windows():
        try:
            if get_window_title(hwnd) == title and get_process_name(hwnd) == process_name:
                return hwnd
        except Exception:
            continue
    return None


def is_monitor_available(monitor_x: int, monitor_y: int) -> bool:
    """Check if a monitor at the given position exists."""
    # Enumerate all monitors to check
    result = [False]

    def monitor_enum_callback(hmonitor, hdc, lprect, lparam):
        mi = MONITORINFO()
        mi.cbSize = ctypes.sizeof(mi)
        if user32.GetMonitorInfoW(hmonitor, ctypes.byref(mi)):
            if mi.rcMonitor.left == monitor_x and mi.rcMonitor.top == monitor_y:
                result[0] = True
        return True

    MonitorEnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HANDLE, wintypes.HDC,
                                          ctypes.POINTER(RECT), wintypes.LPARAM)
    user32.EnumDisplayMonitors(None, None, MonitorEnumProc(monitor_enum_callback), 0)

    return result[0]


def restore_window_position(saved: WindowPosition) -> bool:
    """Restore a single window to its saved position.

    Returns True if restored, False if skipped/failed.
    """
    try:
        # First, try to find the window by hwnd
        hwnd = saved.hwnd
        if not user32.IsWindow(hwnd):
            # Window handle is invalid, try to find by title + process
            hwnd = find_window_by_title_and_process(saved.title, saved.process_name)
            if not hwnd:
                return False

        # Check if the saved monitor is available
        if not is_monitor_available(saved.monitor_x, saved.monitor_y):
            # Monitor not available, keep window where it is
            return False

        # Restore the window placement
        wp = WINDOWPLACEMENT()
        wp.length = ctypes.sizeof(wp)

        # Get current placement to preserve some fields
        if not user32.GetWindowPlacement(hwnd, ctypes.byref(wp)):
            return False

        # Set saved position
        wp.showCmd = saved.state
        wp.rcNormalPosition.left = saved.x
        wp.rcNormalPosition.top = saved.y
        wp.rcNormalPosition.right = saved.x + saved.width
        wp.rcNormalPosition.bottom = saved.y + saved.height

        return bool(user32.SetWindowPlacement(hwnd, ctypes.byref(wp)))
    except Exception:
        return False


def restore_window_positions(saved_positions: List[WindowPosition]) -> int:
    """Restore multiple windows to their saved positions.

    Returns:
        Number of windows successfully restored
    """
    restored = 0
    for saved in saved_positions:
        if restore_window_position(saved):
            restored += 1
    return restored


def get_monitors_info() -> List[Dict]:
    """Get information about all monitors."""
    monitors = []

    def monitor_enum_callback(hmonitor, hdc, lprect, lparam):
        mi = MONITORINFO()
        mi.cbSize = ctypes.sizeof(mi)
        if user32.GetMonitorInfoW(hmonitor, ctypes.byref(mi)):
            monitors.append({
                "x": mi.rcMonitor.left,
                "y": mi.rcMonitor.top,
                "width": mi.rcMonitor.right - mi.rcMonitor.left,
                "height": mi.rcMonitor.bottom - mi.rcMonitor.top,
                "is_primary": bool(mi.dwFlags & 1),
            })
        return True

    MonitorEnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HANDLE, wintypes.HDC,
                                          ctypes.POINTER(RECT), wintypes.LPARAM)
    user32.EnumDisplayMonitors(None, None, MonitorEnumProc(monitor_enum_callback), 0)

    return monitors


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

    print("=" * 60)
    print("MONITORS:")
    print("=" * 60)
    for m in get_monitors_info():
        primary = " [PRIMARY]" if m["is_primary"] else ""
        print(f"  {m['x']},{m['y']} - {m['width']}x{m['height']}{primary}")

    print()
    print("=" * 60)
    print("WINDOWS:")
    print("=" * 60)
    positions = get_window_positions()
    for p in positions:
        title = p.title[:50].encode('ascii', 'replace').decode('ascii')
        print(f"  [{p.process_name}] {title}")
        print(f"    Position: ({p.x}, {p.y}) Size: {p.width}x{p.height}")
        print(f"    Monitor: ({p.monitor_x}, {p.monitor_y}) State: {p.state}")
        print()
