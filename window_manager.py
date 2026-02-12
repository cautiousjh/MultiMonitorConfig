"""Window position management for DisplaySnap."""
import ctypes
from ctypes import wintypes
from dataclasses import dataclass
from typing import List, Optional, Dict, Set, Tuple
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

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000


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
MonitorEnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HANDLE, wintypes.HDC,
                                      ctypes.POINTER(RECT), wintypes.LPARAM)

# Set up QueryFullProcessImageNameW prototype once
kernel32.QueryFullProcessImageNameW.argtypes = [
    wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)
]
kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL

# Reusable buffers for hot-path functions
_process_buf = ctypes.create_unicode_buffer(260)
_process_size = wintypes.DWORD(260)
_pid_dword = wintypes.DWORD()
_wp_struct = WINDOWPLACEMENT()
_wp_struct.length = ctypes.sizeof(_wp_struct)
_mi_struct = MONITORINFO()
_mi_struct.cbSize = ctypes.sizeof(_mi_struct)


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


# --- Process name cache (PID-based) ---
_pid_cache: Dict[int, str] = {}

SKIP_PROCESSES = frozenset([
    "TextInputHost.exe",
    "ShellExperienceHost.exe",
    "SearchHost.exe",
    "StartMenuExperienceHost.exe",
])

SKIP_TITLES = frozenset([
    "Program Manager",
    "Windows Input Experience",
    "Microsoft Text Input Application",
    "Settings",
])


def _get_pid(hwnd: int) -> int:
    """Get process ID for a window handle."""
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(_pid_dword))
    return _pid_dword.value


def _get_process_name_by_pid(pid: int) -> str:
    """Get process name by PID, using cache.

    Uses QueryFullProcessImageNameW with PROCESS_QUERY_LIMITED_INFORMATION
    for lower permission requirements and fewer API calls.
    """
    cached = _pid_cache.get(pid)
    if cached is not None:
        return cached

    name = ""
    try:
        handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if handle:
            try:
                _process_size.value = 260
                if kernel32.QueryFullProcessImageNameW(
                    handle, 0, _process_buf, ctypes.byref(_process_size)
                ):
                    # Extract filename from full path: "C:\...\app.exe" -> "app.exe"
                    full_path = _process_buf.value
                    name = full_path.rsplit('\\', 1)[-1]
            finally:
                kernel32.CloseHandle(handle)
    except Exception:
        pass

    _pid_cache[pid] = name
    return name


def get_process_name(hwnd: int) -> str:
    """Get the process name for a window handle."""
    return _get_process_name_by_pid(_get_pid(hwnd))


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


def _get_monitor_rects() -> List[Tuple[int, int, int, int, int, int]]:
    """Enumerate all monitors once. Returns list of (left, top, right, bottom, left, top).

    Last two values are the monitor origin (same as left, top) for convenience.
    """
    monitors = []

    def callback(hmonitor, hdc, lprect, lparam):
        if user32.GetMonitorInfoW(hmonitor, ctypes.byref(_mi_struct)):
            r = _mi_struct.rcMonitor
            monitors.append((r.left, r.top, r.right, r.bottom))
        return True

    user32.EnumDisplayMonitors(None, None, MonitorEnumProc(callback), 0)
    return monitors


def _find_monitor_for_point(cx: int, cy: int,
                             monitor_rects: List[Tuple[int, int, int, int]]) -> Tuple[int, int]:
    """Find which monitor contains the given point using pre-enumerated rects."""
    for left, top, right, bottom in monitor_rects:
        if left <= cx < right and top <= cy < bottom:
            return (left, top)
    return (0, 0)


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


def _enum_all_visible_hwnds() -> List[int]:
    """Enumerate all visible top-level window handles (minimal filtering)."""
    windows = []

    def enum_callback(hwnd, lParam):
        if not user32.IsWindowVisible(hwnd):
            return True

        # Quick style check (no API calls beyond GetWindowLong)
        ex_style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        if ex_style & WS_EX_TOOLWINDOW and not (ex_style & WS_EX_APPWINDOW):
            return True

        windows.append(hwnd)
        return True

    user32.EnumWindows(EnumWindowsProc(enum_callback), 0)
    return windows


def get_window_positions() -> List[WindowPosition]:
    """Get positions of all visible windows.

    Optimized:
    - Enumerates windows once
    - Caches process names by PID
    - Retrieves title/process only once per window
    - Pre-enumerates monitor rects for coordinate-based lookup (no per-window API calls)
    - Uses QueryFullProcessImageNameW (lower permissions)
    - Reuses ctypes buffers
    """
    global _pid_cache
    _pid_cache = {}  # Fresh cache per call

    # Pre-enumerate monitor rects once (eliminates 2 API calls per window)
    monitor_rects = _get_monitor_rects()

    positions = []

    for hwnd in _enum_all_visible_hwnds():
        try:
            # Get title first (cheap) for filtering
            title = get_window_title(hwnd)
            if not title:
                continue
            if title in SKIP_TITLES:
                continue

            # Get process name (cached by PID)
            pid = _get_pid(hwnd)
            process_name = _get_process_name_by_pid(pid)
            if process_name in SKIP_PROCESSES:
                continue

            # Get window placement (reuse struct)
            if not user32.GetWindowPlacement(hwnd, ctypes.byref(_wp_struct)):
                continue

            rect = _wp_struct.rcNormalPosition

            # Coordinate-based monitor lookup (no API call)
            cx = (rect.left + rect.right) // 2
            cy = (rect.top + rect.bottom) // 2
            monitor_x, monitor_y = _find_monitor_for_point(cx, cy, monitor_rects)

            # Fallback for minimized/off-screen windows
            if monitor_x == 0 and monitor_y == 0 and monitor_rects:
                first = monitor_rects[0]
                if not (first[0] == 0 and first[1] == 0):
                    # (0,0) isn't a real monitor origin, use MonitorFromWindow
                    monitor_x, monitor_y = get_window_monitor_pos(hwnd)

            pos = WindowPosition(
                hwnd=hwnd,
                title=title,
                process_name=process_name,
                x=rect.left,
                y=rect.top,
                width=rect.right - rect.left,
                height=rect.bottom - rect.top,
                state=_wp_struct.showCmd,
                monitor_x=monitor_x,
                monitor_y=monitor_y,
            )
            positions.append(pos)
        except Exception:
            continue

    return positions


def get_all_windows() -> List[int]:
    """Get all visible user windows (uses full filtering)."""
    global _pid_cache
    _pid_cache = {}

    windows = []
    for hwnd in _enum_all_visible_hwnds():
        title = get_window_title(hwnd)
        if not title or title in SKIP_TITLES:
            continue
        process_name = _get_process_name_by_pid(_get_pid(hwnd))
        if process_name in SKIP_PROCESSES:
            continue
        windows.append(hwnd)
    return windows


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
    return True  # Will be refined when we have the actual monitor list


def _move_window_to_primary(hwnd: int, primary_rect: Tuple[int, int, int, int]) -> bool:
    """Move a window to the primary monitor (with pre-fetched primary rect)."""
    try:
        wp = WINDOWPLACEMENT()
        wp.length = ctypes.sizeof(wp)
        if not user32.GetWindowPlacement(hwnd, ctypes.byref(wp)):
            return False

        rect = wp.rcNormalPosition
        width = rect.right - rect.left
        height = rect.bottom - rect.top

        pri_x, pri_y, pri_w, pri_h = primary_rect

        new_x = pri_x + 50 + (hwnd % 10) * 30
        new_y = pri_y + 50 + (hwnd % 10) * 30

        if new_x + width > pri_x + pri_w:
            new_x = pri_x + pri_w - width - 10
        if new_y + height > pri_y + pri_h:
            new_y = pri_y + pri_h - height - 10

        wp.rcNormalPosition.left = new_x
        wp.rcNormalPosition.top = new_y
        wp.rcNormalPosition.right = new_x + width
        wp.rcNormalPosition.bottom = new_y + height

        return bool(user32.SetWindowPlacement(hwnd, ctypes.byref(wp)))
    except Exception:
        return False


def move_window_to_primary(hwnd: int) -> bool:
    """Move a window to the primary monitor."""
    return _move_window_to_primary(hwnd, get_primary_monitor_rect())


def move_windows_from_monitor(monitor_x: int, monitor_y: int) -> int:
    """Move all windows from a specific monitor to primary.

    Args:
        monitor_x: The monitor's left position
        monitor_y: The monitor's top position

    Returns:
        Number of windows moved
    """
    return move_windows_from_monitors({(monitor_x, monitor_y)})


def move_windows_from_monitors(monitor_positions: Set[Tuple[int, int]]) -> int:
    """Move all windows from multiple monitors to primary in a single pass.

    Args:
        monitor_positions: Set of (monitor_x, monitor_y) tuples

    Returns:
        Number of windows moved
    """
    if not monitor_positions:
        return 0

    # Cache primary rect once for the entire batch
    primary_rect = get_primary_monitor_rect()
    # Pre-enumerate monitor rects for coordinate lookup
    monitor_rects = _get_monitor_rects()

    moved = 0
    for hwnd in _enum_all_visible_hwnds():
        try:
            title = get_window_title(hwnd)
            if not title or title in SKIP_TITLES:
                continue

            # Coordinate-based monitor lookup
            if not user32.GetWindowPlacement(hwnd, ctypes.byref(_wp_struct)):
                continue
            rect = _wp_struct.rcNormalPosition
            cx = (rect.left + rect.right) // 2
            cy = (rect.top + rect.bottom) // 2
            win_mon = _find_monitor_for_point(cx, cy, monitor_rects)

            if win_mon in monitor_positions:
                if _move_window_to_primary(hwnd, primary_rect):
                    moved += 1
        except Exception:
            continue
    return moved


def build_window_lookup() -> Dict[Tuple[str, str], int]:
    """Build a (title, process_name) -> hwnd lookup map in a single pass.

    Used by batch restore to avoid re-enumerating all windows per restore target.
    """
    global _pid_cache
    _pid_cache = {}

    lookup: Dict[Tuple[str, str], int] = {}
    for hwnd in _enum_all_visible_hwnds():
        try:
            title = get_window_title(hwnd)
            if not title:
                continue
            process_name = _get_process_name_by_pid(_get_pid(hwnd))
            lookup[(title, process_name)] = hwnd
        except Exception:
            continue
    return lookup


def find_window_by_title_and_process(title: str, process_name: str) -> Optional[int]:
    """Find a window by title and process name."""
    for hwnd in _enum_all_visible_hwnds():
        try:
            if get_window_title(hwnd) == title and get_process_name(hwnd) == process_name:
                return hwnd
        except Exception:
            continue
    return None


def get_available_monitor_positions() -> Set[Tuple[int, int]]:
    """Get set of all available monitor positions. Single enumeration."""
    return {(left, top) for left, top, _, _ in _get_monitor_rects()}


def is_monitor_available(monitor_x: int, monitor_y: int) -> bool:
    """Check if a monitor at the given position exists."""
    return (monitor_x, monitor_y) in get_available_monitor_positions()


def restore_window_position(saved: WindowPosition,
                             available_monitors: Optional[Set[Tuple[int, int]]] = None,
                             window_lookup: Optional[Dict[Tuple[str, str], int]] = None) -> bool:
    """Restore a single window to its saved position.

    Args:
        saved: The saved window position
        available_monitors: Pre-fetched set of (x, y) monitor positions.
                           If None, will query monitors (slower).
        window_lookup: Pre-built (title, process_name)->hwnd map.
                      If None, falls back to full enumeration (slower).

    Returns True if restored, False if skipped/failed.
    """
    try:
        hwnd = saved.hwnd
        if not user32.IsWindow(hwnd):
            # Use pre-built lookup map (O(1)) or fall back to full scan (O(N))
            if window_lookup is not None:
                hwnd = window_lookup.get((saved.title, saved.process_name), 0)
            else:
                hwnd = find_window_by_title_and_process(saved.title, saved.process_name)
            if not hwnd:
                return False

        # Check if the saved monitor is available
        if available_monitors is not None:
            if (saved.monitor_x, saved.monitor_y) not in available_monitors:
                return False
        else:
            if not is_monitor_available(saved.monitor_x, saved.monitor_y):
                return False

        wp = WINDOWPLACEMENT()
        wp.length = ctypes.sizeof(wp)

        if not user32.GetWindowPlacement(hwnd, ctypes.byref(wp)):
            return False

        wp.showCmd = saved.state
        wp.rcNormalPosition.left = saved.x
        wp.rcNormalPosition.top = saved.y
        wp.rcNormalPosition.right = saved.x + saved.width
        wp.rcNormalPosition.bottom = saved.y + saved.height

        return bool(user32.SetWindowPlacement(hwnd, ctypes.byref(wp)))
    except Exception:
        return False


def restore_window_positions(saved_positions: List[WindowPosition],
                              skip_minimized: bool = True) -> int:
    """Restore multiple windows to their saved positions.

    Args:
        saved_positions: List of saved window positions
        skip_minimized: If True, skip minimized windows (default)

    Returns:
        Number of windows successfully restored
    """
    # Pre-fetch available monitors and window lookup map once
    available_monitors = get_available_monitor_positions()
    wl = build_window_lookup()

    restored = 0
    for saved in saved_positions:
        if skip_minimized and saved.state == SW_SHOWMINIMIZED:
            continue
        if restore_window_position(saved, available_monitors=available_monitors, window_lookup=wl):
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

    user32.EnumDisplayMonitors(None, None, MonitorEnumProc(monitor_enum_callback), 0)

    return monitors


if __name__ == "__main__":
    import sys
    import time
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
    start = time.perf_counter()
    positions = get_window_positions()
    elapsed = time.perf_counter() - start
    for p in positions:
        title = p.title[:50].encode('ascii', 'replace').decode('ascii')
        print(f"  [{p.process_name}] {title}")
        print(f"    Position: ({p.x}, {p.y}) Size: {p.width}x{p.height}")
        print(f"    Monitor: ({p.monitor_x}, {p.monitor_y}) State: {p.state}")
        print()
    print(f"Retrieved {len(positions)} windows in {elapsed*1000:.1f}ms")
