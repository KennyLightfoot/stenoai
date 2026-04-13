"""
Meeting detector for StenoAI.

Monitors running processes and (on Windows) window titles to detect when a
video meeting starts or ends.  Emits protocol lines to stdout so the Electron
main process can react:

    MEETING_START:<app_name>
    MEETING_END

Run as a subprocess via the 'monitor' CLI command.
"""

import ctypes
import sys
import time
import logging
from typing import Optional

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

logger = logging.getLogger(__name__)

# Process names that indicate an active meeting (lowercased)
MEETING_PROCESSES: dict[str, str] = {
    'zoom.exe': 'Zoom',
    'zoomlauncher.exe': 'Zoom',
    'discord.exe': 'Discord',
    'teams.exe': 'Microsoft Teams',
    'ms-teams.exe': 'Microsoft Teams',
    'slack.exe': 'Slack',
    'webexmta.exe': 'Cisco WebEx',
    'pcoipagent.exe': 'PCoIP',
    # macOS / Linux equivalents
    'zoom': 'Zoom',
    'discord': 'Discord',
    'slack': 'Slack',
}

# Substrings to match in window titles for browser-based meetings
BROWSER_MEETING_TITLE_PATTERNS = [
    'google meet',
    'zoom meeting',
    'microsoft teams',
    'webex',
    'goto meeting',
    'gotomeeting',
    'bluejeans',
    'whereby',
    'around.co',
]

# Browser process names (lowercased) whose window titles we'll inspect
BROWSER_PROCESSES = {
    'chrome.exe', 'msedge.exe', 'firefox.exe', 'brave.exe', 'opera.exe',
    'chromium.exe',
    # macOS / Linux
    'google chrome', 'microsoft edge', 'firefox', 'brave browser',
}


def _get_running_meeting_process() -> Optional[str]:
    """Return the display name of the first detected meeting process, or None."""
    if not PSUTIL_AVAILABLE:
        return None
    try:
        for proc in psutil.process_iter(['name']):
            name = (proc.info.get('name') or '').lower()
            if name in MEETING_PROCESSES:
                return MEETING_PROCESSES[name]
    except Exception as exc:
        logger.debug(f"Process scan error: {exc}")
    return None


# ---------------------------------------------------------------------------
# Windows-only: enumerate all visible window titles
# ---------------------------------------------------------------------------

def _enum_windows_titles() -> list[str]:
    """Return titles of all visible top-level windows on Windows."""
    if sys.platform != 'win32':
        return []
    titles: list[str] = []
    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))

    def _callback(hwnd, _lparam):
        if ctypes.windll.user32.IsWindowVisible(hwnd):
            length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                ctypes.windll.user32.GetWindowTextW(hwnd, buf, length + 1)
                titles.append(buf.value)
        return True

    ctypes.windll.user32.EnumWindows(EnumWindowsProc(_callback), 0)
    return titles


def _get_browser_meeting_name() -> Optional[str]:
    """
    Check open browser windows for meeting title patterns.
    Returns a human-readable name like 'Google Meet', or None.
    """
    if not PSUTIL_AVAILABLE:
        return None

    # First confirm at least one browser is running — avoids the expensive
    # EnumWindows call when no browser is open.
    browser_running = False
    try:
        for proc in psutil.process_iter(['name']):
            if (proc.info.get('name') or '').lower() in BROWSER_PROCESSES:
                browser_running = True
                break
    except Exception:
        pass

    if not browser_running:
        return None

    for title in _enum_windows_titles():
        tl = title.lower()
        for pattern in BROWSER_MEETING_TITLE_PATTERNS:
            if pattern in tl:
                # Map pattern to a nicer name
                nice = {
                    'google meet': 'Google Meet',
                    'zoom meeting': 'Zoom (browser)',
                    'microsoft teams': 'Microsoft Teams (browser)',
                    'webex': 'Cisco WebEx (browser)',
                    'goto meeting': 'GoTo Meeting',
                    'gotomeeting': 'GoTo Meeting',
                    'bluejeans': 'BlueJeans',
                    'whereby': 'Whereby',
                    'around.co': 'Around',
                }.get(pattern, pattern.title())
                return nice
    return None


def detect_meeting() -> Optional[str]:
    """
    Return the name of an active meeting app, or None if no meeting detected.
    Checks native apps first, then browser titles.
    """
    name = _get_running_meeting_process()
    if name:
        return name
    return _get_browser_meeting_name()


def monitor(poll_interval: float = 5.0) -> None:
    """
    Continuously poll for meetings and write protocol lines to stdout.

    Protocol:
        MEETING_START:<app_name>   — emitted once when a meeting is first detected
        MEETING_END                — emitted once when it ends

    This function runs forever; Electron kills the process when auto-detect is toggled off.
    """
    if not PSUTIL_AVAILABLE:
        print('ERROR:psutil not installed; meeting detection unavailable', flush=True)
        return

    in_meeting: bool = False
    current_app: Optional[str] = None

    while True:
        try:
            app = detect_meeting()
            if app and not in_meeting:
                in_meeting = True
                current_app = app
                print(f'MEETING_START:{app}', flush=True)
            elif not app and in_meeting:
                in_meeting = False
                current_app = None
                print('MEETING_END', flush=True)
        except Exception as exc:
            logger.debug(f"Detection error: {exc}")

        time.sleep(poll_interval)
