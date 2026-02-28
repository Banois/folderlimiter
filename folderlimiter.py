import ctypes
from ctypes import wintypes
from datetime import datetime
import json
import os
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk
if os.name == "nt":
    import winreg
else:
    winreg = None

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.abspath(__file__)
CONFIG_PATH = os.path.join(PROJECT_DIR, "config.json")
RUNNER_VBS_PATH = os.path.join(PROJECT_DIR, "run_monitor_silent.vbs")
DEFAULT_WATCH_PATH = r"C:\Users\conner\Desktop\a) folder central\a)projects\monitor\test"

DEFAULT_CONFIG = {
    "monitored_paths": [],
    "check_interval_seconds": 5,
    "auto_minimize_to_tray": False,
    "start_on_startup": False,
}

STARTUP_REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
STARTUP_VALUE_NAME = "FolderMonitorApp"

SIZE_RE = re.compile(r"^\s*(\d+(?:\.\d+)?)\s*([a-zA-Z]{0,2})\s*$")
UNIT_FACTORS = {
    "b": 1,
    "kb": 1024,
    "mb": 1024 * 1024,
    "gb": 1024 * 1024 * 1024,
    "tb": 1024 * 1024 * 1024 * 1024,
}

LRESULT = ctypes.c_ssize_t


if os.name == "nt":
    ERROR_MORE_DATA = 234
    CCH_RM_MAX_APP_NAME = 255
    CCH_RM_MAX_SVC_NAME = 63
    CCH_RM_SESSION_KEY = 32

    class RM_UNIQUE_PROCESS(ctypes.Structure):
        _fields_ = [
            ("dwProcessId", wintypes.DWORD),
            ("ProcessStartTime", wintypes.FILETIME),
        ]

    class RM_PROCESS_INFO(ctypes.Structure):
        _fields_ = [
            ("Process", RM_UNIQUE_PROCESS),
            ("strAppName", wintypes.WCHAR * (CCH_RM_MAX_APP_NAME + 1)),
            ("strServiceShortName", wintypes.WCHAR * (CCH_RM_MAX_SVC_NAME + 1)),
            ("ApplicationType", wintypes.UINT),
            ("AppStatus", wintypes.DWORD),
            ("TSSessionId", wintypes.DWORD),
            ("bRestartable", wintypes.BOOL),
        ]

    _rstrtmgr = ctypes.WinDLL("Rstrtmgr")
    _rstrtmgr.RmStartSession.argtypes = [ctypes.POINTER(wintypes.DWORD), wintypes.DWORD, wintypes.LPWSTR]
    _rstrtmgr.RmStartSession.restype = wintypes.DWORD
    _rstrtmgr.RmRegisterResources.argtypes = [
        wintypes.DWORD,
        wintypes.UINT,
        ctypes.POINTER(wintypes.LPCWSTR),
        wintypes.UINT,
        ctypes.POINTER(RM_UNIQUE_PROCESS),
        wintypes.UINT,
        ctypes.POINTER(wintypes.LPCWSTR),
    ]
    _rstrtmgr.RmRegisterResources.restype = wintypes.DWORD
    _rstrtmgr.RmGetList.argtypes = [
        wintypes.DWORD,
        ctypes.POINTER(wintypes.UINT),
        ctypes.POINTER(wintypes.UINT),
        ctypes.POINTER(RM_PROCESS_INFO),
        ctypes.POINTER(wintypes.DWORD),
    ]
    _rstrtmgr.RmGetList.restype = wintypes.DWORD
    _rstrtmgr.RmEndSession.argtypes = [wintypes.DWORD]
    _rstrtmgr.RmEndSession.restype = wintypes.DWORD
else:
    _rstrtmgr = None

RM_APP_TYPE_NAMES = {
    0: "Unknown",
    1: "Main Window",
    2: "Other Window",
    3: "Service",
    4: "Explorer",
    5: "Console",
    6: "Critical",
    1000: "Invalid",
}


def is_locked_file_error(exc: OSError) -> bool:
    return getattr(exc, "winerror", None) in (32, 33)


def list_locking_processes(file_path: str) -> list[dict]:
    if os.name != "nt" or _rstrtmgr is None:
        return []
    if not file_path:
        return []

    session_handle = wintypes.DWORD(0)
    session_key = ctypes.create_unicode_buffer(CCH_RM_SESSION_KEY + 1)
    result = _rstrtmgr.RmStartSession(ctypes.byref(session_handle), 0, session_key)
    if result != 0:
        return []

    try:
        resources = (wintypes.LPCWSTR * 1)(file_path)
        result = _rstrtmgr.RmRegisterResources(session_handle, 1, resources, 0, None, 0, None)
        if result != 0:
            return []

        needed = wintypes.UINT(0)
        count = wintypes.UINT(0)
        reboot_reasons = wintypes.DWORD(0)

        result = _rstrtmgr.RmGetList(
            session_handle,
            ctypes.byref(needed),
            ctypes.byref(count),
            None,
            ctypes.byref(reboot_reasons),
        )
        if result == 0:
            return []
        if result != ERROR_MORE_DATA or needed.value == 0:
            return []

        process_info = (RM_PROCESS_INFO * needed.value)()
        count = wintypes.UINT(needed.value)
        result = _rstrtmgr.RmGetList(
            session_handle,
            ctypes.byref(needed),
            ctypes.byref(count),
            process_info,
            ctypes.byref(reboot_reasons),
        )
        if result != 0:
            return []

        output = []
        for index in range(count.value):
            info = process_info[index]
            pid = int(info.Process.dwProcessId)
            name = info.strAppName.strip() or f"PID {pid}"
            service = info.strServiceShortName.strip()
            output.append(
                {
                    "pid": pid,
                    "name": name,
                    "type": RM_APP_TYPE_NAMES.get(int(info.ApplicationType), str(int(info.ApplicationType))),
                    "service": service,
                }
            )
        output.sort(key=lambda item: (item["pid"], item["name"]))
        return output
    except Exception:
        return []
    finally:
        _rstrtmgr.RmEndSession(session_handle)


def parse_size_to_bytes(value: str) -> int:
    text = str(value).strip().lower()
    match = SIZE_RE.match(text)
    if not match:
        raise ValueError("Invalid size format. Example values: 0.5, 5mb, 1.2gb")
    amount = float(match.group(1))
    unit = match.group(2) or "gb"
    if unit not in UNIT_FACTORS:
        raise ValueError("Unsupported unit. Use b, kb, mb, gb, or tb")
    size_bytes = int(amount * UNIT_FACTORS[unit])
    if size_bytes <= 0:
        raise ValueError("Size must be greater than zero")
    return size_bytes


def human_size(num_bytes: int) -> str:
    units = ["B", "KB", "MB", "GB", "TB"]
    value = float(num_bytes)
    index = 0
    while value >= 1024 and index < len(units) - 1:
        value /= 1024.0
        index += 1
    return f"{value:.2f} {units[index]}"


def parse_yes_no(value: str) -> bool:
    text = str(value).strip().lower()
    if text in ("y", "yes"):
        return True
    if text in ("n", "no"):
        return False
    raise ValueError("Please enter y or n")


def normalize_delete_mode(value: str) -> str:
    mode = str(value).strip().lower()
    if mode in ("latest", "newest", "most recent"):
        return "latest"
    if mode in ("earliest", "oldest"):
        return "earliest"
    return "earliest"


def normalize_monitored_path_entry(entry: dict) -> dict | None:
    if not isinstance(entry, dict):
        return None
    raw_path = str(entry.get("path", "")).strip()
    if not raw_path:
        return None
    abs_path = os.path.abspath(raw_path)
    limit_input = str(entry.get("limit_input", "0.5")).strip().lower()
    parse_size_to_bytes(limit_input)
    delete_mode = normalize_delete_mode(entry.get("delete_mode", "earliest"))
    return {
        "path": abs_path,
        "limit_input": limit_input,
        "delete_mode": delete_mode,
    }


def normalize_config(raw_config: dict | None) -> dict:
    raw = raw_config if isinstance(raw_config, dict) else {}
    cfg = dict(DEFAULT_CONFIG)
    if raw:
        cfg.update(raw)

    interval_seconds = int(cfg.get("check_interval_seconds", 5))
    if interval_seconds < 1:
        interval_seconds = 1
    cfg["check_interval_seconds"] = interval_seconds

    cfg["auto_minimize_to_tray"] = bool(cfg.get("auto_minimize_to_tray", False))
    if "start_on_startup" in raw:
        cfg["start_on_startup"] = bool(raw.get("start_on_startup", False))
    else:
        cfg["start_on_startup"] = is_startup_enabled()

    has_modern_paths = isinstance(raw.get("monitored_paths"), list)
    monitored_paths = cfg.get("monitored_paths")
    if not has_modern_paths:
        # Legacy migration from single-path config when modern list is absent.
        legacy_path = str(raw.get("watch_path", "")).strip()
        legacy_limit = str(raw.get("limit_input", "0.5")).strip().lower() or "0.5"
        legacy_mode = normalize_delete_mode(raw.get("delete_mode", "earliest"))
        monitored_paths = []
        if legacy_path:
            abs_legacy_path = os.path.abspath(legacy_path)
            # Do not auto-reinsert the old project default path.
            if abs_legacy_path != os.path.abspath(DEFAULT_WATCH_PATH):
                monitored_paths.append(
                    {
                        "path": abs_legacy_path,
                        "limit_input": legacy_limit,
                        "delete_mode": legacy_mode,
                    }
                )
    elif not isinstance(monitored_paths, list):
        monitored_paths = []

    normalized_paths = []
    seen_paths = set()
    for entry in monitored_paths:
        normalized_entry = normalize_monitored_path_entry(entry)
        if not normalized_entry:
            continue
        if normalized_entry["path"] in seen_paths:
            continue
        seen_paths.add(normalized_entry["path"])
        normalized_paths.append(normalized_entry)

    cfg["monitored_paths"] = normalized_paths
    return cfg


def load_config() -> dict | None:
    if not os.path.exists(CONFIG_PATH):
        return None
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return normalize_config(data)
    except Exception:
        return None


def save_config(config: dict) -> None:
    normalized = normalize_config(config)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(normalized, f, indent=2)


def get_startup_command() -> str:
    return f'wscript.exe "{RUNNER_VBS_PATH}"'


def is_startup_enabled() -> bool:
    if os.name != "nt" or winreg is None:
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_REG_PATH, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, STARTUP_VALUE_NAME)
            return bool(str(value).strip())
    except FileNotFoundError:
        return False
    except OSError:
        return False


def set_startup_enabled(enabled: bool) -> tuple[bool, str]:
    if os.name != "nt" or winreg is None:
        return False, "Windows startup integration is only available on Windows."
    try:
        ensure_runner_script()
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_REG_PATH, 0, winreg.KEY_SET_VALUE) as key:
            if enabled:
                winreg.SetValueEx(key, STARTUP_VALUE_NAME, 0, winreg.REG_SZ, get_startup_command())
            else:
                try:
                    winreg.DeleteValue(key, STARTUP_VALUE_NAME)
                except FileNotFoundError:
                    pass
        return True, ""
    except OSError as exc:
        return False, str(exc)


def ensure_runner_script() -> None:
    script_name = os.path.basename(SCRIPT_PATH)
    content = """Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
cmd = "pythonw \"" & scriptDir & "\\{script_name}\""
CreateObject("Wscript.Shell").Run cmd, 0, False
""".format(script_name=script_name.replace('"', '""'))
    with open(RUNNER_VBS_PATH, "w", encoding="utf-8") as f:
        f.write(content)


def restart_self() -> None:
    subprocess.Popen([get_preferred_gui_python(), SCRIPT_PATH], cwd=PROJECT_DIR)
    os._exit(0)


def get_preferred_gui_python() -> str:
    if os.name != "nt":
        return sys.executable
    exe_dir = os.path.dirname(sys.executable)
    pythonw_path = os.path.join(exe_dir, "pythonw.exe")
    if os.path.exists(pythonw_path):
        return pythonw_path
    return sys.executable


def ensure_gui_only_process() -> None:
    if os.name != "nt":
        return
    if "--console" in sys.argv:
        return
    if os.path.basename(sys.executable).lower() == "pythonw.exe":
        return
    if os.environ.get("MONITOR_GUI_ONLY") == "1":
        return

    kernel32 = ctypes.windll.kernel32
    kernel32.GetConsoleWindow.restype = wintypes.HWND
    console_hwnd = kernel32.GetConsoleWindow()
    if not console_hwnd:
        return

    gui_python = get_preferred_gui_python()
    if os.path.basename(gui_python).lower() == "pythonw.exe":
        env = os.environ.copy()
        env["MONITOR_GUI_ONLY"] = "1"
        subprocess.Popen([gui_python, SCRIPT_PATH], cwd=PROJECT_DIR, env=env)
        os._exit(0)

    # Fallback when pythonw is unavailable.
    SW_HIDE = 0
    ctypes.windll.user32.ShowWindow(console_hwnd, SW_HIDE)


def ask_non_empty_string(root: tk.Tk, title: str, prompt: str, default: str) -> str | None:
    while True:
        result = simpledialog.askstring(title, prompt, parent=root, initialvalue=default)
        if result is None:
            return None
        result = result.strip()
        if result:
            return result
        messagebox.showerror("Setup", "This value cannot be empty.", parent=root)


def run_setup_wizard(root: tk.Tk, starting_values: dict | None = None) -> dict | None:
    defaults = normalize_config(starting_values)

    messagebox.showinfo(
        "Initial Setup",
        "Set up monitor settings. Path/limit/delete behavior is managed from View Paths.",
        parent=root,
    )

    while True:
        interval_input = ask_non_empty_string(
            root,
            "Setup",
            "How often to check (seconds):",
            str(defaults["check_interval_seconds"]),
        )
        if interval_input is None:
            return None
        try:
            interval_seconds = int(interval_input)
            if interval_seconds < 1:
                raise ValueError
            break
        except ValueError:
            messagebox.showerror("Setup", "Enter a whole number >= 1.", parent=root)

    while True:
        tray_input = ask_non_empty_string(
            root,
            "Setup",
            "Auto-minimize to tray on startup? (y/n)",
            "y" if defaults["auto_minimize_to_tray"] else "n",
        )
        if tray_input is None:
            return None
        try:
            auto_minimize = parse_yes_no(tray_input)
            break
        except ValueError as exc:
            messagebox.showerror("Setup", str(exc), parent=root)

    while True:
        startup_input = ask_non_empty_string(
            root,
            "Setup",
            "Start this program when Windows starts? (y/n)",
            "y" if defaults.get("start_on_startup", False) else "n",
        )
        if startup_input is None:
            return None
        try:
            start_on_startup = parse_yes_no(startup_input)
            break
        except ValueError as exc:
            messagebox.showerror("Setup", str(exc), parent=root)

    config = {
        "monitored_paths": defaults["monitored_paths"],
        "check_interval_seconds": interval_seconds,
        "auto_minimize_to_tray": auto_minimize,
        "start_on_startup": start_on_startup,
    }
    return normalize_config(config)


class POINT(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


class MSG(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("message", wintypes.UINT),
        ("wParam", wintypes.WPARAM),
        ("lParam", wintypes.LPARAM),
        ("time", wintypes.DWORD),
        ("pt", POINT),
    ]


WNDPROC = ctypes.WINFUNCTYPE(
    LRESULT,
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM,
)


class WNDCLASSW(ctypes.Structure):
    _fields_ = [
        ("style", wintypes.UINT),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wintypes.HINSTANCE),
        ("hIcon", wintypes.HANDLE),
        ("hCursor", wintypes.HANDLE),
        ("hbrBackground", wintypes.HANDLE),
        ("lpszMenuName", wintypes.LPCWSTR),
        ("lpszClassName", wintypes.LPCWSTR),
    ]


class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hWnd", wintypes.HWND),
        ("uID", wintypes.UINT),
        ("uFlags", wintypes.UINT),
        ("uCallbackMessage", wintypes.UINT),
        ("hIcon", wintypes.HICON),
        ("szTip", wintypes.WCHAR * 128),
        ("dwState", wintypes.DWORD),
        ("dwStateMask", wintypes.DWORD),
        ("szInfo", wintypes.WCHAR * 256),
        ("uTimeoutOrVersion", wintypes.UINT),
        ("szInfoTitle", wintypes.WCHAR * 64),
        ("dwInfoFlags", wintypes.DWORD),
    ]


class TrayIcon:
    WM_USER = 0x0400
    WM_DESTROY = 0x0002
    WM_CLOSE = 0x0010
    WM_LBUTTONDBLCLK = 0x0203
    WM_RBUTTONUP = 0x0205
    WM_APP = 0x8000
    NIN_BALLOONUSERCLICK = WM_USER + 5

    NIM_ADD = 0x00000000
    NIM_MODIFY = 0x00000001
    NIM_DELETE = 0x00000002

    NIF_MESSAGE = 0x00000001
    NIF_ICON = 0x00000002
    NIF_TIP = 0x00000004
    NIF_INFO = 0x00000010

    IDI_APPLICATION = 32512
    IDC_ARROW = 32512

    def __init__(self, on_open, on_menu, on_notification_click=None):
        self.on_open = on_open
        self.on_menu = on_menu
        self.on_notification_click = on_notification_click
        self.user32 = ctypes.windll.user32
        self.shell32 = ctypes.windll.shell32
        self.kernel32 = ctypes.windll.kernel32
        self.callback_message = self.WM_APP + 1
        self.thread = None
        self.hwnd = None
        self.class_name = f"FolderMonitorTrayClass_{os.getpid()}"
        self.nid = None
        self._started = False
        self._ready = threading.Event()
        self._wndproc = None
        self._configure_win_api()

    def _configure_win_api(self) -> None:
        self.user32.DefWindowProcW.argtypes = [
            wintypes.HWND,
            wintypes.UINT,
            wintypes.WPARAM,
            wintypes.LPARAM,
        ]
        self.user32.DefWindowProcW.restype = LRESULT
        self.shell32.Shell_NotifyIconW.argtypes = [wintypes.DWORD, ctypes.POINTER(NOTIFYICONDATAW)]
        self.shell32.Shell_NotifyIconW.restype = wintypes.BOOL

    def show_notification(self, title: str, message: str) -> bool:
        if not self._started or self.nid is None:
            return False
        self.nid.uFlags = self.NIF_INFO
        self.nid.szInfoTitle = str(title)[:63]
        self.nid.szInfo = str(message)[:255]
        self.nid.dwInfoFlags = 0
        self.nid.uTimeoutOrVersion = 10000
        return bool(self.shell32.Shell_NotifyIconW(self.NIM_MODIFY, ctypes.byref(self.nid)))

    def start(self) -> bool:
        if self.thread and self.thread.is_alive():
            return True
        self.thread = threading.Thread(target=self._run_message_loop, daemon=True)
        self.thread.start()
        self._ready.wait(timeout=3)
        return self._started

    def stop(self) -> None:
        if self.hwnd:
            self.user32.PostMessageW(self.hwnd, self.WM_CLOSE, 0, 0)
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=2)

    def _run_message_loop(self) -> None:
        h_instance = self.kernel32.GetModuleHandleW(None)

        @WNDPROC
        def wndproc(hwnd, msg, wparam, lparam):
            if msg == self.callback_message:
                if lparam == self.WM_LBUTTONDBLCLK:
                    self.on_open()
                elif lparam == self.WM_RBUTTONUP:
                    point = POINT()
                    self.user32.GetCursorPos(ctypes.byref(point))
                    self.on_menu(point.x, point.y)
                elif lparam == self.NIN_BALLOONUSERCLICK and self.on_notification_click:
                    self.on_notification_click()
                return 0

            if msg == self.WM_CLOSE:
                if self.nid is not None:
                    self.shell32.Shell_NotifyIconW(self.NIM_DELETE, ctypes.byref(self.nid))
                self.user32.DestroyWindow(hwnd)
                return 0

            if msg == self.WM_DESTROY:
                self.user32.PostQuitMessage(0)
                return 0

            return self.user32.DefWindowProcW(hwnd, msg, wparam, lparam)

        self._wndproc = wndproc

        window_class = WNDCLASSW()
        window_class.style = 0
        window_class.lpfnWndProc = self._wndproc
        window_class.cbClsExtra = 0
        window_class.cbWndExtra = 0
        window_class.hInstance = h_instance
        window_class.hIcon = self.user32.LoadIconW(None, self.IDI_APPLICATION)
        window_class.hCursor = self.user32.LoadCursorW(None, self.IDC_ARROW)
        window_class.hbrBackground = 0
        window_class.lpszMenuName = None
        window_class.lpszClassName = self.class_name

        self.user32.RegisterClassW(ctypes.byref(window_class))

        self.hwnd = self.user32.CreateWindowExW(
            0,
            self.class_name,
            "FolderMonitorTrayWindow",
            0,
            0,
            0,
            0,
            0,
            None,
            None,
            h_instance,
            None,
        )

        if not self.hwnd:
            self._started = False
            self._ready.set()
            return

        icon_data = NOTIFYICONDATAW()
        icon_data.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        icon_data.hWnd = self.hwnd
        icon_data.uID = 1
        icon_data.uFlags = self.NIF_ICON | self.NIF_MESSAGE | self.NIF_TIP
        icon_data.uCallbackMessage = self.callback_message
        icon_data.hIcon = self.user32.LoadIconW(None, self.IDI_APPLICATION)
        icon_data.szTip = "Folder Monitor"

        if not self.shell32.Shell_NotifyIconW(self.NIM_ADD, ctypes.byref(icon_data)):
            self._started = False
            self._ready.set()
            return

        self.nid = icon_data
        self._started = True
        self._ready.set()

        msg = MSG()
        while self.user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            self.user32.TranslateMessage(ctypes.byref(msg))
            self.user32.DispatchMessageW(ctypes.byref(msg))

        self._started = False


class MonitorApp:
    def __init__(self, root: tk.Tk, config: dict):
        self.root = root
        self.config = normalize_config(config)
        self.interval_ms = int(self.config["check_interval_seconds"]) * 1000
        self.is_exiting = False
        self.missing_paths_notified = set()
        self.path_window = None
        self.path_tree = None
        self.footer_var = tk.StringVar(value="")

        self.root.title("Folder Monitor")
        self.root.geometry("800x500")
        self.root.minsize(700, 420)

        self.summary_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.lock_notified_paths = set()
        self.lock_popup = None
        self.lock_popup_file_path = None
        self.lock_popup_text_var = tk.StringVar(value="")
        self.lock_popup_after_id = None
        self.lock_window = None
        self.lock_window_file_path = None
        self.lock_window_file_var = tk.StringVar(value="")
        self.lock_window_status_var = tk.StringVar(value="")
        self.lock_tree = None
        self.lock_refresh_after_id = None
        self.pending_locked_file_notification = None

        self._build_ui()
        self._update_summary()
        self._sync_startup_setting(show_popup=False)

        self.tray_menu = tk.Menu(self.root, tearoff=0)
        self.tray_menu.add_command(label="Open Monitor", command=self.show_window)
        self.tray_menu.add_separator()
        self.tray_menu.add_command(label="Exit", command=self.exit_app)

        self.tray_icon = TrayIcon(self._on_tray_open, self._on_tray_menu, self._on_tray_notification_click)
        self.tray_available = self.tray_icon.start()
        if not self.tray_available:
            self.log("Tray icon failed to start. Close button will exit the app.")

        self.root.protocol("WM_DELETE_WINDOW", self.on_window_close)
        self.log("Monitoring started.")
        self._schedule_check(initial=True)

        if self.config["auto_minimize_to_tray"] and self.tray_available:
            self.root.after(300, self.hide_to_tray)
        else:
            self.root.deiconify()

    def _build_ui(self) -> None:
        header = tk.Frame(self.root, padx=12, pady=10)
        header.pack(fill="x")

        tk.Label(
            header,
            text="Folder Monitor",
            font=("Segoe UI", 14, "bold"),
        ).pack(anchor="w")

        tk.Label(
            header,
            textvariable=self.summary_var,
            justify="left",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(6, 0))

        controls = tk.Frame(self.root, padx=12, pady=8)
        controls.pack(fill="x")

        tk.Button(controls, text="Run Check Now", command=self.check_once, width=16).pack(side="left")
        tk.Button(controls, text="View Paths", command=self.open_paths_window, width=12).pack(
            side="left", padx=(8, 0)
        )
        tk.Button(controls, text="Open Monitored Folder", command=self.open_watch_folder, width=20).pack(
            side="left", padx=(8, 0)
        )
        tk.Button(controls, text="Reconfigure", command=self.reconfigure, width=12).pack(side="left", padx=(8, 0))
        tk.Button(controls, text="Exit", command=self.exit_app, width=10).pack(side="right")

        tk.Label(
            self.root,
            textvariable=self.status_var,
            anchor="w",
            padx=12,
            font=("Segoe UI", 10),
        ).pack(fill="x")

        self.log_box = scrolledtext.ScrolledText(self.root, height=18, font=("Consolas", 10), state="disabled")
        self.log_box.pack(fill="both", expand=True, padx=12, pady=(4, 12))

        tk.Label(
            self.root,
            textvariable=self.footer_var,
            anchor="w",
            padx=12,
            pady=6,
            fg="#666666",
            font=("Segoe UI", 8),
        ).pack(fill="x", side="bottom")

    def _update_summary(self) -> None:
        path_count = len(self.config["monitored_paths"])
        text = f"Monitored paths: {path_count}"
        self.summary_var.set(text)
        self.footer_var.set(
            "Hide to tray on startup: "
            f"{'Yes' if self.config['auto_minimize_to_tray'] else 'No'}"
            " | Start on Windows startup: "
            f"{'Yes' if self.config.get('start_on_startup', False) else 'No'}"
        )

    def log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{timestamp}] {message}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _sync_startup_setting(self, show_popup: bool) -> bool:
        enabled = bool(self.config.get("start_on_startup", False))
        ok, error = set_startup_enabled(enabled)
        if ok:
            return True
        self.log(f"Startup setting could not be applied: {error}")
        if show_popup:
            messagebox.showwarning(
                "Startup Setting",
                "Could not apply startup preference.\n\n" + error,
                parent=self.root,
            )
        return False

    def _delete_mode_label(self, mode: str) -> str:
        return "Latest" if normalize_delete_mode(mode) == "latest" else "Earliest"

    def _close_paths_window(self) -> None:
        if self.path_window and self.path_window.winfo_exists():
            self.path_window.destroy()
        self.path_window = None
        self.path_tree = None

    def _refresh_paths_tree(self) -> None:
        if not self.path_tree:
            return
        for item_id in self.path_tree.get_children():
            self.path_tree.delete(item_id)

        for index, entry in enumerate(self.config["monitored_paths"]):
            limit_bytes = parse_size_to_bytes(entry["limit_input"])
            self.path_tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    entry["path"],
                    f"{entry['limit_input']} ({human_size(limit_bytes)})",
                    self._delete_mode_label(entry["delete_mode"]),
                ),
            )

    def _remove_path_by_click(self, event) -> None:
        if not self.path_tree:
            return
        row_id = self.path_tree.identify_row(event.y)
        if not row_id:
            return
        try:
            index = int(row_id)
        except ValueError:
            return
        if index < 0 or index >= len(self.config["monitored_paths"]):
            return

        entry = self.config["monitored_paths"][index]
        if not messagebox.askyesno(
            "Remove Path",
            f"Remove monitored path?\n\n{entry['path']}",
            parent=self.path_window,
        ):
            return

        self.config["monitored_paths"].pop(index)
        save_config(self.config)
        self.config = normalize_config(self.config)
        self._update_summary()
        self._refresh_paths_tree()
        self.log(f"Removed monitored path: {entry['path']}")
        self.status_var.set("Monitored paths updated.")
        self.root.after(100, self.check_once)

    def _add_path_dialog(self) -> None:
        default_path = self.config["monitored_paths"][0]["path"] if self.config["monitored_paths"] else ""
        abs_path = None
        if os.name == "nt":
            try:
                picker_kwargs = {
                    "parent": self.path_window or self.root,
                    "title": "Select Folder To Monitor",
                    "mustexist": False,
                }
                if default_path:
                    picker_kwargs["initialdir"] = default_path
                selected = filedialog.askdirectory(**picker_kwargs)
                if selected:
                    abs_path = os.path.abspath(selected.strip())
            except Exception as exc:
                self.log(f"Folder picker failed, falling back to manual path input: {exc}")

        if not abs_path:
            path_text = ask_non_empty_string(
                self.root,
                "Add Path",
                "Path to monitor:",
                default_path,
            )
            if path_text is None:
                return
            abs_path = os.path.abspath(path_text.strip())

        existing_paths = {entry["path"].lower() for entry in self.config["monitored_paths"]}
        if abs_path.lower() in existing_paths:
            messagebox.showinfo("Add Path", "That path is already in the monitored list.", parent=self.path_window)
            return

        while True:
            limit_input = ask_non_empty_string(
                self.root,
                "Add Path",
                "Folder size limit (GB by default). Examples: 0.5, 5mb, 1.2gb",
                "0.5",
            )
            if limit_input is None:
                return
            try:
                parse_size_to_bytes(limit_input)
                break
            except ValueError as exc:
                messagebox.showerror("Add Path", str(exc), parent=self.path_window)

        while True:
            mode_choice = ask_non_empty_string(
                self.root,
                "Add Path",
                "Delete mode when over limit: 1 = latest (newest), 2 = earliest (oldest)",
                "2",
            )
            if mode_choice is None:
                return
            mode_choice = mode_choice.strip()
            if mode_choice == "1":
                delete_mode = "latest"
                break
            if mode_choice == "2":
                delete_mode = "earliest"
                break
            messagebox.showerror("Add Path", "Enter 1 or 2.", parent=self.path_window)

        self.config["monitored_paths"].append(
            {
                "path": abs_path,
                "limit_input": limit_input.strip().lower(),
                "delete_mode": delete_mode,
            }
        )
        save_config(self.config)
        self.config = normalize_config(self.config)
        self._update_summary()
        self._refresh_paths_tree()
        self.log(f"Added monitored path: {abs_path} | limit {limit_input} | {self._delete_mode_label(delete_mode)}")
        self.status_var.set("Monitored paths updated.")
        self.root.after(100, self.check_once)

    def open_paths_window(self) -> None:
        if self.path_window and self.path_window.winfo_exists():
            self.path_window.deiconify()
            self.path_window.lift()
            self.path_window.focus_force()
            self._refresh_paths_tree()
            return

        window = tk.Toplevel(self.root)
        self.path_window = window
        window.title("Monitored Paths")
        window.geometry("860x320")
        window.minsize(700, 260)
        window.protocol("WM_DELETE_WINDOW", self._close_paths_window)

        top = tk.Frame(window, padx=12, pady=10)
        top.pack(fill="x")
        tk.Label(top, text="Click a path row to remove it.", font=("Segoe UI", 10)).pack(anchor="w")

        table_frame = tk.Frame(window, padx=12, pady=0)
        table_frame.pack(fill="both", expand=True, pady=(0, 8))
        self.path_tree = ttk.Treeview(
            table_frame,
            columns=("path", "limit", "mode"),
            show="headings",
            selectmode="browse",
        )
        self.path_tree.heading("path", text="Path")
        self.path_tree.heading("limit", text="Limit")
        self.path_tree.heading("mode", text="Delete Over Limit")
        self.path_tree.column("path", width=520, anchor="w")
        self.path_tree.column("limit", width=170, anchor="w")
        self.path_tree.column("mode", width=140, anchor="center")
        self.path_tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.path_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.path_tree.configure(yscrollcommand=scrollbar.set)
        self.path_tree.bind("<ButtonRelease-1>", self._remove_path_by_click)

        actions = tk.Frame(window, padx=12, pady=0)
        actions.pack(fill="x", pady=(0, 12))
        tk.Button(actions, text="Add Path", command=self._add_path_dialog, width=12).pack(side="left")
        tk.Button(actions, text="Close", command=self._close_paths_window, width=10).pack(side="right")

        self._refresh_paths_tree()

    def _position_lock_popup(self) -> None:
        if not self.lock_popup or not self.lock_popup.winfo_exists():
            return
        self.lock_popup.update_idletasks()
        width = self.lock_popup.winfo_width()
        height = self.lock_popup.winfo_height()
        screen_width = self.lock_popup.winfo_screenwidth()
        screen_height = self.lock_popup.winfo_screenheight()
        x = max(0, screen_width - width - 20)
        y = max(0, screen_height - height - 60)
        self.lock_popup.geometry(f"+{x}+{y}")

    def _dismiss_lock_popup(self) -> None:
        if self.lock_popup_after_id and self.lock_popup and self.lock_popup.winfo_exists():
            self.lock_popup.after_cancel(self.lock_popup_after_id)
        self.lock_popup_after_id = None
        if self.lock_popup and self.lock_popup.winfo_exists():
            self.lock_popup.destroy()
        self.lock_popup = None
        self.lock_popup_file_path = None

    def _open_lock_details(self, file_path: str | None) -> None:
        if not file_path:
            return
        self.show_window()
        self.open_lock_window(file_path)

    def _open_lock_details_from_popup(self, _event=None) -> None:
        file_path = self.lock_popup_file_path
        self._dismiss_lock_popup()
        self._open_lock_details(file_path)

    def _on_tray_notification_click(self) -> None:
        self.root.after(0, self._open_lock_details_from_tray_notification)

    def _open_lock_details_from_tray_notification(self) -> None:
        file_path = self.pending_locked_file_notification
        self.pending_locked_file_notification = None
        self._open_lock_details(file_path)

    def _notify_locked_file(self, file_path: str) -> None:
        self.pending_locked_file_notification = file_path
        name = os.path.basename(file_path) or file_path
        if os.name == "nt" and self.tray_available:
            if self.tray_icon.show_notification("File Locked", f"{name} is in use. Click notification for details."):
                return
        self._show_locked_file_popup(file_path)

    def _show_locked_file_popup(self, file_path: str) -> None:
        self.lock_popup_file_path = file_path
        name = os.path.basename(file_path) or file_path
        self.lock_popup_text_var.set(
            f"File is being used by another process:\n{name}\n\nClick for process details."
        )

        if self.lock_popup and self.lock_popup.winfo_exists():
            self._position_lock_popup()
            self.lock_popup.deiconify()
            self.lock_popup.lift()
            return

        popup = tk.Toplevel(self.root)
        self.lock_popup = popup
        popup.title("File In Use")
        popup.resizable(False, False)
        popup.attributes("-topmost", True)
        popup.protocol("WM_DELETE_WINDOW", self._dismiss_lock_popup)

        container = tk.Frame(popup, padx=12, pady=10)
        container.pack(fill="both", expand=True)

        label = tk.Label(
            container,
            textvariable=self.lock_popup_text_var,
            justify="left",
            wraplength=320,
            font=("Segoe UI", 10),
        )
        label.pack(anchor="w")

        actions = tk.Frame(container, pady=8)
        actions.pack(fill="x")
        open_button = tk.Button(actions, text="Open Details", command=self._open_lock_details_from_popup, width=14)
        open_button.pack(side="left")
        dismiss_button = tk.Button(actions, text="Dismiss", command=self._dismiss_lock_popup, width=10)
        dismiss_button.pack(side="right")

        label.bind("<Button-1>", self._open_lock_details_from_popup)

        self._position_lock_popup()
        self.lock_popup_after_id = popup.after(12000, self._dismiss_lock_popup)

    def _clear_lock_tree(self) -> None:
        if not self.lock_tree:
            return
        for item_id in self.lock_tree.get_children():
            self.lock_tree.delete(item_id)

    def _schedule_lock_refresh(self) -> None:
        if not self.lock_window or not self.lock_window.winfo_exists():
            return
        self.lock_refresh_after_id = self.lock_window.after(2000, self._lock_refresh_tick)

    def _lock_refresh_tick(self) -> None:
        self.lock_refresh_after_id = None
        if not self.lock_window or not self.lock_window.winfo_exists():
            return
        self.refresh_lock_window()
        self._schedule_lock_refresh()

    def _open_locked_file_folder(self) -> None:
        if not self.lock_window_file_path:
            return
        folder = os.path.dirname(self.lock_window_file_path)
        if folder and os.path.exists(folder):
            os.startfile(folder)

    def _retry_delete_locked_file(self) -> None:
        file_path = self.lock_window_file_path
        if not file_path:
            return
        if not os.path.exists(file_path):
            self.lock_window_status_var.set("File no longer exists.")
            return
        if list_locking_processes(file_path):
            self.lock_window_status_var.set("File is still locked by one or more processes.")
            return
        try:
            size = os.path.getsize(file_path)
        except OSError:
            size = 0
        try:
            os.remove(file_path)
            self.lock_notified_paths.discard(file_path)
            self.log(f"Deleted after unlock: {file_path} ({human_size(size)})")
            self.lock_window_status_var.set("Deleted successfully after unlock.")
            self.status_var.set("Deleted previously locked file.")
        except OSError as exc:
            self.log(f"Failed deleting unlocked file {file_path}: {exc}")
            self.lock_window_status_var.set(f"Delete failed: {exc}")

    def _terminate_selected_lock_processes(self) -> None:
        if not self.lock_tree:
            return
        selected = self.lock_tree.selection()
        if not selected:
            messagebox.showinfo("Locked File Inspector", "Select one or more processes first.", parent=self.lock_window)
            return

        pids = sorted({int(self.lock_tree.item(item_id, "values")[0]) for item_id in selected})
        if not pids:
            return

        if not messagebox.askyesno(
            "Terminate Processes",
            f"Terminate selected process(es): {', '.join(str(pid) for pid in pids)}?",
            parent=self.lock_window,
        ):
            return

        success_count = 0
        failures = []
        for pid in pids:
            if pid == os.getpid():
                failures.append(f"PID {pid}: skipping monitor app process.")
                continue
            try:
                result = subprocess.run(
                    ["taskkill", "/PID", str(pid), "/F"],
                    capture_output=True,
                    text=True,
                    timeout=8,
                )
                if result.returncode == 0:
                    success_count += 1
                    self.log(f"Terminated process PID {pid} to release file lock.")
                else:
                    error_text = (result.stderr or result.stdout or "Unknown failure").strip()
                    failures.append(f"PID {pid}: {error_text}")
            except Exception as exc:
                failures.append(f"PID {pid}: {exc}")

        if success_count:
            self.lock_window_status_var.set(f"Terminated {success_count} process(es). Refreshing lock state...")
        if failures:
            messagebox.showwarning(
                "Terminate Results",
                "Some processes could not be terminated:\n\n" + "\n".join(failures),
                parent=self.lock_window,
            )

        self.refresh_lock_window()
        self._retry_delete_locked_file()
        self.root.after(100, self.check_once)

    def _close_lock_window(self) -> None:
        if self.lock_refresh_after_id and self.lock_window and self.lock_window.winfo_exists():
            self.lock_window.after_cancel(self.lock_refresh_after_id)
        self.lock_refresh_after_id = None
        if self.lock_window and self.lock_window.winfo_exists():
            self.lock_window.destroy()
        self.lock_window = None
        self.lock_tree = None

    def open_lock_window(self, file_path: str) -> None:
        self.lock_window_file_path = file_path
        self.lock_window_file_var.set(file_path)

        if self.lock_window and self.lock_window.winfo_exists():
            self.lock_window.deiconify()
            self.lock_window.lift()
            self.lock_window.focus_force()
            self.refresh_lock_window()
            return

        window = tk.Toplevel(self.root)
        self.lock_window = window
        window.title("Locked File Inspector")
        window.geometry("860x420")
        window.minsize(760, 320)
        window.protocol("WM_DELETE_WINDOW", self._close_lock_window)

        top = tk.Frame(window, padx=12, pady=10)
        top.pack(fill="x")
        tk.Label(top, text="File:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(top, textvariable=self.lock_window_file_var, anchor="w", justify="left", wraplength=820).pack(
            fill="x", pady=(2, 6)
        )
        tk.Label(top, textvariable=self.lock_window_status_var, anchor="w", font=("Segoe UI", 10)).pack(fill="x")

        table_frame = tk.Frame(window, padx=12, pady=0)
        table_frame.pack(fill="both", expand=True, pady=(2, 8))

        self.lock_tree = ttk.Treeview(
            table_frame,
            columns=("pid", "name", "type", "service"),
            show="headings",
            selectmode="extended",
        )
        self.lock_tree.heading("pid", text="PID")
        self.lock_tree.heading("name", text="Process")
        self.lock_tree.heading("type", text="Type")
        self.lock_tree.heading("service", text="Service")
        self.lock_tree.column("pid", width=90, anchor="center")
        self.lock_tree.column("name", width=300, anchor="w")
        self.lock_tree.column("type", width=140, anchor="center")
        self.lock_tree.column("service", width=220, anchor="w")
        self.lock_tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.lock_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.lock_tree.configure(yscrollcommand=scrollbar.set)

        actions = tk.Frame(window, padx=12, pady=0)
        actions.pack(fill="x", pady=(0, 12))
        tk.Button(actions, text="Refresh Now", command=self.refresh_lock_window, width=14).pack(side="left")
        tk.Button(actions, text="Terminate Selected", command=self._terminate_selected_lock_processes, width=18).pack(
            side="left", padx=(8, 0)
        )
        tk.Button(actions, text="Retry Delete", command=self._retry_delete_locked_file, width=14).pack(
            side="left", padx=(8, 0)
        )
        tk.Button(actions, text="Open Folder", command=self._open_locked_file_folder, width=12).pack(
            side="right", padx=(8, 0)
        )
        tk.Button(actions, text="Close", command=self._close_lock_window, width=10).pack(side="right")

        self.refresh_lock_window()
        self._schedule_lock_refresh()

    def refresh_lock_window(self) -> None:
        if not self.lock_window or not self.lock_window.winfo_exists() or not self.lock_tree:
            return

        selected_pids = set()
        for item_id in self.lock_tree.selection():
            values = self.lock_tree.item(item_id, "values")
            if not values:
                continue
            try:
                selected_pids.add(int(values[0]))
            except (TypeError, ValueError):
                continue

        self._clear_lock_tree()
        file_path = self.lock_window_file_path
        self.lock_window_file_var.set(file_path or "")
        if not file_path:
            self.lock_window_status_var.set("No file selected.")
            return

        if not os.path.exists(file_path):
            self.lock_window_status_var.set("File no longer exists.")
            return

        locks = list_locking_processes(file_path)
        if not locks:
            self.lock_window_status_var.set("No active process lock detected.")
            self.lock_notified_paths.discard(file_path)
            return

        for index, item in enumerate(locks):
            self.lock_tree.insert(
                "",
                "end",
                iid=f"pid-{item['pid']}-{index}",
                values=(item["pid"], item["name"], item["type"], item["service"] or "-"),
            )

        if selected_pids:
            for item_id in self.lock_tree.get_children():
                values = self.lock_tree.item(item_id, "values")
                if not values:
                    continue
                try:
                    pid = int(values[0])
                except (TypeError, ValueError):
                    continue
                if pid in selected_pids:
                    self.lock_tree.selection_add(item_id)

        self.lock_window_status_var.set(f"{len(locks)} process(es) currently locking this file.")

    def _handle_locked_file(self, file_path: str, exc: OSError) -> None:
        self.log(f"Failed deleting {file_path}: {exc}")
        self.status_var.set("Over limit. One or more files are locked by running processes.")
        if file_path not in self.lock_notified_paths:
            self._notify_locked_file(file_path)

    def _schedule_check(self, initial: bool = False) -> None:
        delay = 1000 if initial else self.interval_ms
        self.root.after(delay, self._check_tick)

    def _check_tick(self) -> None:
        if self.is_exiting:
            return
        self.check_once()
        self._schedule_check(initial=False)

    def get_folder_state(self, watch_path: str) -> tuple[int, list[dict]]:
        total_size = 0
        files = []

        for current_root, _, file_names in os.walk(watch_path):
            for name in file_names:
                path = os.path.join(current_root, name)
                try:
                    stats = os.stat(path)
                except OSError:
                    continue
                total_size += stats.st_size
                files.append(
                    {
                        "path": path,
                        "size": stats.st_size,
                        "mtime": stats.st_mtime,
                    }
                )
        return total_size, files

    def check_once(self) -> None:
        monitored_paths = self.config["monitored_paths"]
        if not monitored_paths:
            self.status_var.set("No monitored paths configured. Use View Paths.")
            return

        locked_paths_this_cycle = set()
        deleted_count = 0
        paths_over_limit = 0
        missing_paths = []

        for entry in monitored_paths:
            watch_path = entry["path"]
            if not os.path.exists(watch_path):
                missing_paths.append(watch_path)
                if watch_path not in self.missing_paths_notified:
                    self.log(f"Folder not found: {watch_path}")
                    self.missing_paths_notified.add(watch_path)
                continue

            self.missing_paths_notified.discard(watch_path)
            limit_bytes = parse_size_to_bytes(entry["limit_input"])
            total_size, files = self.get_folder_state(watch_path)
            over_by = total_size - limit_bytes
            if over_by <= 0:
                continue

            paths_over_limit += 1
            reverse = normalize_delete_mode(entry["delete_mode"]) == "latest"
            files.sort(key=lambda item: item["mtime"], reverse=reverse)

            for item in files:
                if total_size <= limit_bytes:
                    break
                try:
                    os.remove(item["path"])
                    total_size -= item["size"]
                    deleted_count += 1
                    self.log(f"Deleted: {item['path']} ({human_size(item['size'])})")
                except OSError as exc:
                    if is_locked_file_error(exc):
                        locked_paths_this_cycle.add(item["path"])
                        self._handle_locked_file(item["path"], exc)
                    else:
                        self.log(f"Failed deleting {item['path']}: {exc}")

        self.lock_notified_paths = locked_paths_this_cycle

        if locked_paths_this_cycle:
            self.status_var.set(
                f"Over limit. {len(locked_paths_this_cycle)} locked file(s) are still in use by another process."
            )
        elif deleted_count:
            self.status_var.set(
                f"Over limit. Deleted {deleted_count} file(s) across monitored paths."
            )
        elif paths_over_limit:
            self.status_var.set("Over limit, but no files could be deleted.")
        elif missing_paths:
            self.status_var.set(f"{len(missing_paths)} monitored path(s) are missing.")
        else:
            self.status_var.set("All monitored paths are within limits.")

    def open_watch_folder(self) -> None:
        if not self.config["monitored_paths"]:
            self.status_var.set("No monitored paths configured.")
            return
        path = self.config["monitored_paths"][0]["path"]
        os.makedirs(path, exist_ok=True)
        os.startfile(path)

    def reconfigure(self) -> None:
        new_config = run_setup_wizard(self.root, self.config)
        if new_config is None:
            return
        self.config = normalize_config(new_config)
        self._update_summary()
        self._sync_startup_setting(show_popup=True)
        save_config(new_config)
        messagebox.showinfo("Setup", "Configuration saved. Restarting now.", parent=self.root)
        restart_self()

    def _on_tray_open(self) -> None:
        self.root.after(0, self.show_window)

    def _on_tray_menu(self, x: int, y: int) -> None:
        self.root.after(0, lambda: self.show_tray_menu(x, y))

    def show_tray_menu(self, x: int, y: int) -> None:
        try:
            self.tray_menu.tk_popup(x, y)
        finally:
            self.tray_menu.grab_release()

    def show_window(self) -> None:
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def hide_to_tray(self) -> None:
        if self.tray_available:
            self.root.withdraw()
            self.status_var.set("Running in tray. Double-click tray icon to open.")
        else:
            self.root.iconify()

    def on_window_close(self) -> None:
        # Closing the window always minimizes to tray when available.
        if self.tray_available:
            self.hide_to_tray()
        else:
            self.exit_app()

    def exit_app(self) -> None:
        self.is_exiting = True
        self._close_paths_window()
        self._dismiss_lock_popup()
        self._close_lock_window()
        if self.tray_available:
            self.tray_icon.stop()
        self.root.destroy()


def main() -> None:
    ensure_gui_only_process()
    ensure_runner_script()

    root = tk.Tk()
    root.withdraw()

    config = load_config()
    if config is None:
        setup_config = run_setup_wizard(root, DEFAULT_CONFIG)
        if setup_config is None:
            root.destroy()
            return
        ok, error = set_startup_enabled(bool(setup_config.get("start_on_startup", False)))
        if not ok:
            messagebox.showwarning(
                "Startup Setting",
                "Could not apply startup preference.\n\n" + error,
                parent=root,
            )
        save_config(setup_config)
        messagebox.showinfo("Setup", "Configuration saved. Restarting now.", parent=root)
        root.destroy()
        restart_self()
        return

    app = MonitorApp(root, config)
    root.mainloop()


if __name__ == "__main__":
    main()
