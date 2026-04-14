"""
Parsec Connection Monitor — GUI Edition
PyQt6 · System Tray · Whitelist · Auto-unlock / Auto-lock
"""

import os, sys, time, json, re, threading, ctypes, ctypes.wintypes
from datetime import datetime
from pathlib import Path

try:
    from PyQt6.QtWidgets import (
        QApplication, QSystemTrayIcon, QMenu, QDialog, QVBoxLayout,
        QHBoxLayout, QLabel, QPushButton, QLineEdit, QListWidget,
        QListWidgetItem, QWidget, QCheckBox, QMessageBox, QFrame,
        QScrollArea, QAbstractItemView
    )
    from PyQt6.QtGui import QIcon, QPixmap, QAction, QColor, QPainter
    from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QByteArray
    from PyQt6.QtSvgWidgets import QSvgWidget
    from PyQt6.QtSvg import QSvgRenderer
except ImportError:
    print("Run:  pip install PyQt6")
    sys.exit(1)

# ── Paths ──────────────────────────────────────────────────────────────────────

BASE_DIR    = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "config.json"

PARSEC_PATHS = {
    "per_computer": Path(os.environ.get("PROGRAMDATA", r"C:\ProgramData")) / "Parsec",
    "per_user":     Path(os.environ.get("APPDATA",     ""))               / "Parsec",
}

LOG_PATTERN = re.compile(
    r"^\[I (?P<ts>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]\s+"
    r"(?P<user>\S+)\s+(?P<event>connected|disconnected)\.$"
)

POLL_INTERVAL = 1.0

# ── SVG Icons ──────────────────────────────────────────────────────────────────

SVG_NORMAL = b"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">
  <defs>
    <radialGradient id="bg" cx="50%" cy="50%" r="50%">
      <stop offset="0%" stop-color="#1a1a2e"/>
      <stop offset="100%" stop-color="#0d0d1a"/>
    </radialGradient>
    <radialGradient id="glow" cx="50%" cy="40%" r="50%">
      <stop offset="0%" stop-color="#00d4ff" stop-opacity="0.3"/>
      <stop offset="100%" stop-color="#00d4ff" stop-opacity="0"/>
    </radialGradient>
  </defs>
  <circle cx="32" cy="32" r="30" fill="url(#bg)" stroke="#00d4ff" stroke-width="1.5" stroke-opacity="0.6"/>
  <circle cx="32" cy="32" r="30" fill="url(#glow)"/>
  <rect x="14" y="16" width="36" height="24" rx="3" fill="none" stroke="#00d4ff" stroke-width="2"/>
  <rect x="16" y="18" width="32" height="20" rx="2" fill="#00d4ff" fill-opacity="0.08"/>
  <polygon points="26,23 26,33 37,28" fill="#00d4ff" fill-opacity="0.9"/>
  <line x1="32" y1="40" x2="32" y2="46" stroke="#00d4ff" stroke-width="2" stroke-opacity="0.7"/>
  <line x1="25" y1="46" x2="39" y2="46" stroke="#00d4ff" stroke-width="2" stroke-opacity="0.7"/>
  <circle cx="20" cy="51" r="2" fill="#00d4ff" fill-opacity="0.5"/>
  <circle cx="32" cy="53" r="2.5" fill="#00ff88"/>
  <circle cx="44" cy="51" r="2" fill="#00d4ff" fill-opacity="0.5"/>
  <line x1="22" y1="51" x2="30" y2="53" stroke="#00d4ff" stroke-width="1" stroke-opacity="0.4"/>
  <line x1="34" y1="53" x2="42" y2="51" stroke="#00d4ff" stroke-width="1" stroke-opacity="0.4"/>
</svg>"""

SVG_LOCKED = b"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">
  <defs>
    <radialGradient id="bg2" cx="50%" cy="50%" r="50%">
      <stop offset="0%" stop-color="#1a1a2e"/>
      <stop offset="100%" stop-color="#0d0d1a"/>
    </radialGradient>
  </defs>
  <circle cx="32" cy="32" r="30" fill="url(#bg2)" stroke="#ff4466" stroke-width="1.5" stroke-opacity="0.7"/>
  <rect x="14" y="14" width="36" height="24" rx="3" fill="none" stroke="#ff4466" stroke-width="2"/>
  <rect x="16" y="16" width="32" height="20" rx="2" fill="#ff4466" fill-opacity="0.07"/>
  <rect x="24" y="26" width="16" height="10" rx="2" fill="#ff4466" fill-opacity="0.85"/>
  <path d="M27 26 v-4 a5 5 0 0 1 10 0 v4" fill="none" stroke="#ff4466" stroke-width="2.2"/>
  <circle cx="32" cy="30" r="2" fill="#0d0d1a"/>
  <line x1="32" y1="32" x2="32" y2="34" stroke="#0d0d1a" stroke-width="1.5"/>
  <line x1="32" y1="38" x2="32" y2="44" stroke="#ff4466" stroke-width="2" stroke-opacity="0.7"/>
  <line x1="24" y1="44" x2="40" y2="44" stroke="#ff4466" stroke-width="2" stroke-opacity="0.7"/>
</svg>"""

def svg_to_icon(svg_bytes: bytes, size: int = 64) -> QIcon:
    renderer = QSvgRenderer(QByteArray(svg_bytes))
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    renderer.render(painter)
    painter.end()
    return QIcon(pixmap)

# ── Config ─────────────────────────────────────────────────────────────────────

DEFAULT_CONFIG = {
    "install_type":  "",
    "parsec_folder": "",
    # Whitelist entries:
    # {
    #   "parsec_user": "Houwy#10157355",
    #   "auto_unlock": true,
    #   "auto_lock":   true
    # }
    "whitelist": []
}

def detect_parsec_folder() -> tuple[str, str]:
    for itype, folder in PARSEC_PATHS.items():
        if (folder / "log.txt").exists():
            return itype, str(folder)
    return "", ""

def load_config() -> dict:
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        for k, v in DEFAULT_CONFIG.items():
            cfg.setdefault(k, v)
    else:
        cfg = dict(DEFAULT_CONFIG)
        itype, folder = detect_parsec_folder()
        cfg["install_type"]  = itype
        cfg["parsec_folder"] = folder

    save_config(cfg)
    return cfg

def save_config(cfg: dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)

# ── Windows session ────────────────────────────────────────────────────────────

# ── Software lock (input blocker) ─────────────────────────────────────────────

class SoftLock:
    """
    Blocks all mouse and keyboard input by installing low-level Windows hooks.
    - Any blocked input attempt is logged silently (no visual feedback to user).
    - Unlock hotkey: Ctrl+Alt+F13 — releases hooks and restores input.
    - Parsec whitelist connect → auto-unlock.
    """
    _instance = None

    def __init__(self):
        self._active    = False
        self._hook_kb   = None
        self._hook_ms   = None
        self._thread    = None
        self._stop_evt  = threading.Event()
        SoftLock._instance = self

    def lock(self):
        if self._active:
            _log("[SOFTLOCK] Already locked")
            return
        _log("[SOFTLOCK] Engaging input block")
        self._active   = True
        self._stop_evt.clear()
        self._thread   = threading.Thread(target=self._hook_thread, daemon=True)
        self._thread.start()

    def unlock(self):
        if not self._active:
            return
        _log("[SOFTLOCK] Releasing input block")
        self._active = False
        self._stop_evt.set()
        # Post WM_QUIT to break GetMessageW loop in hook thread
        if self._thread and self._thread.is_alive():
            ctypes.windll.user32.PostThreadMessageW(
                self._thread.ident, 0x0012, 0, 0  # WM_QUIT
            )

    def is_locked(self):
        return self._active

    def _hook_thread(self):
        """Run Windows message loop with low-level hooks installed."""
        from ctypes import wintypes

        WH_KEYBOARD_LL = 13
        WH_MOUSE_LL    = 14
        HC_ACTION      = 0
        WM_KEYDOWN     = 0x0100
        WM_SYSKEYDOWN  = 0x0104
        VK_MENU        = 0x12
        VK_CONTROL     = 0x11
        VK_F13         = 0x7C

        user32   = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        # Low-level hooks require NULL hmod and thread_id=0 for system-wide
        # Must use WINFUNCTYPE with c_long return and keep references alive
        HOOKPROC = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_int,
                                       wintypes.WPARAM, wintypes.LPARAM)

        class KBDLLHOOKSTRUCT(ctypes.Structure):
            _fields_ = [("vkCode",      wintypes.DWORD),
                        ("scanCode",    wintypes.DWORD),
                        ("flags",       wintypes.DWORD),
                        ("time",        wintypes.DWORD),
                        ("dwExtraInfo", ctypes.c_ulonglong)]

        # Track modifier state manually inside the hook
        # (GetAsyncKeyState is unreliable inside LL hooks)
        pressed = set()

        VK_LCTRL  = 0xA2
        VK_RCTRL  = 0xA3
        VK_LALT   = 0xA4
        VK_RALT   = 0xA5
        VK_LSHIFT = 0xA0
        VK_RSHIFT = 0xA1
        WM_KEYUP      = 0x0101
        WM_SYSKEYUP   = 0x0105

        def kb_proc(nCode, wParam, lParam):
            if nCode == HC_ACTION and self._active:
                kb  = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
                vk  = kb.vkCode

                # Track press/release
                if wParam in (WM_KEYDOWN, WM_SYSKEYDOWN):
                    pressed.add(vk)
                elif wParam in (WM_KEYUP, WM_SYSKEYUP):
                    pressed.discard(vk)

                # Hotkey: (LCtrl or RCtrl) + (LAlt or RAlt) + F13
                ctrl_down = VK_LCTRL in pressed or VK_RCTRL in pressed or VK_CONTROL in pressed
                alt_down  = VK_LALT  in pressed or VK_RALT  in pressed or VK_MENU    in pressed

                if vk == VK_F13 and ctrl_down and alt_down:
                    _log(f"[SOFTLOCK] Hotkey Ctrl+Alt+F13 detected — unlocking (pressed={[hex(v) for v in pressed]})")
                    threading.Thread(target=self.unlock, daemon=True).start()
                    return 1

                if wParam in (WM_KEYDOWN, WM_SYSKEYDOWN):
                    _log(f"[SOFTLOCK] Blocked key vk=0x{vk:02X}")
                return 1
            return user32.CallNextHookEx(None, nCode, wParam, lParam)

        # Throttle mouse log: max 1 entry per 10 seconds
        _ms_last_log = [0.0]

        def ms_proc(nCode, wParam, lParam):
            if nCode == HC_ACTION and self._active:
                now = time.time()
                if wParam in (0x0200, 0x0201, 0x0202, 0x0203, 0x0204, 0x0205, 0x020A):
                    if now - _ms_last_log[0] >= 10.0:
                        _log(f"[SOFTLOCK] Blocked mouse msg=0x{wParam:04X}")
                        _ms_last_log[0] = now
                return 1
            return user32.CallNextHookEx(None, nCode, wParam, lParam)

        # Keep function references alive — GC would break hooks otherwise
        self._kb_func = HOOKPROC(kb_proc)
        self._ms_func = HOOKPROC(ms_proc)

        # Low-level hooks: hmod=NULL, thread_id=0 (system-wide)
        user32.SetWindowsHookExW.restype  = ctypes.c_void_p
        user32.SetWindowsHookExW.argtypes = [
            ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint
        ]
        self._hook_kb = user32.SetWindowsHookExW(WH_KEYBOARD_LL, self._kb_func, None, 0)
        self._hook_ms = user32.SetWindowsHookExW(WH_MOUSE_LL,    self._ms_func, None, 0)
        err_kb = ctypes.GetLastError()
        err_ms = ctypes.GetLastError()
        _log(f"[SOFTLOCK] Hooks: kb={self._hook_kb} (err={err_kb}) ms={self._hook_ms} (err={err_ms})")

        if not self._hook_kb or not self._hook_ms:
            _log("[SOFTLOCK] Hook install failed — aborting")
            self._active = False
            return

        # Message loop — mandatory for LL hooks to receive events
        msg = wintypes.MSG()
        while not self._stop_evt.is_set():
            # GetMessage blocks until a message arrives — perfect for hooks
            ret = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if ret == 0 or ret == -1:
                break
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

        if self._hook_kb:
            user32.UnhookWindowsHookEx(self._hook_kb)
            self._hook_kb = None
        if self._hook_ms:
            user32.UnhookWindowsHookEx(self._hook_ms)
            self._hook_ms = None
        _log("[SOFTLOCK] Hooks removed")


_soft_lock = SoftLock()


def lock_workstation():
    _soft_lock.lock()

LOG_FILE = BASE_DIR / "unlock_debug.log"

def _log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def unlock_workstation():
    _log("[UNLOCK] Releasing soft lock")
    _soft_lock.unlock()


# ── Log monitor thread ─────────────────────────────────────────────────────────

class LogMonitorThread(QThread):
    event_detected = pyqtSignal(str, str, str)
    status_changed = pyqtSignal(str)

    def __init__(self, log_path: Path):
        super().__init__()
        self.log_path = log_path
        self._stop    = threading.Event()

    def stop(self): self._stop.set()

    def run(self):
        last_pos = last_size = 0
        try:
            last_pos = os.path.getsize(self.log_path)
        except FileNotFoundError:
            pass
        last_size = last_pos
        self.status_changed.emit("Monitoring…")

        while not self._stop.is_set():
            time.sleep(POLL_INTERVAL)

            if not self.log_path.exists():
                self.status_changed.emit("log.txt missing…")
                while not self.log_path.exists() and not self._stop.is_set():
                    time.sleep(2)
                last_pos = last_size = 0
                continue

            try:
                current_size = os.path.getsize(self.log_path)
            except FileNotFoundError:
                continue

            if current_size < last_size:
                last_pos = last_size = 0

            if current_size == last_pos:
                last_size = current_size
                continue

            try:
                with open(self.log_path, "r", encoding="utf-8", errors="replace") as f:
                    f.seek(last_pos)
                    new_data = f.read()
                    last_pos = f.tell()
            except OSError:
                continue

            last_size = current_size
            for line in new_data.splitlines():
                m = LOG_PATTERN.match(line.rstrip())
                if m:
                    self.event_detected.emit(m.group("user"), m.group("event"), m.group("ts"))

# ── Styles ─────────────────────────────────────────────────────────────────────

DARK = """
QWidget          { background:#0d0d1a; color:#e0e0e0; font-family:'Segoe UI',Consolas,monospace; font-size:12px; }
QDialog          { background:#0d0d1a; }
QPushButton      { background:#131325; color:#00d4ff; border:1px solid #00d4ff33; border-radius:4px; padding:5px 14px; }
QPushButton:hover{ background:#00d4ff18; border-color:#00d4ff88; }
QPushButton#del  { color:#ff4466; border-color:#ff446633; }
QPushButton#del:hover { background:#ff446618; }
QLineEdit        { background:#0a0a16; color:#e0e0e0; border:1px solid #2a2a3a; border-radius:4px; padding:5px 8px; }
QLineEdit:focus  { border-color:#00d4ff55; }
QListWidget      { background:#0a0a16; border:1px solid #1e1e30; border-radius:4px; }
QListWidget::item:selected { background:#00d4ff18; color:#00d4ff; }
QCheckBox::indicator       { width:13px; height:13px; border:1px solid #333; border-radius:3px; background:#111; }
QCheckBox::indicator:checked { background:#00d4ff; border-color:#00d4ff; }
QScrollBar:vertical { background:#0a0a16; width:6px; border-radius:3px; }
QScrollBar::handle:vertical { background:#1e1e30; border-radius:3px; }
"""

# ── Whitelist dialog ───────────────────────────────────────────────────────────

class WLDialog(QDialog):
    def __init__(self, parent=None, entry: dict = None):
        super().__init__(parent)
        self.setWindowTitle("Whitelist Entry")
        self.setFixedWidth(400)
        self.setStyleSheet(DARK)
        self.result_data = None

        lay = QVBoxLayout(self)
        lay.setSpacing(10)
        lay.setContentsMargins(18, 18, 18, 18)

        def field(label, placeholder):
            row = QHBoxLayout()
            lbl = QLabel(label); lbl.setFixedWidth(120); lbl.setStyleSheet("color:#666;font-size:11px;")
            edit = QLineEdit(); edit.setPlaceholderText(placeholder)
            row.addWidget(lbl); row.addWidget(edit)
            lay.addLayout(row)
            return edit

        self.f_parsec = field("Parsec user:", "e.g. Houwy#10157355")

        self.cb_unlock = QCheckBox("Auto-unlock on connect");  self.cb_unlock.setChecked(True)
        self.cb_lock   = QCheckBox("Auto-lock on disconnect"); self.cb_lock.setChecked(True)
        lay.addWidget(self.cb_unlock)
        lay.addWidget(self.cb_lock)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine); sep.setStyleSheet("color:#1e1e30;")
        lay.addWidget(sep)

        btns = QHBoxLayout()
        ok = QPushButton("Save"); cancel = QPushButton("Cancel")
        ok.clicked.connect(self._ok); cancel.clicked.connect(self.reject)
        btns.addStretch(); btns.addWidget(cancel); btns.addWidget(ok)
        lay.addLayout(btns)

        if entry:
            self.f_parsec.setText(entry.get("parsec_user",""))
            self.cb_unlock.setChecked(entry.get("auto_unlock", True))
            self.cb_lock.setChecked(entry.get("auto_lock", True))

    def _ok(self):
        if not self.f_parsec.text().strip():
            QMessageBox.warning(self, "Error", "Parsec username is required."); return
        self.result_data = {
            "parsec_user": self.f_parsec.text().strip(),
            "auto_unlock": self.cb_unlock.isChecked(),
            "auto_lock":   self.cb_lock.isChecked(),
        }
        self.accept()

# ── Event row ──────────────────────────────────────────────────────────────────

class EventRow(QWidget):
    def __init__(self, user, event, ts, wl):
        super().__init__()
        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 5, 12, 5)
        color = "#00ff88" if event == "connected" else "#ff4466"

        dot = QLabel("●"); dot.setFixedWidth(16); dot.setStyleSheet(f"color:{color};font-size:13px;")
        name = QLabel(user); name.setStyleSheet("color:#ddd;font-weight:600;"); name.setFixedWidth(170)
        act  = QLabel(event); act.setStyleSheet(f"color:{color};font-size:11px;"); act.setFixedWidth(95)
        ts_l = QLabel(ts);    ts_l.setStyleSheet("color:#555;font-size:10px;")
        wl_l = QLabel("★" if wl else "");  wl_l.setStyleSheet("color:#f0a500;"); wl_l.setFixedWidth(18)

        for w in (dot, name, act, ts_l): lay.addWidget(w)
        lay.addStretch(); lay.addWidget(wl_l)
        self.setStyleSheet("QWidget{background:transparent;border-bottom:1px solid #141420;}"
                           "QWidget:hover{background:rgba(0,212,255,0.03);}")

# ── Main window ────────────────────────────────────────────────────────────────

class MainWindow(QDialog):
    def __init__(self, config):
        super().__init__()
        self.config = config
        self.setWindowTitle("Parsec Monitor")
        self.setMinimumSize(660, 500)
        self.setStyleSheet(DARK)
        self.setWindowIcon(svg_to_icon(SVG_NORMAL))

        root = QVBoxLayout(self); root.setContentsMargins(0,0,0,0); root.setSpacing(0)

        # Header
        hdr = QWidget(); hdr.setStyleSheet("background:#080812;border-bottom:1px solid #1a1a28;")
        hl  = QHBoxLayout(hdr); hl.setContentsMargins(16,12,16,12)
        ico = QSvgWidget(); ico.load(QByteArray(SVG_NORMAL)); ico.setFixedSize(38,38)
        tc  = QVBoxLayout(); tc.setSpacing(1)
        tc.addWidget(QLabel("PARSEC MONITOR", styleSheet="color:#00d4ff;font-size:17px;font-weight:700;letter-spacing:2px;"))
        tc.addWidget(QLabel("connection watchdog", styleSheet="color:#333;font-size:10px;letter-spacing:1px;"))
        self.s_dot  = QLabel("●", styleSheet="color:#00ff88;font-size:18px;")
        self.s_text = QLabel("Monitoring", styleSheet="color:#00ff88;font-size:10px;")
        sc = QVBoxLayout(); sc.setAlignment(Qt.AlignmentFlag.AlignRight)
        sc.addWidget(self.s_dot, alignment=Qt.AlignmentFlag.AlignRight)
        sc.addWidget(self.s_text, alignment=Qt.AlignmentFlag.AlignRight)
        hl.addWidget(ico); hl.addSpacing(10); hl.addLayout(tc); hl.addStretch(); hl.addLayout(sc)
        root.addWidget(hdr)

        # Body
        body = QWidget(); bl = QHBoxLayout(body); bl.setContentsMargins(0,0,0,0); bl.setSpacing(0)

        # Left — log
        left = QWidget(); left.setStyleSheet("border-right:1px solid #1a1a28;")
        ll   = QVBoxLayout(left); ll.setContentsMargins(0,0,0,0); ll.setSpacing(0)
        ll.addWidget(QLabel("  EVENT LOG", styleSheet="color:#333;font-size:10px;letter-spacing:2px;padding:7px 12px;background:#080812;border-bottom:1px solid #1a1a28;"))
        self.scroll = QScrollArea(); self.scroll.setWidgetResizable(True)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.ev_cont = QWidget(); self.ev_lay = QVBoxLayout(self.ev_cont)
        self.ev_lay.setContentsMargins(0,0,0,0); self.ev_lay.setSpacing(0); self.ev_lay.addStretch()
        self.scroll.setWidget(self.ev_cont)
        ll.addWidget(self.scroll)

        # Right — whitelist
        right = QWidget(); right.setFixedWidth(210)
        rl = QVBoxLayout(right); rl.setContentsMargins(0,0,0,0); rl.setSpacing(0)
        rl.addWidget(QLabel("  WHITELIST", styleSheet="color:#333;font-size:10px;letter-spacing:2px;padding:7px 12px;background:#080812;border-bottom:1px solid #1a1a28;"))
        self.wl_list = QListWidget()
        self.wl_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        rl.addWidget(self.wl_list)
        wb = QWidget(); wb.setStyleSheet("background:#080812;border-top:1px solid #1a1a28;")
        wbl = QHBoxLayout(wb); wbl.setContentsMargins(8,5,8,5)
        for txt, slot in [("+ Add", self._wl_add), ("Edit", self._wl_edit)]:
            b = QPushButton(txt); b.clicked.connect(slot); wbl.addWidget(b)
        d = QPushButton("✕"); d.setObjectName("del"); d.clicked.connect(self._wl_del); wbl.addWidget(d)
        rl.addWidget(wb)

        bl.addWidget(left, stretch=1); bl.addWidget(right)
        root.addWidget(body, stretch=1)

        # Footer
        ft = QWidget(); ft.setStyleSheet("background:#08081200;border-top:1px solid #1a1a28;")
        fl = QHBoxLayout(ft); fl.setContentsMargins(14,7,14,7)
        fl.addWidget(QLabel(f"📂 {config.get('parsec_folder','?')}", styleSheet="color:#333;font-size:10px;"))
        fl.addStretch()
        hb = QPushButton("Hide to Tray"); hb.clicked.connect(self.hide); fl.addWidget(hb)
        root.addWidget(ft)

        self._refresh_wl()

    # Whitelist
    def _refresh_wl(self):
        self.wl_list.clear()
        for e in self.config.get("whitelist", []):
            flags = ("🔓" if e.get("auto_unlock") else "") + ("🔒" if e.get("auto_lock") else "")
            item = QListWidgetItem(f"  {e['parsec_user']}  {flags}")
            item.setForeground(QColor("#00d4ff"))
            self.wl_list.addItem(item)

    def _wl_add(self):
        dlg = WLDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted and dlg.result_data:
            self.config["whitelist"].append(dlg.result_data)
            save_config(self.config); self._refresh_wl()

    def _wl_edit(self):
        row = self.wl_list.currentRow()
        if row < 0: return
        dlg = WLDialog(self, self.config["whitelist"][row])
        if dlg.exec() == QDialog.DialogCode.Accepted and dlg.result_data:
            self.config["whitelist"][row] = dlg.result_data
            save_config(self.config); self._refresh_wl()

    def _wl_del(self):
        row = self.wl_list.currentRow()
        if row < 0: return
        e = self.config["whitelist"][row]
        if QMessageBox.question(self, "Delete", f"Remove {e['parsec_user']}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        ) == QMessageBox.StandardButton.Yes:
            self.config["whitelist"].pop(row); save_config(self.config); self._refresh_wl()

    # Events
    def add_event(self, user, event, ts, wl):
        row = EventRow(user, event, ts, wl)
        self.ev_lay.insertWidget(self.ev_lay.count()-1, row)
        QTimer.singleShot(50, lambda: self.scroll.verticalScrollBar().setValue(
            self.scroll.verticalScrollBar().maximum()))

    def set_status(self, text, ok=True):
        c = "#00ff88" if ok else "#ff4466"
        self.s_dot.setStyleSheet(f"color:{c};font-size:18px;")
        self.s_text.setStyleSheet(f"color:{c};font-size:10px;")
        self.s_text.setText(text)

    def closeEvent(self, e): e.ignore(); self.hide()

# ── App ────────────────────────────────────────────────────────────────────────

class App:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.app.setQuitOnLastWindowClosed(False)
        self.config = load_config()
        self._ensure_parsec_folder()

        self.ico_ok  = svg_to_icon(SVG_NORMAL)
        self.ico_lck = svg_to_icon(SVG_LOCKED)

        self.win = MainWindow(self.config)

        # Tray
        self.tray = QSystemTrayIcon(self.ico_ok, self.app)
        self.tray.setToolTip("Parsec Monitor")
        m = QMenu()
        m.addAction(QAction("Show", self.app, triggered=self._show))
        m.addSeparator()
        m.addAction(QAction("Quit", self.app, triggered=self._quit))
        self.tray.setContextMenu(m)
        self.tray.activated.connect(lambda r: self._show() if r == QSystemTrayIcon.ActivationReason.DoubleClick else None)
        self.tray.show()

        # Monitor
        log_path = Path(self.config["parsec_folder"]) / "log.txt"
        self.mon = LogMonitorThread(log_path)
        self.mon.event_detected.connect(self._on_event)
        self.mon.status_changed.connect(lambda t: self.win.set_status(t))
        self.mon.start()

    def _ensure_parsec_folder(self):
        if not self.config["parsec_folder"]:
            itype, folder = detect_parsec_folder()
            if folder:
                self.config["install_type"]  = itype
                self.config["parsec_folder"] = folder
                save_config(self.config)
            else:
                QMessageBox.warning(None, "Parsec Monitor",
                    "Parsec log.txt not found.\nSet 'parsec_folder' in config.json manually.")

    def _show(self):
        self.win.show(); self.win.raise_(); self.win.activateWindow()

    def _quit(self):
        self.mon.stop(); self.mon.wait(2000); self.app.quit()

    def _on_event(self, user: str, event: str, ts: str):
        wl = next((e for e in self.config.get("whitelist",[])
                   if e.get("parsec_user","").lower() == user.lower()), None)

        self.win.add_event(user, event, ts, wl is not None)
        self.tray.showMessage("Parsec Monitor",
            f"{'🟢' if event=='connected' else '🔴'} {user} {event}",
            QSystemTrayIcon.MessageIcon.Information, 3000)

        if wl:
            if event == "connected":
                self.tray.setIcon(self.ico_ok)
                if wl.get("auto_unlock"):
                    threading.Thread(target=unlock_workstation, daemon=True).start()
            else:
                self.tray.setIcon(self.ico_lck)
                if wl.get("auto_lock"):
                    threading.Thread(target=lock_workstation, daemon=True).start()
        else:
            self.tray.setIcon(self.ico_ok if event == "connected" else self.ico_lck)

    def run(self): sys.exit(self.app.exec())

# ── Entry ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if sys.platform != "win32":
        print("Windows only.")
        input("Press Enter to exit...")
        sys.exit(1)
    try:
        App().run()
    except Exception:
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")