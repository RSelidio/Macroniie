import json
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox

import psutil
from pynput import keyboard, mouse

try:
    import ctypes
    from ctypes import wintypes
except Exception as exc:
    raise RuntimeError("This app requires Windows.") from exc


MOVE_SAMPLE_SEC = 0.02
SM_XVIRTUALSCREEN = 76
SM_YVIRTUALSCREEN = 77
SM_CXVIRTUALSCREEN = 78
SM_CYVIRTUALSCREEN = 79

MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
MOUSEEVENTF_MIDDLEDOWN = 0x0020
MOUSEEVENTF_MIDDLEUP = 0x0040
MOUSEEVENTF_WHEEL = 0x0800
MOUSEEVENTF_ABSOLUTE = 0x8000
MOUSEEVENTF_VIRTUALDESK = 0x4000


def set_process_dpi_awareness():
    try:
        ctypes.WinDLL("shcore", use_last_error=True).SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.WinDLL("user32", use_last_error=True).SetProcessDPIAware()
    except Exception:
        pass


def get_screen_size():
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)


def get_virtual_screen_rect():
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    left = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
    top = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
    width = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
    height = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
    return left, top, width, height


def get_cursor_pos():
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    pt = wintypes.POINT()
    if not user32.GetCursorPos(ctypes.byref(pt)):
        return 0, 0
    return pt.x, pt.y


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", getattr(wintypes, "ULONG_PTR", ctypes.c_size_t)),
    ]


class INPUT(ctypes.Structure):
    _fields_ = [
        ("type", wintypes.DWORD),
        ("mi", MOUSEINPUT),
    ]


set_process_dpi_awareness()


def get_foreground_process_and_title():
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return None, ""

    length = user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)

    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if not pid.value:
        return None, buf.value

    try:
        process = psutil.Process(pid.value)
        return process, buf.value
    except psutil.Error:
        return None, buf.value


class MacroApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Macronnie")

        self.target_title_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Idle")
        self.event_count_var = tk.StringVar(value="0")
        self.duration_var = tk.StringVar(value="0.00s")
        self.loop_var = tk.BooleanVar(value=False)
        self.loop_interval_var = tk.StringVar(value="0.0")
        self.loop_count_var = tk.StringVar(value="1")
        self.absolute_mouse_var = tk.BooleanVar(value=True)
        self.hotkey_record_var = tk.StringVar(value="ctrl+1")
        self.hotkey_stop_var = tk.StringVar(value="ctrl+2")
        self.hotkey_play_var = tk.StringVar(value="ctrl+3")
        self.hotkey_kill_var = tk.StringVar(value="ctrl+esc")
        self.hotkey_info_var = tk.StringVar(value="")

        self.bg_key_enabled_vars = [tk.BooleanVar(value=False) for _ in range(10)]
        self.bg_key_name_vars = [tk.StringVar(value="f" + str(i+1)) for i in range(10)]
        self.bg_key_interval_vars = [tk.StringVar(value="10.0") for i in range(10)]
        self.bg_key_last_press = [0.0] * 10

        self.events = []
        self.recording = False
        self.playing = False
        self.record_start = 0.0
        self.last_move_time = 0.0
        self.stop_event = threading.Event()

        self.mouse_listener = None
        self.keyboard_listener = None
        self.hotkey_listener = None
        self.play_thread = None
        self.hotkey_bindings = {}
        self.pressed_tokens = set()
        self.active_hotkeys = set()
        self.loop_wait_seconds = 0.0
        self.loop_iterations = 0
        self.loop_limit = 1
        self.bg_keys_thread = None
        self.kill_hotkey_binding = frozenset()

        self._build_ui()
        self.apply_hotkeys(show_message=False)
        self._start_hotkey_listener()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self):
        frame = tk.Frame(self.root, padx=12, pady=12)
        frame.pack(fill=tk.BOTH, expand=True)

        info_label = tk.Label(
            frame,
            textvariable=self.hotkey_info_var,
            bg="lightblue",
            fg="black",
            padx=8,
            pady=4,
        )
        info_label.grid(row=0, column=0, columnspan=2, sticky="we", pady=(0, 8))

        tk.Label(frame, text="Window title contains").grid(row=1, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.target_title_var, width=28).grid(
            row=1, column=1, sticky="we", padx=(8, 0)
        )

        button_row = tk.Frame(frame)
        button_row.grid(row=2, column=0, columnspan=2, pady=(12, 0), sticky="we")

        tk.Button(button_row, text="Start Record", command=self.start_record).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        tk.Button(button_row, text="Stop Record", command=self.stop_record).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        tk.Button(button_row, text="Play", command=self.start_playback).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        tk.Button(button_row, text="Stop", command=self.stop_all).pack(
            side=tk.LEFT
        )

        loop_row = tk.Frame(frame)
        loop_row.grid(row=3, column=0, columnspan=2, pady=(6, 0), sticky="w")
        tk.Checkbutton(loop_row, text="Loop playback", variable=self.loop_var).pack(
            side=tk.LEFT
        )
        tk.Label(loop_row, text="Loop interval (s)").pack(side=tk.LEFT, padx=(10, 4))
        tk.Entry(loop_row, textvariable=self.loop_interval_var, width=8).pack(side=tk.LEFT)
        tk.Label(loop_row, text="Times").pack(side=tk.LEFT, padx=(10, 4))
        tk.Entry(loop_row, textvariable=self.loop_count_var, width=6).pack(side=tk.LEFT)
        tk.Checkbutton(
            loop_row,
            text="High precision mouse",
            variable=self.absolute_mouse_var,
        ).pack(side=tk.LEFT, padx=(10, 0))

        hotkey_row = tk.Frame(frame)
        hotkey_row.grid(row=4, column=0, columnspan=2, pady=(10, 0), sticky="we")
        tk.Label(hotkey_row, text="Record").grid(row=0, column=0, sticky="w")
        tk.Entry(hotkey_row, textvariable=self.hotkey_record_var, width=12).grid(row=0, column=1, padx=(6, 12))
        tk.Label(hotkey_row, text="Stop").grid(row=0, column=2, sticky="w")
        tk.Entry(hotkey_row, textvariable=self.hotkey_stop_var, width=12).grid(row=0, column=3, padx=(6, 12))
        tk.Label(hotkey_row, text="Play").grid(row=0, column=4, sticky="w")
        tk.Entry(hotkey_row, textvariable=self.hotkey_play_var, width=12).grid(row=0, column=5, padx=(6, 12))
        tk.Button(hotkey_row, text="Apply Hotkeys", command=self.apply_hotkeys).grid(row=0, column=6)

        hotkey_kill_row = tk.Frame(frame)
        hotkey_kill_row.grid(row=5, column=0, columnspan=2, pady=(4, 0), sticky="w")
        tk.Label(hotkey_kill_row, text="Kill (Stop everything):").pack(side=tk.LEFT, padx=(0, 6))
        tk.Entry(hotkey_kill_row, textvariable=self.hotkey_kill_var, width=12).pack(side=tk.LEFT)

        bg_keys_label = tk.Label(frame, text="Background Keys (press at intervals during playback):", fg="darkgreen", font=("Arial", 9, "bold"))
        bg_keys_label.grid(row=6, column=0, columnspan=2, sticky="w", pady=(10, 4))

        bg_keys_frame = tk.Frame(frame)
        bg_keys_frame.grid(row=7, column=0, columnspan=2, sticky="we", pady=(0, 10))

        for idx in range(10):
            row_idx = idx // 5
            col_idx = idx % 5

            cell_frame = tk.Frame(bg_keys_frame)
            cell_frame.grid(row=row_idx, column=col_idx, padx=4, pady=2, sticky="w")

            tk.Checkbutton(cell_frame, variable=self.bg_key_enabled_vars[idx]).pack(side=tk.LEFT)
            tk.Label(cell_frame, text="Key:").pack(side=tk.LEFT, padx=(2, 2))
            tk.Entry(cell_frame, textvariable=self.bg_key_name_vars[idx], width=5).pack(side=tk.LEFT)
            tk.Label(cell_frame, text="s:").pack(side=tk.LEFT, padx=(4, 2))
            tk.Entry(cell_frame, textvariable=self.bg_key_interval_vars[idx], width=6).pack(side=tk.LEFT)

        file_row = tk.Frame(frame)
        file_row.grid(row=8, column=0, columnspan=2, pady=(10, 0), sticky="we")
        tk.Button(file_row, text="Load", command=self.load_macro).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(file_row, text="Save", command=self.save_macro).pack(side=tk.LEFT)

        info = tk.Frame(frame)
        info.grid(row=9, column=0, columnspan=2, pady=(12, 0), sticky="we")
        tk.Label(info, text="Events:").pack(side=tk.LEFT)
        tk.Label(info, textvariable=self.event_count_var, width=8).pack(side=tk.LEFT)
        tk.Label(info, text="Duration:").pack(side=tk.LEFT, padx=(12, 0))
        tk.Label(info, textvariable=self.duration_var, width=10).pack(side=tk.LEFT)

        status_row = tk.Frame(frame)
        status_row.grid(row=10, column=0, columnspan=2, pady=(10, 0), sticky="we")
        tk.Label(status_row, text="Status:").pack(side=tk.LEFT)
        tk.Label(status_row, textvariable=self.status_var).pack(side=tk.LEFT, padx=(8, 0))

        frame.columnconfigure(1, weight=1)

    def _start_hotkey_listener(self):
        def on_press(key):
            token = self._token_from_key(key)
            if not token:
                return
            self.pressed_tokens.add(token)

            if self._is_hotkey_pressed(self.kill_hotkey_binding, "kill"):
                self.root.after(0, self.stop_all)

            if self._is_hotkey_pressed(self.hotkey_bindings.get("record"), "record"):
                self.root.after(0, self.start_record)
            if self._is_hotkey_pressed(self.hotkey_bindings.get("stop"), "stop"):
                self.root.after(0, self.stop_record)
            if self._is_hotkey_pressed(self.hotkey_bindings.get("play"), "play"):
                self.root.after(0, self.start_playback)

        def on_release(key):
            token = self._token_from_key(key)
            if not token:
                return
            self.pressed_tokens.discard(token)
            self._refresh_active_hotkeys()

        self.hotkey_listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        self.hotkey_listener.daemon = True
        self.hotkey_listener.start()

    def _token_from_key(self, key):
        if key in (keyboard.Key.ctrl_l, keyboard.Key.ctrl_r):
            return "ctrl"
        if key in (keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r):
            return "shift"
        if key in (keyboard.Key.alt_l, keyboard.Key.alt_r, keyboard.Key.alt_gr):
            return "alt"
        if isinstance(key, keyboard.KeyCode):
            if key.char:
                return key.char.lower()
            vk = getattr(key, "vk", None)
            if vk is not None and 32 <= vk <= 126:
                return chr(vk).lower()
            return None
        return getattr(key, "name", None)

    def _normalize_hotkey_text(self, value):
        alias = {
            "control": "ctrl",
            "ctl": "ctrl",
            "return": "enter",
            "escape": "esc",
        }
        parts = [part.strip().lower() for part in value.split("+") if part.strip()]
        if not parts:
            return None
        normalized = []
        for part in parts:
            normalized.append(alias.get(part, part))
        return frozenset(normalized)

    def _format_hotkey(self, combo):
        if not combo:
            return ""
        order = ["ctrl", "shift", "alt"]
        modifiers = [name for name in order if name in combo]
        others = sorted([name for name in combo if name not in order])
        parts = modifiers + others
        return "+".join(part.upper() if len(part) == 1 else part.title() for part in parts)

    def _refresh_hotkey_info(self):
        text = (
            f"Hotkeys: {self._format_hotkey(self.hotkey_bindings.get('record'))}=Record | "
            f"{self._format_hotkey(self.hotkey_bindings.get('stop'))}=Stop | "
            f"{self._format_hotkey(self.hotkey_bindings.get('play'))}=Play | {self._format_hotkey(self.kill_hotkey_binding)}=Kill"
        )
        self.hotkey_info_var.set(text)

    def _is_hotkey_pressed(self, combo, action_name):
        if not combo:
            return False
        active_key = ":" + action_name
        if combo.issubset(self.pressed_tokens):
            if active_key in self.active_hotkeys:
                return False
            self.active_hotkeys.add(active_key)
            return True
        return False

    def _refresh_active_hotkeys(self):
        keep = set()
        if self.kill_hotkey_binding and self.kill_hotkey_binding.issubset(self.pressed_tokens):
            keep.add(":kill")
        for action_name, combo in self.hotkey_bindings.items():
            if combo and combo.issubset(self.pressed_tokens):
                keep.add(":" + action_name)
        self.active_hotkeys = keep

    def apply_hotkeys(self, show_message=True):
        record_combo = self._normalize_hotkey_text(self.hotkey_record_var.get())
        stop_combo = self._normalize_hotkey_text(self.hotkey_stop_var.get())
        play_combo = self._normalize_hotkey_text(self.hotkey_play_var.get())
        kill_combo = self._normalize_hotkey_text(self.hotkey_kill_var.get())

        if not record_combo or not stop_combo or not play_combo or not kill_combo:
            messagebox.showerror("Invalid hotkey", "Record, Stop, Play, and Kill hotkeys are required.")
            return

        combos = {
            "record": record_combo,
            "stop": stop_combo,
            "play": play_combo,
            "kill": kill_combo,
        }

        if len(set(combos.values())) != len(combos):
            messagebox.showerror("Invalid hotkey", "All hotkeys must be unique.")
            return

        self.hotkey_bindings = {"record": record_combo, "stop": stop_combo, "play": play_combo}
        self.kill_hotkey_binding = kill_combo
        self.pressed_tokens.clear()
        self.active_hotkeys.clear()
        self._refresh_hotkey_info()

        if show_message:
            self.status_var.set("Hotkeys updated")

    def is_target_active(self):
        target_title = self.target_title_var.get().strip().lower()

        # No filter set → record/play everywhere
        if not target_title:
            return True

        process, title = get_foreground_process_and_title()
        if not process:
            return False

        title_match = True
        if target_title:
            title_match = target_title in title.lower()

        return title_match

    def start_record(self):
        if self.recording:
            return
        if self.playing:
            messagebox.showwarning("Busy", "Stop playback before recording.")
            return
        self.events = []
        self.event_count_var.set("0")
        self.duration_var.set("0.00s")
        self.record_start = time.monotonic()
        self.last_move_time = 0.0
        self.recording = True
        stop_hint = self._format_hotkey(self.hotkey_bindings.get("stop"))
        self.status_var.set(f"Recording ({stop_hint} or Ctrl+Esc to stop)")

        start_x, start_y = get_cursor_pos()
        self.events.append({
            "t": 0.0,
            "type": "move",
            "x": start_x,
            "y": start_y,
            "coord": "abs",
        })
        self._update_stats()

        self.mouse_listener = mouse.Listener(
            on_move=self.on_move,
            on_click=self.on_click,
            on_scroll=self.on_scroll,
        )
        self.mouse_listener.daemon = True
        self.mouse_listener.start()

        self.keyboard_listener = keyboard.Listener(
            on_press=self.on_key_press,
            on_release=self.on_key_release,
        )
        self.keyboard_listener.daemon = True
        self.keyboard_listener.start()

    def stop_record(self):
        if not self.recording:
            return
        self.recording = False
        if self.mouse_listener:
            self.mouse_listener.stop()
            self.mouse_listener = None
        if self.keyboard_listener:
            self.keyboard_listener.stop()
            self.keyboard_listener = None
        self._update_stats()
        self.status_var.set("Idle")

    def start_playback(self):
        if self.playing:
            return
        if self.recording:
            messagebox.showwarning("Busy", "Stop recording before playback.")
            return
        if not self.events:
            messagebox.showinfo("Empty", "No macro recorded or loaded.")
            return

        loop_interval_seconds = self._parse_loop_interval_seconds()
        if loop_interval_seconds is None:
            return

        loop_count = self._parse_loop_count()
        if loop_count is None:
            return

        self.playing = True
        self.stop_event.clear()
        self.loop_wait_seconds = loop_interval_seconds
        self.loop_iterations = 0
        self.loop_limit = loop_count
        status = "Playing (Ctrl+Esc to stop)"
        if self.loop_var.get():
            if self.loop_limit > 1:
                status = f"Playing (Loop {self.loop_limit}x every {self.loop_wait_seconds:.2f}s, Ctrl+Esc to stop)"
            else:
                status = f"Playing (1 run, Ctrl+Esc to stop)"
        self.status_var.set(status)

        self.play_thread = threading.Thread(target=self._playback_worker, daemon=True)
        self.play_thread.start()

        self.bg_key_last_press = [time.monotonic()] * 10
        self.bg_keys_thread = threading.Thread(target=self._bg_keys_worker, daemon=True)
        self.bg_keys_thread.start()

    def stop_all(self):
        if self.recording:
            self.stop_record()
        if self.playing:
            self.stop_event.set()
            self.playing = False
            self.status_var.set("Idle")

    def on_move(self, x, y):
        if not self.recording or not self.is_target_active():
            return
        now = time.monotonic()
        if now - self.last_move_time < MOVE_SAMPLE_SEC:
            return
        self.last_move_time = now
        abs_x, abs_y = get_cursor_pos()
        self.events.append({
            "t": now - self.record_start,
            "type": "move",
            "x": abs_x,
            "y": abs_y,
            "coord": "abs",
        })
        self._update_stats()

    def on_click(self, x, y, button, pressed):
        if not self.recording or not self.is_target_active():
            return
        now = time.monotonic()
        abs_x, abs_y = get_cursor_pos()
        self.events.append({
            "t": now - self.record_start,
            "type": "click",
            "x": abs_x,
            "y": abs_y,
            "coord": "abs",
            "button": str(button),
            "pressed": pressed,
        })
        self._update_stats()

    def on_scroll(self, x, y, dx, dy):
        if not self.recording or not self.is_target_active():
            return
        now = time.monotonic()
        self.events.append({
            "t": now - self.record_start,
            "type": "scroll",
            "x": x,
            "y": y,
            "dx": dx,
            "dy": dy,
        })
        self._update_stats()

    def on_key_press(self, key):
        self._record_key_event(key, True)

    def on_key_release(self, key):
        self._record_key_event(key, False)

    def _record_key_event(self, key, pressed):
        if not self.recording or not self.is_target_active():
            return
        now = time.monotonic()
        if isinstance(key, keyboard.KeyCode) and key.char:
            entry = {
                "t": now - self.record_start,
                "type": "key",
                "is_char": True,
                "key": key.char,
                "pressed": pressed,
            }
        else:
            name = getattr(key, "name", None)
            if not name:
                return
            entry = {
                "t": now - self.record_start,
                "type": "key",
                "is_char": False,
                "key": name,
                "pressed": pressed,
            }
        self.events.append(entry)
        self._update_stats()

    def _playback_worker(self):
        controller = mouse.Controller()
        keyboard_ctrl = keyboard.Controller()

        while True:
            self.loop_iterations += 1
            base_time = time.monotonic()
            for event in self.events:
                if self.stop_event.is_set():
                    break

                if event.get("t") == 0.0:
                    continue

                target_time = base_time + event["t"]
                wait = target_time - time.monotonic()
                if wait > 0:
                    time.sleep(wait)

                try:
                    if event["type"] == "move":
                        point = self._resolve_point(event)
                        controller.position = point
                    elif event["type"] == "click":
                        point = self._resolve_point(event)
                        controller.position = point
                        time.sleep(0.02)
                        button = getattr(mouse.Button, event["button"].split(".")[-1], mouse.Button.left)
                        if event["pressed"]:
                            controller.press(button)
                        else:
                            controller.release(button)
                    elif event["type"] == "scroll":
                        controller.scroll(event["dx"], event["dy"])
                    elif event["type"] == "key":
                        if event.get("is_char"):
                            key_obj = keyboard.KeyCode.from_char(event.get("key", ""))
                        else:
                            key_name = event.get("key")
                            key_obj = getattr(keyboard.Key, key_name, None)
                        if key_obj:
                            if event.get("pressed"):
                                keyboard_ctrl.press(key_obj)
                            else:
                                keyboard_ctrl.release(key_obj)
                except Exception:
                    continue

            if self.stop_event.is_set() or not self.loop_var.get():
                break

            if self.loop_iterations >= self.loop_limit:
                break

            pause_until = time.monotonic() + self.loop_wait_seconds
            while time.monotonic() < pause_until:
                if self.stop_event.is_set():
                    break
                time.sleep(0.05)

            if self.stop_event.is_set():
                break

        self.playing = False
        self.status_var.set("Idle")

    def _parse_loop_interval_seconds(self):
        raw = self.loop_interval_var.get().strip()
        if not raw:
            raw = "0"
            self.loop_interval_var.set(raw)
        try:
            value = float(raw)
        except ValueError:
            messagebox.showerror("Invalid loop interval", "Loop interval must be a number (seconds).")
            return None
        if value < 0:
            messagebox.showerror("Invalid loop interval", "Loop interval cannot be negative.")
            return None
        return value

    def _parse_loop_count(self):
        raw = self.loop_count_var.get().strip()
        if not raw:
            raw = "1"
            self.loop_count_var.set(raw)
        try:
            value = int(raw)
        except ValueError:
            messagebox.showerror("Invalid loop count", "Loop count must be a whole number.")
            return None
        if value < 1:
            messagebox.showerror("Invalid loop count", "Loop count must be at least 1.")
            return None
        return value

    def _bg_keys_worker(self):
        keyboard_ctrl = keyboard.Controller()
        while self.playing:
            for idx in range(10):
                if not self.bg_key_enabled_vars[idx].get():
                    continue

                key_name = self.bg_key_name_vars[idx].get().strip().lower()
                if not key_name:
                    continue

                try:
                    interval = float(self.bg_key_interval_vars[idx].get().strip())
                except ValueError:
                    continue

                if interval <= 0:
                    continue

                now = time.monotonic()
                if now - self.bg_key_last_press[idx] >= interval:
                    try:
                        key_obj = getattr(keyboard.Key, key_name, None)
                        if not key_obj and len(key_name) == 1:
                            key_obj = keyboard.KeyCode.from_char(key_name)
                        if key_obj:
                            keyboard_ctrl.press(key_obj)
                            time.sleep(0.05)
                            keyboard_ctrl.release(key_obj)
                            self.bg_key_last_press[idx] = now
                    except Exception:
                        pass
            time.sleep(0.1)

    def _normalize_point(self, x, y):
        _, _, width, height = get_virtual_screen_rect()
        if width <= 0 or height <= 0:
            return float(x), float(y)
        return float(x) / width, float(y) / height

    def _resolve_point(self, event):
        x = event.get("x")
        y = event.get("y")
        if x is None or y is None:
            return 0, 0
        if event.get("coord") == "norm":
            left, top, width, height = get_virtual_screen_rect()
            return int(round(left + x * width)), int(round(top + y * height))
        return int(round(x)), int(round(y))

    def _send_mouse_move(self, x, y):
        left, top, width, height = get_virtual_screen_rect()
        if width <= 1 or height <= 1:
            return
        abs_x = int(round((x - left) * 65535 / (width - 1)))
        abs_y = int(round((y - top) * 65535 / (height - 1)))
        flags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK
        self._send_mouse_input(flags, abs_x, abs_y, 0)

    def _send_mouse_click(self, button_str, pressed):
        flags = 0
        name = button_str.split(".")[-1]
        if name == "left":
            flags = MOUSEEVENTF_LEFTDOWN if pressed else MOUSEEVENTF_LEFTUP
        elif name == "right":
            flags = MOUSEEVENTF_RIGHTDOWN if pressed else MOUSEEVENTF_RIGHTUP
        elif name == "middle":
            flags = MOUSEEVENTF_MIDDLEDOWN if pressed else MOUSEEVENTF_MIDDLEUP
        if flags:
            self._send_mouse_input(flags, 0, 0, 0)

    def _send_mouse_input(self, flags, x, y, data):
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        extra = ctypes.c_ulong(0)
        mi = MOUSEINPUT(x, y, data, flags, 0, ctypes.cast(ctypes.pointer(extra), wintypes.ULONG_PTR))
        inp = INPUT(0, mi)
        user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))

    def _play_key_event(self, event):
        controller = keyboard.Controller()
        if event.get("is_char"):
            key_obj = keyboard.KeyCode.from_char(event.get("key", ""))
        else:
            key_name = event.get("key")
            key_obj = getattr(keyboard.Key, key_name, None)
        if not key_obj:
            return
        if event.get("pressed"):
            controller.press(key_obj)
        else:
            controller.release(key_obj)

    def _update_stats(self):
        self.event_count_var.set(str(len(self.events)))
        if self.events:
            duration = self.events[-1]["t"]
        else:
            duration = 0.0
        self.duration_var.set(f"{duration:.2f}s")

    def save_macro(self):
        if not self.events:
            messagebox.showinfo("Empty", "No macro recorded to save.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("Macro files", "*.json")],
        )
        if not path:
            return
        data = {
            "created": time.strftime("%Y-%m-%d %H:%M:%S"),
            "target_title": self.target_title_var.get().strip(),
            "hotkeys": {
                "record": self.hotkey_record_var.get().strip(),
                "stop": self.hotkey_stop_var.get().strip(),
                "play": self.hotkey_play_var.get().strip(),
                "kill": self.hotkey_kill_var.get().strip(),
            },
            "loop_enabled": self.loop_var.get(),
            "loop_interval_seconds": self.loop_interval_var.get().strip(),
            "loop_count": self.loop_count_var.get().strip(),
            "background_keys": [
                {
                    "enabled": self.bg_key_enabled_vars[i].get(),
                    "key": self.bg_key_name_vars[i].get().strip(),
                    "interval": self.bg_key_interval_vars[i].get().strip(),
                }
                for i in range(10)
            ],
            "events": self.events,
        }
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2)
        self.status_var.set("Saved")

    def load_macro(self):
        path = filedialog.askopenfilename(
            filetypes=[("Macro files", "*.json")]
        )
        if not path:
            return
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
        self.events = data.get("events", [])
        self.target_title_var.set(data.get("target_title", self.target_title_var.get()))
        hotkeys = data.get("hotkeys", {})
        self.hotkey_record_var.set(hotkeys.get("record", self.hotkey_record_var.get()))
        self.hotkey_stop_var.set(hotkeys.get("stop", self.hotkey_stop_var.get()))
        self.hotkey_play_var.set(hotkeys.get("play", self.hotkey_play_var.get()))
        self.hotkey_kill_var.set(hotkeys.get("kill", self.hotkey_kill_var.get()))
        self.loop_var.set(bool(data.get("loop_enabled", self.loop_var.get())))
        saved_interval = data.get("loop_interval_seconds", self.loop_interval_var.get())
        self.loop_interval_var.set(str(saved_interval))
        
        saved_count = data.get("loop_count", self.loop_count_var.get())
        self.loop_count_var.set(str(saved_count))
        
        bg_keys = data.get("background_keys", [])
        for i in range(10):
            if i < len(bg_keys):
                bg_key = bg_keys[i]
                self.bg_key_enabled_vars[i].set(bool(bg_key.get("enabled", False)))
                self.bg_key_name_vars[i].set(bg_key.get("key", f"f{i+1}"))
                self.bg_key_interval_vars[i].set(str(bg_key.get("interval", "10.0")))
        
        self.apply_hotkeys(show_message=False)
        self._update_stats()
        self.status_var.set("Loaded")

    def on_close(self):
        print("[DEBUG] Closing app, stopping all listeners...")
        self.stop_all()
        if self.mouse_listener:
            try:
                self.mouse_listener.stop()
            except Exception:
                pass
            self.mouse_listener = None
        if self.keyboard_listener:
            try:
                self.keyboard_listener.stop()
            except Exception:
                pass
            self.keyboard_listener = None
        if self.hotkey_listener:
            try:
                self.hotkey_listener.stop()
            except Exception:
                pass
            self.hotkey_listener = None
        time.sleep(0.2)
        print("[DEBUG] All listeners stopped")
        self.root.destroy()


def main():
    root = tk.Tk()
    app = MacroApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
