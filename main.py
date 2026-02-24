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
PLAYBACK_DELAY_SEC = 0.0
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
        self.absolute_mouse_var = tk.BooleanVar(value=True)

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

        self._build_ui()
        self._start_hotkey_listener()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self):
        frame = tk.Frame(self.root, padx=12, pady=12)
        frame.pack(fill=tk.BOTH, expand=True)

        info_label = tk.Label(frame, text="Hotkeys: Ctrl+1=Record | Ctrl+2=Stop | Ctrl+3=Play | Ctrl+Esc=Kill", 
                             bg="lightblue", fg="black", padx=8, pady=4)
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
        tk.Checkbutton(
            loop_row,
            text="High precision mouse",
            variable=self.absolute_mouse_var,
        ).pack(side=tk.LEFT, padx=(8, 0))

        file_row = tk.Frame(frame)
        file_row.grid(row=4, column=0, columnspan=2, pady=(10, 0), sticky="we")
        tk.Button(file_row, text="Load", command=self.load_macro).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(file_row, text="Save", command=self.save_macro).pack(side=tk.LEFT)

        info = tk.Frame(frame)
        info.grid(row=5, column=0, columnspan=2, pady=(12, 0), sticky="we")
        tk.Label(info, text="Events:").pack(side=tk.LEFT)
        tk.Label(info, textvariable=self.event_count_var, width=8).pack(side=tk.LEFT)
        tk.Label(info, text="Duration:").pack(side=tk.LEFT, padx=(12, 0))
        tk.Label(info, textvariable=self.duration_var, width=10).pack(side=tk.LEFT)

        status_row = tk.Frame(frame)
        status_row.grid(row=6, column=0, columnspan=2, pady=(10, 0), sticky="we")
        tk.Label(status_row, text="Status:").pack(side=tk.LEFT)
        tk.Label(status_row, textvariable=self.status_var).pack(side=tk.LEFT, padx=(8, 0))

        frame.columnconfigure(1, weight=1)

    def _start_hotkey_listener(self):
        def on_press(key):
            if key in (keyboard.Key.ctrl_l, keyboard.Key.ctrl_r):
                pressed.add("ctrl")
            if key == keyboard.Key.esc and "ctrl" in pressed:
                self.stop_all()
            if "ctrl" in pressed and isinstance(key, keyboard.KeyCode):
                vk = getattr(key, "vk", None)
                if vk == 0x31:  # '1'
                    self.root.after(0, self.start_record)
                elif vk == 0x32:  # '2'
                    self.root.after(0, self.stop_record)
                elif vk == 0x33:  # '3'
                    self.root.after(0, self.start_playback)

        def on_release(key):
            if key in (keyboard.Key.ctrl_l, keyboard.Key.ctrl_r):
                pressed.discard("ctrl")

        pressed = set()
        self.hotkey_listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        self.hotkey_listener.daemon = True
        self.hotkey_listener.start()

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
        self.status_var.set("Recording (Ctrl+2 or Ctrl+Esc to stop)")

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
        self.playing = True
        self.stop_event.clear()
        status = "Playing (Ctrl+Esc to stop)"
        if self.loop_var.get():
            status = "Playing (Loop, Ctrl+Esc to stop)"
        self.status_var.set(status)

        self.play_thread = threading.Thread(target=self._playback_worker, daemon=True)
        self.play_thread.start()

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
            print(f"[DEBUG] Waiting {PLAYBACK_DELAY_SEC}s before playback...")
            delay_end = time.monotonic() + PLAYBACK_DELAY_SEC
            while time.monotonic() < delay_end:
                if self.stop_event.is_set():
                    break
                time.sleep(0.05)
            if self.stop_event.is_set():
                break

            print(f"[DEBUG] Starting playback with {len(self.events)} events")
            base_time = time.monotonic()
            for i, event in enumerate(self.events):
                if self.stop_event.is_set():
                    break

                if event.get("t") == 0.0:
                    print(f"[DEBUG] Skipping initial position event")
                    continue

                target_time = base_time + event["t"]
                wait = target_time - time.monotonic()
                if wait > 0:
                    time.sleep(wait)

                try:
                    print(f"[DEBUG] Event {i}: {event['type']}")
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
                except Exception as e:
                    print(f"[DEBUG] Error in event {i}: {e}")

            if self.stop_event.is_set() or not self.loop_var.get():
                break

        print("[DEBUG] Playback finished")
        self.playing = False
        self.status_var.set("Idle")

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
