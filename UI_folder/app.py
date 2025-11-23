"""
app.py

Tkinter UI + local portal + background activity tracking.

Flow:
 - Login page (username, password) -> Questions page -> Submit
 - On Submit: open local portal (portal.html) served by local HTTP server
 - Start background tracking (mouse clicks, key presses, idle)
 - Save only activity events to activity_log.xlsx (no answers, no ids)

How to run:
 - Put app.py and portal.html in the same folder (UI_folder)
 - (Optional) you may keep the image referenced in portal.html at:
     /mnt/data/12963430-c73a-46ba-8baa-19480aa2ecb7.png
 - Install dependencies:
     pip install pynput pandas openpyxl
 - Run:
     python app.py
"""

import os
import threading
import time
import webbrowser
from http.server import HTTPServer, SimpleHTTPRequestHandler
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox

from pynput import mouse, keyboard

# ---------- CONFIG ----------
OUTPUT_FILE = "activity_log.xlsx"
IDLE_THRESHOLD = 5.0       # seconds threshold to register idle event
SERVER_PORT = 8000         # fixed port for local portal
PORTAL_FILENAME = "portal.html"
# ----------------------------

# Global event storage (only events; NOT saving ids or answers)
event_log = []
_last_event_time = None
_tracking_started = False

# --------------- Activity logging ---------------
def _now_iso():
    return datetime.utcnow().isoformat()

def _record_idle_if_needed(now_ts):
    global _last_event_time
    if _last_event_time is None:
        _last_event_time = now_ts
        return
    gap = (now_ts - _last_event_time).total_seconds()
    if gap > IDLE_THRESHOLD:
        event_log.append({
            "timestamp": now_ts.isoformat(),
            "event_type": "idle",
            "detail": f"idle_for_{gap:.2f}_sec",
            "x": None,
            "y": None
        })
    _last_event_time = now_ts

def record_event(event_type, detail=None, x=None, y=None):
    """
    Record event and detect idle gaps > IDLE_THRESHOLD.
    Stored fields: timestamp, event_type, detail, x, y
    """
    now_ts = datetime.utcnow()
    _record_idle_if_needed(now_ts)
    event_log.append({
        "timestamp": now_ts.isoformat(),
        "event_type": event_type,
        "detail": detail,
        "x": x,
        "y": y
    })

# pynput callbacks
def _on_click(x, y, button, pressed):
    if pressed:
        record_event("mouse_click", str(button), x, y)

def _on_scroll(x, y, dx, dy):
    # optional: record scroll events as well
    record_event("mouse_scroll", f"dx={dx},dy={dy}", x, y)

def _on_press(key):
    try:
        k = key.char
    except AttributeError:
        k = str(key)
    record_event("key_press", k, None, None)

# --------------- Background listeners ---------------
def start_listeners():
    """Start pynput listeners (mouse, keyboard) in daemon threads."""
    global _tracking_started
    if _tracking_started:
        return
    _tracking_started = True

    mouse_listener = mouse.Listener(on_click=_on_click, on_scroll=_on_scroll)
    keyboard_listener = keyboard.Listener(on_press=_on_press)
    mouse_listener.daemon = True
    keyboard_listener.daemon = True
    mouse_listener.start()
    keyboard_listener.start()

# --------------- Save to Excel ---------------
def save_events_to_excel():
    """Append the current event_log to OUTPUT_FILE (activity_log.xlsx)."""
    df = pd.DataFrame(event_log, columns=["timestamp", "event_type", "detail", "x", "y"])
    if os.path.exists(OUTPUT_FILE):
        try:
            existing = pd.read_excel(OUTPUT_FILE)
            combined = pd.concat([existing, df], ignore_index=True, sort=False)
            combined.to_excel(OUTPUT_FILE, index=False)
        except Exception as e:
            # Fallback: save alternate file
            alt = f"activity_log_{int(time.time())}.xlsx"
            df.to_excel(alt, index=False)
            raise RuntimeError(f"Could not append to {OUTPUT_FILE}: {e}. Saved to {alt} instead.")
    else:
        df.to_excel(OUTPUT_FILE, index=False)

# --------------- Local HTTP server ---------------
class QuietHandler(SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        # suppress default console logging
        pass

def start_local_server(port=SERVER_PORT):
    """
    Serve files from the directory where app.py is located.
    This function spawns a background thread to run the HTTP server.
    """
    # ensure server serves from same directory as app.py
    web_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(web_dir)

    def _serve():
        server = HTTPServer(("localhost", port), QuietHandler)
        try:
            server.serve_forever()
        except Exception:
            pass

    thread = threading.Thread(target=_serve, daemon=True)
    thread.start()
    # give server a moment to start
    time.sleep(0.3)
    return f"http://localhost:{port}/{PORTAL_FILENAME}"

# --------------- Tkinter UI ---------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Office Portal Tracker")
        self.geometry("720x520")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        self.frames = {}

        for F in (LoginPage, QuestionsPage, SubmitPage):
            page = F(container, self)
            self.frames[F.__name__] = page
            page.grid(row=0, column=0, sticky="nsew")

        self.show_frame("LoginPage")

    def show_frame(self, name):
        frame = self.frames[name]
        frame.tkraise()

    def _on_close(self):
        if messagebox.askokcancel("Quit", "Exit application and save logs?"):
            # best-effort save
            try:
                save_events_to_excel()
            except Exception:
                pass
            self.destroy()

class LoginPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=18)
        self.controller = controller
        ttk.Label(self, text="Login", font=("Helvetica", 20, "bold")).pack(pady=8)

        frm = ttk.Frame(self)
        frm.pack(pady=12, fill="x", padx=20)
        ttk.Label(frm, text="Username:").grid(row=0, column=0, sticky="w")
        self.username = tk.StringVar()
        ttk.Entry(frm, textvariable=self.username).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Label(frm, text="Password:").grid(row=1, column=0, sticky="w")
        self.password = tk.StringVar()
        ttk.Entry(frm, textvariable=self.password, show="*").grid(row=1, column=1, sticky="ew", padx=8)
        frm.columnconfigure(1, weight=1)

        ttk.Button(self, text="Login", command=self._do_login).pack(pady=14)

    def _do_login(self):
        if not self.username.get().strip() or not self.password.get().strip():
            messagebox.showerror("Missing", "Please enter username and password.")
            return
        # NOTE: We do NOT save credentials anywhere.
        self.controller.show_frame("QuestionsPage")

class QuestionsPage(ttk.Frame):
    QUESTIONS = [
        "What is your age?",
        "What are your interests?",
        "What's your favourite food?",
        "What's your favourite movie?",
        "What do you enjoy doing in free time?",
        "Any hobbies?"
    ]
    def __init__(self, parent, controller):
        super().__init__(parent, padding=14)
        self.controller = controller
        ttk.Label(self, text="Quick Questions (Not saved)", font=("Helvetica", 16)).pack(pady=6)
        self.inputs = []
        qframe = ttk.Frame(self)
        qframe.pack(fill="both", expand=True, padx=12)
        for q in self.QUESTIONS:
            ttk.Label(qframe, text=q).pack(anchor="w", pady=(8,2))
            ent = ttk.Entry(qframe)
            ent.pack(fill="x")
            self.inputs.append(ent)

        btn = ttk.Button(self, text="Submit and Enter Portal", command=self._submit)
        btn.pack(pady=12)

    def _submit(self):
        # We intentionally do NOT store the answers anywhere.
        self.controller.show_frame("SubmitPage")

class SubmitPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=14)
        self.controller = controller

        ttk.Label(self, text="Portal Ready", font=("Helvetica", 16)).pack(pady=6)
        ttk.Label(self, text="Click the button below to open the Office Portal. Tracking will start when portal opens.").pack(pady=6)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=12)
        ttk.Button(btn_frame, text="Open Portal & Start Tracking", command=self.open_portal).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Save Activity Log", command=self.save_log).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Quit (Save & Exit)", command=self.quit_and_save).pack(side="left", padx=6)

        self.summary = tk.Text(self, height=10)
        self.summary.pack(fill="both", expand=True, pady=10, padx=6)
        self.after(1500, self._periodic_update)

    def open_portal(self):
        # Start local server and open portal page
        url = start_local_server()
        webbrowser.open(url)
        # Start background input tracking
        start_listeners()
        messagebox.showinfo("Tracking started", "Portal opened and tracking started in the background.")
        self._update_summary()

    def save_log(self):
        try:
            save_events_to_excel()
            messagebox.showinfo("Saved", f"Events saved to {OUTPUT_FILE}")
        except Exception as e:
            messagebox.showerror("Save failed", str(e))

    def quit_and_save(self):
        try:
            save_events_to_excel()
        except Exception:
            pass
        self.controller._on_close()

    def _update_summary(self):
        self.summary.delete("1.0", "end")
        self.summary.insert("1.0", f"Captured events: {len(event_log)}\n")
        if event_log:
            last = event_log[-1]
            self.summary.insert("end", f"Last: {last['timestamp']} | {last['event_type']} | {last.get('detail')}\n")
        self.summary.insert("end", "\nNote: Username, password, and answers are NOT saved to Excel.\n")

    def _periodic_update(self):
        self._update_summary()
        self.after(1500, self._periodic_update)

# --------------- Entrypoint ---------------
def start_local_server():
    """
    Start server and return URL to portal file.
    Uses SERVER_PORT and serves from the script directory.
    """
    web_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(web_dir)
    # Make sure portal file exists
    portal_path = os.path.join(web_dir, PORTAL_FILENAME)
    if not os.path.exists(portal_path):
        raise FileNotFoundError(f"{PORTAL_FILENAME} not found in {web_dir}. Please place the portal HTML file there.")
    # start server thread
    def _serve():
        httpd = HTTPServer(("localhost", SERVER_PORT), QuietHandler)
        try:
            httpd.serve_forever()
        except Exception:
            pass
    t = threading.Thread(target=_serve, daemon=True)
    t.start()
    time.sleep(0.25)
    return f"http://localhost:{SERVER_PORT}/{PORTAL_FILENAME}"

if __name__ == "__main__":
    # Start the app (server will start when user opens portal from Submit page)
    App().mainloop()
