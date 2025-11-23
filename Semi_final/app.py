"""
app.py
Comprehensive Office Portal tracker:
 - Login (username/password) -> Questions -> Submit -> Open portal
 - Background tracking: pynput (mouse clicks/scrolls, keyboard), idle detection
 - Client-side events (scroll, nav, portal buttons) are sent to /log_event
 - Exports: per-user Excel (with session sheet), CSV, SQLite
 - Dashboard with matplotlib (events over time, by type, click scatter)
"""
import os
import threading
import time
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, unquote
from datetime import datetime
import webbrowser
import json
import sqlite3
import tempfile

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd
# matplotlib is optional for environments where it's not available to the linter/runtime;
# attempt to import and configure it, otherwise provide minimal fallbacks so the UI still runs.
try:
    import matplotlib
    matplotlib.use('Agg')  # use non-interactive backend for safe rendering to canvas
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import matplotlib.pyplot as plt
    MATPLOTLIB_AVAILABLE = True
except Exception:
    MATPLOTLIB_AVAILABLE = False
    import traceback
    print("Warning: matplotlib is not available; dashboard plotting will be disabled.")
    traceback.print_exc()
    # Minimal fallback implementations to allow the rest of the app to import and run.
    class _DummyAxes:
        def clear(self): pass
        def plot(self, *a, **k): pass
        def bar(self, *a, **k): pass
        def text(self, *a, **k): pass
        def set_title(self, *a, **k): pass
        def tick_params(self, *a, **k): pass

    class _DummyFigure:
        def __init__(self, figsize=None): pass
        def add_subplot(self, *a, **k):
            return _DummyAxes()

    class _DummyCanvas:
        def __init__(self, fig, master=None):
            self._master = master
        def get_tk_widget(self):
            # return an empty frame so pack/place calls work
            return tk.Frame(self._master)
        def draw(self): pass

    FigureCanvasTkAgg = _DummyCanvas
    class _DummyPLT:
        Figure = _DummyFigure
    plt = _DummyPLT()

from pynput import mouse, keyboard
import uuid

# ---------------- CONFIG ----------------
OUTPUT_DIR = os.path.abspath(".")
IDLE_THRESHOLD = 5.0  # seconds
SERVER_PORT = 8000
PORTAL_FILENAME = "portal.html"
# ----------------------------------------

# Shared global event list (holds dicts)
event_log = []   # each: {timestamp, event_type, detail, x, y, source (client/server)}
_event_log_lock = threading.Lock()

# Session metadata
current_session = {
    "session_id": None,
    "username": None,
    "start_time": None,
    "end_time": None
}

_tracking_started = False
_server_thread = None
_httpd = None

# ---------- Utility functions ----------
def now_iso():
    return datetime.utcnow().isoformat()

def append_event(ev: dict):
    """Thread-safe append event to global log."""
    with _event_log_lock:
        event_log.append(ev)

def record_idle_if_needed(now_ts):
    """Detect idle gap based on last recorded timestamp in event_log."""
    with _event_log_lock:
        if not event_log:
            return
        last_ts = datetime.fromisoformat(event_log[-1]["timestamp"])
    gap = (now_ts - last_ts).total_seconds()
    if gap > IDLE_THRESHOLD:
        append_event({
            "timestamp": now_ts.isoformat(),
            "event_type": "idle",
            "detail": f"idle_for_{gap:.2f}_sec",
            "x": None,
            "y": None,
            "source": "server"
        })

# ---------- pynput callbacks ----------
def on_click(x, y, button, pressed):
    if not pressed:
        return
    ts = datetime.utcnow()
    record_idle_if_needed(ts)
    append_event({
        "timestamp": ts.isoformat(),
        "event_type": "mouse_click",
        "detail": str(button),
        "x": float(x),
        "y": float(y),
        "source": "server"
    })

def on_scroll(x, y, dx, dy):
    ts = datetime.utcnow()
    record_idle_if_needed(ts)
    append_event({
        "timestamp": ts.isoformat(),
        "event_type": "mouse_scroll",
        "detail": f"dx={dx},dy={dy}",
        "x": float(x),
        "y": float(y),
        "source": "server"
    })

def on_press(key):
    ts = datetime.utcnow()
    record_idle_if_needed(ts)
    try:
        k = key.char
    except AttributeError:
        k = str(key)
    append_event({
        "timestamp": ts.isoformat(),
        "event_type": "key_press",
        "detail": k,
        "x": None,
        "y": None,
        "source": "server"
    })

# ---------- Start listeners ----------
def start_listeners():
    global _tracking_started
    if _tracking_started:
        return
    _tracking_started = True
    m = mouse.Listener(on_click=on_click, on_scroll=on_scroll)
    k = keyboard.Listener(on_press=on_press)
    m.daemon = True
    k.daemon = True
    m.start()
    k.start()

# ---------- Simple HTTP server accepts /log_event ---------
class PortalHandler(SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        # suppress console logs
        pass

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path.startswith("/log_event"):
            # Expect query params like ?type=scroll&pos=200&... or JSON in query 'payload'
            params = parse_qs(parsed.query)
            # If 'payload' param with JSON
            payload = {}
            if 'payload' in params:
                try:
                    payload = json.loads(unquote(params['payload'][0]))
                except Exception:
                    payload = {}
            # Merge params into payload
            for k, v in params.items():
                if k == 'payload': continue
                payload[k] = v[0] if v else ''
            # Build event
            ts = datetime.utcnow().isoformat()
            ev = {
                "timestamp": ts,
                "event_type": payload.get("type", "client_event"),
                "detail": payload.get("detail", payload.get("info", "")),
                "x": float(payload["x"]) if payload.get("x") else None,
                "y": float(payload["y"]) if payload.get("y") else None,
                "source": "client"
            }
            append_event(ev)
            # Simple ack
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(b'{"status":"ok"}')
            return
        else:
            # serve files normally (portal.html etc)
            return super().do_GET()

def start_local_server(port=SERVER_PORT):
    global _server_thread, _httpd
    # serve from script directory
    web_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(web_dir)
    _httpd = HTTPServer(("localhost", port), PortalHandler)
    _server_thread = threading.Thread(target=_httpd.serve_forever, daemon=True)
    _server_thread.start()
    time.sleep(0.2)
    return f"http://localhost:{port}/{PORTAL_FILENAME}"

# ---------- Saving: Excel, CSV, SQLite ----------
def export_data(username=None):
    """Export event_log to files (Excel with session sheet, CSV, SQLite)."""
    with _event_log_lock:
        df = pd.DataFrame(event_log)
    # ensure columns exist
    if df.empty:
        df = pd.DataFrame(columns=["timestamp","event_type","detail","x","y","source"])
    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    safe_user = username if username else "anonymous"
    base = os.path.join(OUTPUT_DIR, f"activity_log_{safe_user}_{ts}")
    # Excel with two sheets: events + session_info
    excel_path = base + ".xlsx"
    try:
        session_info = {
            "session_id": current_session.get("session_id"),
            "username": current_session.get("username"),
            "start_time": current_session.get("start_time"),
            "end_time": current_session.get("end_time")
        }
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="events", index=False)
            pd.DataFrame([session_info]).to_excel(writer, sheet_name="session_info", index=False)
    except Exception as e:
        raise
    # CSV
    csv_path = base + ".csv"
    df.to_csv(csv_path, index=False)
    # SQLite
    db_path = os.path.join(OUTPUT_DIR, "activity_events.db")
    conn = sqlite3.connect(db_path)
    df.to_sql("events", conn, if_exists="append", index=False)
    conn.close()
    return excel_path, csv_path, db_path

# ---------- Dashboard helpers ----------
def events_per_minute(df):
    if df.empty: return pd.DataFrame(columns=["minute","count"])
    df2 = df.copy()
    df2["ts"] = pd.to_datetime(df2["timestamp"])
    df2["minute"] = df2["ts"].dt.floor("T")
    return df2.groupby("minute").size().reset_index(name="count")

def event_type_counts(df):
    if df.empty: return pd.DataFrame(columns=["event_type","count"])
    return df.groupby("event_type").size().reset_index(name="count")

# ---------- Tkinter UI ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Office Portal Tracker - Advanced")
        self.geometry("1000x700")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        self.frames = {}
        for F in (LoginPage, QuestionsPage, SubmitPage, DashboardPage):
            page = F(container, self)
            self.frames[F.__name__] = page
            page.grid(row=0, column=0, sticky="nsew")
        self.show_frame("LoginPage")

    def show_frame(self, name):
        frame = self.frames[name]
        frame.tkraise()

    def _on_close(self):
        if messagebox.askokcancel("Exit", "Save logs and exit?"):
            try:
                save_path = export_data(current_session.get("username"))
            except Exception as e:
                print("Export failed:", e)
            # shutdown server if running
            try:
                if _httpd:
                    _httpd.shutdown()
            except Exception:
                pass
            self.destroy()

class LoginPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=20)
        self.controller = controller
        ttk.Label(self, text="Login", font=("Helvetica", 22, "bold")).pack(pady=10)
        frm = ttk.Frame(self)
        frm.pack(pady=8, padx=8, fill="x")
        ttk.Label(frm, text="Username").grid(row=0, column=0, sticky="w")
        self.user = tk.StringVar()
        ttk.Entry(frm, textvariable=self.user).grid(row=0, column=1, sticky="ew")
        ttk.Label(frm, text="Password").grid(row=1, column=0, sticky="w")
        self.pwd = tk.StringVar()
        ttk.Entry(frm, textvariable=self.pwd, show="*").grid(row=1, column=1, sticky="ew")
        frm.columnconfigure(1, weight=1)
        ttk.Button(self, text="Login", command=self.do_login).pack(pady=12)

    def do_login(self):
        u = self.user.get().strip()
        p = self.pwd.get().strip()
        if not u or not p:
            messagebox.showerror("Missing", "Enter username and password")
            return
        # set session metadata
        current_session["username"] = u
        current_session["session_id"] = str(uuid.uuid4())
        current_session["start_time"] = now_iso()
        current_session["end_time"] = None
        # proceed
        self.controller.show_frame("QuestionsPage")

class QuestionsPage(ttk.Frame):
    QUESTIONS = [
        "What is your age?",
        "What are your interests?",
        "Favourite food?",
        "Favourite movie?",
        "What do you enjoy doing?",
        "Any hobbies?"
    ]
    def __init__(self, parent, controller):
        super().__init__(parent, padding=14)
        self.controller = controller
        ttk.Label(self, text="Quick Questions (NOT saved)", font=("Helvetica", 18)).pack(pady=6)
        self.entries = []
        content = ttk.Frame(self)
        content.pack(fill="both", expand=True, padx=10)
        for q in self.QUESTIONS:
            ttk.Label(content, text=q).pack(anchor="w", pady=(8,2))
            e = ttk.Entry(content); e.pack(fill="x")
            self.entries.append(e)
        ttk.Button(self, text="Submit -> Open Portal", command=self.submit).pack(pady=10)

    def submit(self):
        # intentionally NOT saving answers
        self.controller.show_frame("SubmitPage")

class SubmitPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=12)
        self.controller = controller
        ttk.Label(self, text="Ready to Open Portal", font=("Helvetica", 18)).pack(pady=8)
        ttk.Label(self, text="Click the button below to open the portal and start tracking.").pack()
        btn_frame = ttk.Frame(self); btn_frame.pack(pady=12)
        ttk.Button(btn_frame, text="Open Portal & Start Tracking", command=self.open_and_start).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Go to Dashboard", command=lambda: controller.show_frame("DashboardPage")).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Export Now", command=self.export_now).pack(side="left", padx=6)

    def open_and_start(self):
        # start server (if not started)
        url = start_local_server()
        webbrowser.open(url)
        # start listeners
        start_listeners()
        # record session_start event
        append_event({
            "timestamp": now_iso(),
            "event_type": "session_start",
            "detail": f"session_id={current_session.get('session_id')}",
            "x": None,
            "y": None,
            "source": "app"
        })
        messagebox.showinfo("Started", f"Portal opened at {url}\nTracking started.")
        self.controller.show_frame("DashboardPage")

    def export_now(self):
        try:
            excel, csvp, dbp = export_data(current_session.get("username"))
            messagebox.showinfo("Exported", f"Excel: {excel}\nCSV: {csvp}\nDB: {dbp}")
        except Exception as e:
            messagebox.showerror("Export error", str(e))

class DashboardPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=10)
        self.controller = controller
        top = ttk.Frame(self)
        top.pack(fill="x")
        ttk.Label(top, text="Dashboard", font=("Helvetica", 20)).pack(side="left")
        ttk.Button(top, text="Refresh", command=self.refresh).pack(side="right")
        ttk.Button(top, text="Save & End Session", command=self.end_session).pack(side="right", padx=6)
        # left: charts, right: event list
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)
        left = ttk.Frame(body)
        left.pack(side="left", fill="both", expand=True)
        right = ttk.Frame(body, width=320)
        right.pack(side="right", fill="y")
        # matplotlib canvas placeholders
        self.fig1 = plt.Figure(figsize=(6,3))
        self.ax1 = self.fig1.add_subplot(111)
        self.canvas1 = FigureCanvasTkAgg(self.fig1, master=left)
        self.canvas1.get_tk_widget().pack(fill="both", expand=True)
        self.fig2 = plt.Figure(figsize=(6,3))
        self.ax2 = self.fig2.add_subplot(111)
        self.canvas2 = FigureCanvasTkAgg(self.fig2, master=left)
        self.canvas2.get_tk_widget().pack(fill="both", expand=True)
        # right: event count and list
        ttk.Label(right, text="Event Summary").pack(anchor="nw")
        self.summary_txt = tk.Text(right, width=40, height=20)
        self.summary_txt.pack(fill="y", expand=True)
        self.refresh()

    def refresh(self):
        # read events snapshot
        with _event_log_lock:
            df = pd.DataFrame(event_log)
        if df.empty:
            df = pd.DataFrame(columns=["timestamp","event_type","detail","x","y","source"])
        # Timeseries
        ts_df = events_per_minute(df)
        self.ax1.clear()
        if not ts_df.empty:
            self.ax1.plot(ts_df["minute"].astype(str), ts_df["count"])
            self.ax1.tick_params(axis='x', rotation=45)
            self.ax1.set_title("Events per minute")
        else:
            self.ax1.text(0.5,0.5,"No data", ha='center')
        self.canvas1.draw()
        # Event type counts
        et = event_type_counts(df)
        self.ax2.clear()
        if not et.empty:
            self.ax2.bar(et["event_type"], et["count"])
            self.ax2.set_title("Events by type")
            self.ax2.tick_params(axis='x', rotation=45)
        else:
            self.ax2.text(0.5,0.5,"No data", ha='center')
        self.canvas2.draw()
        # summary text
        summary_lines = []
        summary_lines.append(f"Total events: {len(df)}")
        summary_lines.append(f"Session id: {current_session.get('session_id')}")
        summary_lines.append(f"Username: {current_session.get('username')}")
        summary_lines.append(f"Start: {current_session.get('start_time')}")
        summary_lines.append(f"End: {current_session.get('end_time')}")
        if not df.empty:
            last = df.iloc[-1]
            summary_lines.append("")
            summary_lines.append("Last event:")
            summary_lines.append(f"{last['timestamp']} | {last['event_type']} | {last['detail']}")
        self.summary_txt.delete("1.0", "end")
        self.summary_txt.insert("1.0", "\n".join(summary_lines))

    def end_session(self):
        # record session end from app side
        current_session["end_time"] = now_iso()
        append_event({
            "timestamp": now_iso(),
            "event_type": "session_end",
            "detail": f"session_id={current_session.get('session_id')}",
            "x": None,
            "y": None,
            "source": "app"
        })
        try:
            excel, csvp, dbp = export_data(current_session.get("username"))
            messagebox.showinfo("Saved", f"Saved files:\n{excel}\n{csvp}\n{dbp}")
        except Exception as e:
            messagebox.showerror("Save failed", str(e))

# ---------- Entrypoint ----------
def main():
    # ensure portal file exists
    web_dir = os.path.dirname(os.path.abspath(__file__))
    portal_path = os.path.join(web_dir, PORTAL_FILENAME)
    if not os.path.exists(portal_path):
        print(f"portal.html not found in {web_dir}. Create the file and rerun.")
        return
    start_local_server(SERVER_PORT)  # server runs in background and accepts /log_event
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
