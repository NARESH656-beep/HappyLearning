import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import pandas as pd
import threading
import time
import os
import webbrowser
from http.server import SimpleHTTPRequestHandler, HTTPServer
from pynput import mouse, keyboard

# ================= CONFIG =================
OUTPUT_FILE = "activity_log.xlsx"
IDLE_THRESHOLD = 5
PORTAL_URL = "http://localhost:8000/portal.html"

event_log = []
_last_event_time = None
TRACKING_STARTED = False
# ==========================================


# ------------ EVENT TRACKING FUNCTIONS ------------

def log_event(event_type, detail=None, x=None, y=None):
    global _last_event_time
    ts = datetime.utcnow()

    if _last_event_time:
        idle_gap = (ts - _last_event_time).total_seconds()
        if idle_gap > IDLE_THRESHOLD:
            event_log.append({
                "timestamp": ts.isoformat(),
                "event_type": "idle",
                "detail": f"idle_{idle_gap:.2f}_sec",
                "x": None,
                "y": None
            })

    _last_event_time = ts

    event_log.append({
        "timestamp": ts.isoformat(),
        "event_type": event_type,
        "detail": detail,
        "x": x,
        "y": y
    })


def on_mouse_click(x, y, button, pressed):
    if pressed:
        log_event("mouse_click", str(button), x, y)


def on_key_press(key):
    try:
        detail = key.char
    except:
        detail = str(key)
    log_event("key_press", detail)


def start_tracking():
    global TRACKING_STARTED
    if TRACKING_STARTED:
        return
    TRACKING_STARTED = True

    threading.Thread(target=lambda: mouse.Listener(on_click=on_mouse_click).run(), daemon=True).start()
    threading.Thread(target=lambda: keyboard.Listener(on_press=on_key_press).run(), daemon=True).start()


def save_to_excel():
    df = pd.DataFrame(event_log)

    if os.path.exists(OUTPUT_FILE):
        old = pd.read_excel(OUTPUT_FILE)
        df = pd.concat([old, df], ignore_index=True)

    df.to_excel(OUTPUT_FILE, index=False)


# ------------ LOCAL WEB SERVER ------------
def start_server():
    server = HTTPServer(("localhost", 8000), SimpleHTTPRequestHandler)
    threading.Thread(target=server.serve_forever, daemon=True).start()


# ------------ TKINTER UI PAGES ------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Office Portal Access")
        self.geometry("600x400")

        self.container = ttk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self.frames = {}
        for F in (LoginPage, QuestionsPage, PortalPage):
            frame = F(parent=self.container, controller=self)
            self.frames[F.__name__] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show("LoginPage")

    def show(self, name):
        self.frames[name].tkraise()


# ------------ LOGIN PAGE ------------
class LoginPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=20)
        self.controller = controller

        ttk.Label(self, text="Employee Login", font=("Arial", 18)).pack(pady=10)

        ttk.Label(self, text="Username").pack()
        self.user = tk.StringVar()
        ttk.Entry(self, textvariable=self.user).pack(pady=5)

        ttk.Label(self, text="Password").pack()
        self.pwd = tk.StringVar()
        ttk.Entry(self, textvariable=self.pwd, show="*").pack(pady=5)

        ttk.Button(self, text="Login", command=self.login).pack(pady=20)

    def login(self):
        if not self.user.get() or not self.pwd.get():
            messagebox.showerror("Error", "Enter username and password.")
            return

        self.controller.show("QuestionsPage")


# ------------ QUESTIONS PAGE ------------
class QuestionsPage(ttk.Frame):
    QUESTIONS = [
        "1. What is your age?",
        "2. What are your interests?",
        "3. Favourite food?",
        "4. Favourite movie?",
        "5. What do you enjoy doing?",
        "6. Any hobbies?"
    ]

    def __init__(self, parent, controller):
        super().__init__(parent, padding=20)
        self.controller = controller

        ttk.Label(self, text="Random Questions", font=("Arial", 16)).pack()

        self.entries = []
        for q in self.QUESTIONS:
            ttk.Label(self, text=q).pack(anchor="w")
            ent = ttk.Entry(self)
            ent.pack(fill="x", pady=5)
            self.entries.append(ent)

        ttk.Button(self, text="Submit", command=self.go_to_portal).pack(pady=20)

    def go_to_portal(self):
        # Do NOT save answers.
        self.controller.show("PortalPage")


# ------------ PORTAL PAGE ------------
class PortalPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=20)
        self.controller = controller

        ttk.Label(self, text="Office Portal", font=("Arial", 18)).pack(pady=10)

        ttk.Button(self, text="Enter Portal Website", command=self.open_portal).pack(pady=10)
        ttk.Button(self, text="Save Tracking Log", command=self.save_log).pack(pady=10)
        ttk.Button(self, text="Exit", command=self.quit).pack(pady=10)

    def open_portal(self):
        start_tracking()
        webbrowser.open(PORTAL_URL)
        messagebox.showinfo("Tracking Started", "Mouse, keyboard & idle activity now being recorded.")

    def save_log(self):
        save_to_excel()
        messagebox.showinfo("Saved", "Activity Log saved to activity_log.xlsx")


# ---------- RUN ----------
if __name__ == "__main__":
    start_server()
    App().mainloop()
