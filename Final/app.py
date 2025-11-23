import tkinter as tk
from tkinter import messagebox
import threading
import time
from http.server import HTTPServer, SimpleHTTPRequestHandler
import json
import urllib.parse
import pandas as pd
from pynput.mouse import Listener as MouseListener
from pynput.keyboard import Listener as KeyListener
import os
import webbrowser

# ==============================
# CONFIG
# ==============================
PORT = 8000
LOG_FILE = "activity_log.xlsx"

# This will store all tracking events before saving to Excel
event_buffer = []
last_event_time = time.time()


# ==============================
# EXCEL APPENDER FUNCTION
# ==============================
def append_to_excel(data):
    df_new = pd.DataFrame([data])

    # If file exists, append
    if os.path.exists(LOG_FILE):
        df_old = pd.read_excel(LOG_FILE)
        df_out = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df_out = df_new

    df_out.to_excel(LOG_FILE, index=False)


# ==============================
# CUSTOM REQUEST HANDLER (for portal)
# ==============================
class PortalHandler(SimpleHTTPRequestHandler):

    def do_GET(self):
        global event_buffer, last_event_time

        parsed = urllib.parse.urlparse(self.path)

        # Handle portal -> app logging route
        if parsed.path == "/log_event":
            query = urllib.parse.parse_qs(parsed.query)
            if "payload" in query:
                payload = json.loads(query["payload"][0])

                now = time.time()
                idle_time = round(now - last_event_time, 3)
                last_event_time = now

                log_entry = {
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "event_type": payload.get("type", ""),
                    "description": payload.get("detail", ""),
                    "mouse_x": payload.get("x", ""),
                    "mouse_y": payload.get("y", ""),
                    "idle_time": idle_time
                }

                append_to_excel(log_entry)

            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"OK")
            return

        # otherwise serve HTML/CSS normally
        return super().do_GET()


# ==============================
# SERVER THREAD
# ==============================
def start_server():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    server = HTTPServer(("localhost", PORT), PortalHandler)
    server.serve_forever()


# ==============================
# TKINTER APP (login -> questions -> portal)
# ==============================
class App:

    def __init__(self, root):
        self.root = root
        root.title("Tracking App")
        root.geometry("400x300")

        self.login_ui()

    # Login UI
    def login_ui(self):
        self.clear()
        tk.Label(self.root, text="Login", font=("Arial", 18)).pack(pady=10)

        tk.Label(self.root, text="Username").pack()
        self.username = tk.Entry(self.root)
        self.username.pack()

        tk.Label(self.root, text="Password").pack()
        self.password = tk.Entry(self.root, show="*")
        self.password.pack()

        tk.Button(self.root, text="Login", command=self.goto_questions).pack(pady=15)

    # Questions UI
    def goto_questions(self):
        if not self.username.get() or not self.password.get():
            messagebox.showwarning("Error", "Enter username & password")
            return

        self.clear()
        tk.Label(self.root, text="Questions", font=("Arial", 16)).pack(pady=10)

        self.answers = []
        questions = [
            "What is your age?",
            "Your favourite subject?",
            "Your favourite color?",
            "Your hobbies?",
            "Are you working or student?",
            "Your favourite app?"
        ]

        for q in questions:
            tk.Label(self.root, text=q).pack()
            e = tk.Entry(self.root)
            e.pack()
            self.answers.append(e)

        tk.Button(self.root, text="Submit", command=self.open_portal).pack(pady=15)

    # Open portal + start tracking
    def open_portal(self):
        self.clear()
        tk.Label(self.root, text="Portal Opened", font=("Arial", 16)).pack(pady=20)
        tk.Label(self.root, text="Browser opened. Tracking started...", fg="green").pack()

        # open colorful portal
        webbrowser.open(f"http://localhost:{PORT}/portal.html")

    # Utility
    def clear(self):
        for widget in self.root.winfo_children():
            widget.destroy()


# ==============================
# MOUSE + KEYBOARD TRACKERS
# ==============================
def on_mouse(x, y):
    append_to_excel({
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "event_type": "mouse_move",
        "description": "",
        "mouse_x": x,
        "mouse_y": y,
        "idle_time": ""
    })


def on_click(x, y, button, pressed):
    if pressed:
        append_to_excel({
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "event_type": "mouse_click",
            "description": str(button),
            "mouse_x": x,
            "mouse_y": y,
            "idle_time": ""
        })


def on_key(key):
    append_to_excel({
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "event_type": "key_press",
        "description": str(key),
        "mouse_x": "",
        "mouse_y": "",
        "idle_time": ""
    })


# ==============================
# MAIN
# ==============================
if __name__ == "__main__":
    # Start server thread
    threading.Thread(target=start_server, daemon=True).start()

    # Start mouse listener
    MouseListener(on_move=on_mouse, on_click=on_click).start()

    # Start keyboard listener
    KeyListener(on_press=on_key).start()

    # Launch Tkinter
    root = tk.Tk()
    App(root)
    root.mainloop()
