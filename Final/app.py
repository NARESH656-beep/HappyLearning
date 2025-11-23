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

mouse_listener = None
keyboard_listener = None

last_event_time = time.time()


# ==============================
# EXCEL LOGGING FUNCTION
# ==============================
def append_to_excel(data):
    df_new = pd.DataFrame([data])

    if os.path.exists(LOG_FILE):
        df_old = pd.read_excel(LOG_FILE)
        df_out = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df_out = df_new

    df_out.to_excel(LOG_FILE, index=False)


# ==============================
# DESCRIPTION MAPPER
# ==============================
def convert_description(raw_type, raw_detail):
    mapping = {
        "emp": "Opened Employee Details Section",
        "hr": "Opened HR Policies Section",
        "leave": "Opened Leave Application Page",
        "business": "Opened Business Overview",
        "help": "Opened Help Section",
        "about": "Opened About Company Section",
        "open_dashboard": "Opened Dashboard",
        "edit_profile": "Clicked Edit Profile",
        "download_policy": "Downloaded HR Policy Document",
        "leave_apply": "Applied Leave",
    }

    # Scroll events
    if raw_type == "scroll":
        return f"Scrolled Page ({raw_detail})"

    if raw_detail in mapping:
        return mapping[raw_detail]

    if raw_type == "button_click":
        return f"Button Clicked: {raw_detail}"

    if raw_type == "link_click":
        return f"Navigation Click: {raw_detail}"

    return raw_detail or raw_type


# ==============================
# HTTP HANDLER
# ==============================
class PortalHandler(SimpleHTTPRequestHandler):

    def do_GET(self):
        global last_event_time

        parsed = urllib.parse.urlparse(self.path)

        if parsed.path == "/log_event":
            query = urllib.parse.parse_qs(parsed.query)

            if "payload" in query:
                payload = json.loads(query["payload"][0])

                now = time.time()
                idle_time = round(now - last_event_time, 3)
                last_event_time = now

                readable_description = convert_description(
                    payload.get("type", ""),
                    payload.get("detail", "")
                )

                log_row = {
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "event_type": payload.get("type", ""),
                    "description": readable_description,
                    "mouse_x": payload.get("x", ""),
                    "mouse_y": payload.get("y", ""),
                    "idle_time": idle_time
                }

                append_to_excel(log_row)

            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"OK")
            return

        return super().do_GET()


# ==============================
# START SERVER
# ==============================
def start_server():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    server = HTTPServer(("localhost", PORT), PortalHandler)
    server.serve_forever()


# ==============================
# INPUT LISTENERS
# ==============================
def on_mouse(x, y):
    append_to_excel({
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "event_type": "mouse_move",
        "description": "Mouse Moved",
        "mouse_x": x,
        "mouse_y": y,
        "idle_time": ""
    })


def on_click(x, y, button, pressed):
    if pressed:
        append_to_excel({
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "event_type": "mouse_click",
            "description": f"Mouse Clicked ({button})",
            "mouse_x": x,
            "mouse_y": y,
            "idle_time": ""
        })


def on_key(key):
    append_to_excel({
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "event_type": "key_press",
        "description": f"Key Pressed ({key})",
        "mouse_x": "",
        "mouse_y": "",
        "idle_time": ""
    })


# ==============================
# TKINTER APPLICATION
# ==============================
class App:

    def __init__(self, root):
        self.root = root
        root.title("Tracking App")
        root.geometry("450x350")

        self.login_ui()

    def clear(self):
        for w in self.root.winfo_children():
            w.destroy()

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

    def goto_questions(self):
        if not self.username.get() or not self.password.get():
            messagebox.showwarning("Error", "Enter username & password")
            return

        self.clear()
        tk.Label(self.root, text="Questions", font=("Arial", 16)).pack(pady=10)

        self.questions = [
            "What is your age?",
            "Your favourite subject?",
            "Your favourite color?",
            "Your hobbies?",
            "Are you working or student?",
            "Your favourite app?"
        ]

        for q in self.questions:
            tk.Label(self.root, text=q).pack()
            tk.Entry(self.root).pack()

        tk.Button(self.root, text="Submit", command=self.open_portal).pack(pady=15)

    def open_portal(self):
        self.clear()
        tk.Label(self.root, text="Portal Opened", font=("Arial", 16)).pack(pady=20)

        tk.Label(self.root, text="Browser opened. Tracking Started", fg="green").pack()

        tk.Button(self.root, text="End Session & Exit", fg="white", bg="red",
                  command=self.stop_and_exit).pack(pady=30)

        webbrowser.open(f"http://localhost:{PORT}/portal.html")

    def stop_and_exit(self):
        global mouse_listener, keyboard_listener

        try:
            if mouse_listener:
                mouse_listener.stop()
            if keyboard_listener:
                keyboard_listener.stop()
        except:
            pass

        messagebox.showinfo("Stopped", "Tracking stopped and session saved.")
        self.root.destroy()


# ==============================
# START EVERYTHING
# ==============================
if __name__ == "__main__":
    threading.Thread(target=start_server, daemon=True).start()

    mouse_listener = MouseListener(on_move=on_mouse, on_click=on_click)
    mouse_listener.start()

    keyboard_listener = KeyListener(on_press=on_key)
    keyboard_listener.start()

    root = tk.Tk()
    App(root)
    root.mainloop()
