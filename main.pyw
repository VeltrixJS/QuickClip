import customtkinter as ctk
import keyboard
import mss
import cv2
import numpy as np
import os
import ctypes
import threading
import time
import webbrowser
import subprocess
import sys
import winerror
from collections import deque
from datetime import datetime
import win32api
import win32event
import win32con
import win32gui
import pystray
from PIL import Image


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


def get_videos_folder():
    try:
        buf = ctypes.create_unicode_buffer(300)
        ctypes.windll.shell32.SHGetFolderPathW(None, 0x000E, None, 0, buf)
        path = buf.value
        if path and os.path.exists(path):
            return path
    except:
        pass
    for name in ["Videos", "Vidéos", "Vidéo"]:
        p = os.path.join(os.path.expanduser("~"), name)
        if os.path.exists(p):
            return p
    return os.path.join(os.path.expanduser("~"), "Videos")


def check_single_instance():
    mutex = win32event.CreateMutex(None, False, "QuickClip_SingleInstance")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        return None
    return mutex


def bring_existing_to_front():
    def callback(hwnd, _):
        if win32gui.GetWindowText(hwnd) == "QuickClip":
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
    win32gui.EnumWindows(callback, None)


class HelpWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Comment utiliser QuickClip ?")
        self.geometry("420x400")
        self.resizable(False, False)
        self.grab_set()

        frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=15, pady=15)

        ctk.CTkLabel(frame, text="Guide QuickClip", font=("Arial", 15, "bold")).pack(pady=(0, 10))

        content = """
RACCOURCIS CLAVIER
━━━━━━━━━━━━━━━━━━━━━━
Maintiens C enfoncé + appuie sur des chiffres
pour accumuler la durée du clip.
Le clip est sauvegardé quand tu relâches C.

  C + 1         ->  10 secondes
  C + 2         ->  20 secondes
  C + 9         ->  90 secondes
  C + 1 + 1     ->  20 secondes
  C + 9 + 9     ->  180 secondes

Limite maximale : 10 minutes (600s)

DOSSIERS DE SAUVEGARDE
━━━━━━━━━━━━━━━━━━━━━━━━━━
Les clips sont sauvegardés dans :
  Videos\\QuickClips\\

  10s à 90s  ->  dossier "X sec"
  Autres     ->  dossier "Longs clips"

CHOIX DE L'ECRAN
━━━━━━━━━━━━━━━━━━━━
  Ecran principal  : ton écran de bureau
  Ecran secondaire : deuxième moniteur
  Dernier écran souris : capture l'écran
    où se trouvait ta souris en dernier

BARRE DES TACHES
━━━━━━━━━━━━━━━━━━━━
Fermer la fenêtre ne quitte pas le programme.
QuickClip continue en arrière-plan.
Clic droit sur l'icône -> Quitter pour fermer.
"""
        ctk.CTkLabel(frame, text=content, font=("Arial", 12), justify="left", anchor="w").pack(fill="x")
        ctk.CTkButton(self, text="Fermer", command=self.destroy, width=100).pack(pady=10)


class QuickClip(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("QuickClip")
        self.iconbitmap(resource_path("icône_QUICKCLIP.ico"))
        self.geometry("380x280")
        self.resizable(False, False)
        ctk.set_appearance_mode("dark")

        self.base_dir = os.path.join(get_videos_folder(), "QuickClips")
        self.fps = 10
        self.scale = 0.5
        self.max_seconds = 600

        self.buffer = deque(maxlen=self.max_seconds * self.fps)
        self.buffer_size = (0, 0)
        self.is_running = True

        self.accum_seconds = 0
        self.c_is_held = False
        self.last_mouse_monitor = 1

        self.setup_folders()
        self.build_ui()

        self.protocol("WM_DELETE_WINDOW", self.hide_window)

        threading.Thread(target=self.bg_capture, daemon=True).start()
        threading.Thread(target=self.track_mouse_monitor, daemon=True).start()
        threading.Thread(target=self.init_shortcuts, daemon=True).start()
        threading.Thread(target=self.setup_tray, daemon=True).start()

    def build_ui(self):
        ctk.CTkButton(self, text="?", width=28, height=28,
                      font=("Arial", 13, "bold"),
                      fg_color="#444", hover_color="#666",
                      corner_radius=14,
                      command=self.open_help).place(x=10, y=10)

        ctk.CTkButton(self, text="",
                      image=ctk.CTkImage(Image.open(resource_path("github.png")), size=(18, 18)),
                      command=lambda: webbrowser.open('https://github.com/VeltrixJS'),
                      fg_color="transparent", hover_color="#2c3e50",
                      border_width=0,
                      width=28, height=28).place(x=10, y=245)

        ctk.CTkLabel(self, text="QUICKCLIP", font=("Arial", 16, "bold")).pack(pady=(10, 5))

        self.screen_options = self.build_screen_options()
        self.screen_select = ctk.CTkOptionMenu(self, values=self.screen_options, width=300)
        self.screen_select.pack(pady=5)

        self.accum_label = ctk.CTkLabel(self, text="", text_color="#e67e22", font=("Arial", 12))
        self.accum_label.pack(pady=2)

        self.status_label = ctk.CTkLabel(self, text="", text_color="#3498db", font=("Arial", 11))
        self.status_label.pack(pady=2)

        ctk.CTkButton(self, text="  Ouvrir le dossier des clips",
                      command=self.open_clips_folder,
                      fg_color="#2c3e50", hover_color="#3d5166",
                      width=220, height=32).pack(pady=12)

    def build_screen_options(self):
        options = ["Ecran principal (1)"]
        try:
            with mss.mss() as sct:
                count = len(sct.monitors) - 1
            for i in range(2, count + 1):
                options.append(f"Ecran secondaire ({i})" if i == 2 else f"Ecran {i} ({i})")
        except:
            pass
        options.append("Dernier ecran souris")
        return options

    def open_help(self):
        HelpWindow(self)

    def open_clips_folder(self):
        os.makedirs(self.base_dir, exist_ok=True)
        subprocess.Popen(f'explorer "{self.base_dir}"')

    def setup_tray(self):
        image = Image.open(resource_path("icône_QUICKCLIP.ico"))
        menu = pystray.Menu(
            pystray.MenuItem("Ouvrir", self.show_window),
            pystray.MenuItem("Quitter", self.quit_app)
        )
        self.tray_icon = pystray.Icon("QuickClip", image, "QuickClip", menu)
        self.tray_icon.run()

    def hide_window(self):
        self.withdraw()

    def show_window(self):
        self.deiconify()
        self.lift()

    def quit_app(self):
        self.is_running = False
        self.tray_icon.stop()
        self.destroy()

    def setup_folders(self):
        try:
            os.makedirs(self.base_dir, exist_ok=True)
            for i in range(1, 10):
                os.makedirs(os.path.join(self.base_dir, f"{i*10} sec"), exist_ok=True)
            os.makedirs(os.path.join(self.base_dir, "Longs clips"), exist_ok=True)
        except Exception as e:
            print(f"Erreur dossier : {e}")

    def get_selected_monitor_index(self):
        choice = self.screen_select.get()
        if choice == "Dernier ecran souris":
            return self.last_mouse_monitor
        try:
            return int(choice.split('(')[-1].replace(')', '').strip())
        except:
            return 1

    def track_mouse_monitor(self):
        while self.is_running:
            try:
                x, y = win32api.GetCursorPos()
                with mss.mss() as sct:
                    for i, mon in enumerate(sct.monitors[1:], start=1):
                        if (mon["left"] <= x < mon["left"] + mon["width"] and
                                mon["top"] <= y < mon["top"] + mon["height"]):
                            self.last_mouse_monitor = i
                            break
            except:
                pass
            time.sleep(1)

    def bg_capture(self):
        interval = 1.0 / self.fps
        with mss.mss() as sct:
            while self.is_running:
                try:
                    idx = self.get_selected_monitor_index()
                    monitor = sct.monitors[idx]
                    screenshot = sct.grab(monitor)

                    img = np.frombuffer(screenshot.bgra, dtype=np.uint8).reshape(
                        screenshot.height, screenshot.width, 4)

                    self.buffer_size = (screenshot.width, screenshot.height)

                    small = cv2.resize(img, (0, 0), fx=self.scale, fy=self.scale,
                                       interpolation=cv2.INTER_LINEAR)
                    self.buffer.append(small)
                    time.sleep(interval)
                except:
                    time.sleep(1)

    def on_number_press(self, value):
        self.accum_seconds = min(self.accum_seconds + value, 600)
        self.c_is_held = True
        mins = self.accum_seconds // 60
        secs = self.accum_seconds % 60
        if mins > 0:
            label = f"{mins}m{secs:02d}s  —  relache C pour sauvegarder"
        else:
            label = f"{secs}s  —  relache C pour sauvegarder"
        self.accum_label.configure(text=label)

    def on_c_release(self):
        if self.c_is_held and self.accum_seconds > 0:
            seconds = self.accum_seconds
            self.accum_seconds = 0
            self.c_is_held = False
            self.accum_label.configure(text="")
            self.save_replay(seconds)
        else:
            self.c_is_held = False
            self.accum_seconds = 0

    def save_replay(self, seconds):
        if len(self.buffer) < 5:
            return

        actual_frames = min(len(self.buffer), seconds * self.fps)
        frames_to_save = list(self.buffer)[-actual_frames:]
        timestamp = datetime.now().strftime("%H%M%S")

        if seconds % 10 == 0 and 10 <= seconds <= 90:
            folder = os.path.join(self.base_dir, f"{seconds} sec")
        else:
            folder = os.path.join(self.base_dir, "Longs clips")

        os.makedirs(folder, exist_ok=True)
        filename = f"clip_{seconds}s_{timestamp}.mp4"
        file_path = os.path.join(folder, filename)
        orig_w, orig_h = self.buffer_size

        def write():
            try:
                self.status_label.configure(text="Sauvegarde en cours...", text_color="#e67e22")
                fourcc = cv2.VideoWriter_fourcc(*'mp4v')
                out = cv2.VideoWriter(file_path, fourcc, self.fps, (orig_w, orig_h))
                for f in frames_to_save:
                    full = cv2.resize(f, (orig_w, orig_h), interpolation=cv2.INTER_LINEAR)
                    bgr = cv2.cvtColor(full, cv2.COLOR_BGRA2BGR)
                    out.write(bgr)
                out.release()
                if seconds >= 60:
                    m, s = seconds // 60, seconds % 60
                    label = f"Clip {m}m{s:02d}s enregistre"
                else:
                    label = f"Clip {seconds}s enregistre"
                self.status_label.configure(text=label, text_color="#3498db")
            except Exception as e:
                print(f"Erreur : {e}")

        threading.Thread(target=write).start()

    def init_shortcuts(self):
        for i in range(1, 10):
            keyboard.add_hotkey(f'c+{i}', lambda x=i: self.on_number_press(x * 10))
        keyboard.on_release_key('c', lambda _: self.on_c_release())
        keyboard.wait()


if __name__ == "__main__":
    mutex = check_single_instance()
    if mutex is None:
        bring_existing_to_front()
        sys.exit(0)

    app = QuickClip()
    app.mainloop()
