import customtkinter as ctk
import win32gui
import win32con
import win32process
import win32api
import psutil
import keyboard
import threading
import time
import os
import json
import tkinter as tk
from tkinter import simpledialog
import pystray
from PIL import Image
import sys

CONFIG_FILE = "config.json"
APP_VERSION = "v0.1.0"

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def suppress_callback_exception(*args):
    pass

tk.Tk.report_callback_exception = staticmethod(suppress_callback_exception)

class FullscreenApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Borderless Fullscreen Mode")
        self.geometry("450x410")
        icon_path = resource_path("bfm.ico")
        self.iconbitmap(icon_path)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.protocol("WM_DELETE_WINDOW", self.hide_to_tray)

        self.target_exe = "Game.exe"
        self.selected_title = None
        self._after_id = None
        self.auto_refresh_active = False

        self.is_fullscreen = False
        self.current_hwnd = None
        self.original_style = None
        self.original_exstyle = None
        self.original_rect = None
        self.original_window_size = None
        self.blackbars = []

        font_setting = ("맑은 고딕", 15, "bold")
        small_font = ("맑은 고딕", 13, "bold")

        # 핫키 입력창 (좌측 상단으로 이동)
        self.hotkey_entry = ctk.CTkEntry(self, width=50, height=26, font=("맑은 고딕", 10))
        self.hotkey_entry.insert(0, "F5")
        self.hotkey_entry.place(x=10, y=10)
        self.hotkey_entry.bind("<FocusOut>", lambda e: self.save_config())

        self.label = ctk.CTkLabel(self, text="Borderless Fullscreen Mode", font=("맑은 고딕", 20, "bold"))
        self.label.pack(pady=20)

        self.exe_entry = ctk.CTkEntry(self, placeholder_text="exe or 창 이름 (예: Game.exe, Game)", font=font_setting, height=32, width=320)
        self.exe_entry.pack(pady=10)
        self.exe_entry.insert(0, self.target_exe)
        self.exe_entry.bind("<FocusOut>", lambda e: self.save_config())

        self.window_select_frame = ctk.CTkFrame(self)
        self.window_select_frame.pack(pady=10)
        self.window_listbox = None
        self.refresh_button = ctk.CTkButton(
            self.window_select_frame,
            text="창 목록 새로고침",
            command=self.refresh_window_list,
            width=240,
            height=32,
            font=small_font,
            fg_color="#2C2C2C",
            hover_color="#3C3C3C"
        )
        self.refresh_button.pack(pady=5)

        self.button = ctk.CTkButton(
            self,
            text=f"BFM 토글 ({self.hotkey_entry.get().upper()})",
            command=self.toggle_fullscreen,
            font=font_setting,
            height=40,
            width=240,
            fg_color="#000000",
            hover_color="#171717"
        )
        self.button.pack(pady=20)

        self.status = ctk.CTkLabel(self, text=f"창을 선택 후 버튼 클릭 또는 {self.hotkey_entry.get().upper()}를 누르세요.", font=small_font, wraplength=400, justify="center")
        self.status.pack(pady=10)

        self.load_config()
        self.refresh_window_list()
        threading.Thread(target=self.listen_hotkey, daemon=True).start()
        threading.Thread(target=self.watch_window_closure, daemon=True).start()

        # 핫키 변경 감지용 타이머
        self.monitor_hotkey_change()

        self.version_label = ctk.CTkLabel(self, text=APP_VERSION)
        self.version_label.place(relx=1.0, rely=0.0, x=-10, y=5, anchor="ne")

    def monitor_hotkey_change(self):
        current_hotkey = self.hotkey_entry.get().strip().upper()
        self.button.configure(text=f"BFM 토글 ({current_hotkey})")
        self.status.configure(text=f"창을 선택 후 버튼 클릭 또는 {current_hotkey}를 누르세요.")
        self.after(1000, self.monitor_hotkey_change)

    def refresh_window_list(self):
        if not self.winfo_exists():
            return

        exe_name = self.exe_entry.get().strip()
        matched = []

        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    proc = psutil.Process(pid)
                    proc_name = proc.name().lower()
                    title = win32gui.GetWindowText(hwnd).lower()
                    if exe_name == "" or exe_name.lower() in title or exe_name.lower() in proc_name:
                        hwnds.append((hwnd, win32gui.GetWindowText(hwnd)))
                except psutil.NoSuchProcess:
                    pass
            return True

        hwnds = []
        win32gui.EnumWindows(callback, hwnds)

        if self.window_listbox:
            self.window_listbox.destroy()

        if hwnds:
            titles = [t for _, t in hwnds]
            self.window_listbox = ctk.CTkOptionMenu(
                self.window_select_frame,
                values=titles,
                command=self.set_selected_title,
                font=("맑은 고딕", 14),
                fg_color="black",
                button_color="black"
            )
            self.window_listbox.pack(pady=2)

            if self.selected_title in titles:
                self.window_listbox.set(self.selected_title)
            else:
                self.selected_title = titles[0]
                self.window_listbox.set(self.selected_title)
                self.save_config()
        else:
            self.window_listbox = ctk.CTkLabel(
                self.window_select_frame,
                text="창 감지 안됨.",
                font=("맑은 고딕", 14),
                text_color="gray"
            )
            self.window_listbox.pack(pady=2)

        if hwnds and self.auto_refresh_active:
            if self._after_id:
                self.after_cancel(self._after_id)
                self._after_id = None
            self.auto_refresh_active = False
        elif not hwnds and not self.auto_refresh_active:
            self.auto_refresh_active = True
            self.auto_refresh_loop()

    def auto_refresh_loop(self):
        if not self.auto_refresh_active:
            return
        self.refresh_window_list()
        self._after_id = self.after(1000, self.auto_refresh_loop)

    def hide_to_tray(self):
        self.withdraw()
        icon_path = resource_path("bfm.ico")
        image = Image.open(icon_path)
        menu = pystray.Menu(
            pystray.MenuItem("BFM 토글", lambda: self.toggle_from_tray()),
            pystray.MenuItem("펼치기", lambda: self.restore_window_gui()),
            pystray.MenuItem("종료", lambda: self.quit_app())
        )
        self.tray_icon = pystray.Icon("bfm", image, "BFM", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def toggle_from_tray(self):
        self.after(0, self.toggle_fullscreen)

    def restore_window_gui(self):
        self.deiconify()
        if hasattr(self, 'tray_icon'):
            try:
                self.tray_icon.stop()
            except:
                pass

    def quit_app(self):
        try:
            if self._after_id:
                self.after_cancel(self._after_id)
            for bar in self.blackbars:
                try:
                    bar.destroy()
                except:
                    pass
            self.blackbars.clear()
            if hasattr(self, 'tray_icon'):
                try:
                    self.tray_icon.stop()
                except:
                    pass
            self.destroy()
        except:
            pass
        finally:
            os._exit(0)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.target_exe = data.get("target_exe", "Game.exe")
                self.selected_title = data.get("selected_title")
                if hasattr(self, 'hotkey_entry'):
                    self.hotkey_entry.delete(0, 'end')
                    self.hotkey_entry.insert(0, data.get("hotkey", "F5"))
                if hasattr(self, 'exe_entry'):
                    self.exe_entry.delete(0, 'end')
                    self.exe_entry.insert(0, self.target_exe)

    def save_config(self):
        config = {
            "target_exe": self.exe_entry.get().strip(),
            "selected_title": self.selected_title,
            "hotkey": self.hotkey_entry.get().strip()
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f)

    def set_selected_title(self, value):
        self.selected_title = value
        self.save_config()

    def find_window_by_exe(self):
        exe_name = self.exe_entry.get().strip()
        selected_title = self.selected_title

        matched_hwnds = []

        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    proc = psutil.Process(pid)
                    proc_name = proc.name().lower()
                    title = win32gui.GetWindowText(hwnd).lower()
                    if exe_name == "" or exe_name.lower() in title or exe_name.lower() in proc_name:
                        title = win32gui.GetWindowText(hwnd)
                        if title == selected_title:
                            hwnds.append(hwnd)
                except psutil.NoSuchProcess:
                    pass
            return True

        win32gui.EnumWindows(callback, matched_hwnds)
        return matched_hwnds[0] if matched_hwnds else None

    def get_monitor_full_area(self, hwnd):
        monitor = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)
        info = win32api.GetMonitorInfo(monitor)
        full_area = info['Monitor']
        width = full_area[2] - full_area[0]
        height = full_area[3] - full_area[1]
        return width, height, full_area[0], full_area[1]

    def apply_borderless_fullscreen(self, hwnd):
        self.original_style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        self.original_exstyle = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        rect = win32gui.GetClientRect(hwnd)
        self.original_rect = win32gui.GetWindowRect(hwnd)
        win_w = rect[2] - rect[0]
        win_h = rect[3] - rect[1]
        self.original_window_size = (win_w, win_h)

        aspect_ratio = win_w / win_h
        full_w, full_h, offset_x, offset_y = self.get_monitor_full_area(hwnd)

        if full_w / full_h > aspect_ratio:
            new_height = full_h
            new_width = int(new_height * aspect_ratio)
        else:
            new_width = full_w
            new_height = int(new_width / aspect_ratio)

        x = offset_x + (full_w - new_width) // 2
        y = offset_y + (full_h - new_height) // 2

        win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, win32con.WS_POPUP | win32con.WS_VISIBLE)
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, x, y, new_width, new_height,
                              win32con.SWP_FRAMECHANGED | win32con.SWP_SHOWWINDOW)
        win32gui.UpdateWindow(hwnd)

        self.destroy_blackbars()
        self.create_blackbars(x, y, new_width, new_height, full_w, full_h, offset_x, offset_y)

        self.status.configure(text="BFM 적용됨.")

    def restore_window(self, hwnd):
        if not win32gui.IsWindow(hwnd):
            self.status.configure(text="창 감지 안됨.")
            return

        win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, self.original_style)
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, self.original_exstyle)
        x, y, right, bottom = self.original_rect
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, x, y, right - x, bottom - y,
                              win32con.SWP_FRAMECHANGED | win32con.SWP_SHOWWINDOW)
        self.destroy_blackbars()
        self.status.configure(text="BFM 해제됨.")

    def create_blackbars(self, game_x, game_y, game_w, game_h, monitor_w, monitor_h, offset_x, offset_y):
        left_width = game_x - offset_x
        right_width = (offset_x + monitor_w) - (game_x + game_w)

        if left_width > 0:
            self.blackbars.append(self.make_blackbar(offset_x, offset_y, left_width, monitor_h))
        if right_width > 0:
            self.blackbars.append(self.make_blackbar(game_x + game_w, offset_y, right_width, monitor_h))

    def make_blackbar(self, x, y, width, height):
        bar = ctk.CTk()
        bar.overrideredirect(True)
        bar.geometry(f"{width}x{height}+{x}+{y}")
        bar.configure(bg_color="black", fg_color="black")
        bar.attributes("-topmost", True)
        bar.lift()
        bar.attributes("-topmost", True)
        bar.deiconify()
        bar.update()
        return bar

    def destroy_blackbars(self):
        for bar in self.blackbars:
            try:
                if bar.winfo_exists():
                    bar.destroy()
            except:
                pass
        self.blackbars.clear()

    def toggle_fullscreen(self):
        if not self.is_fullscreen:
            hwnd = self.find_window_by_exe()
            if not hwnd:
                self.status.configure(text="창 감지 안됨.")
                return
            self.current_hwnd = hwnd
            self.target_exe = self.exe_entry.get().strip()
            self.save_config()
            self.apply_borderless_fullscreen(hwnd)
            threading.Thread(target=self.monitor_focus_for_topmost, daemon=True).start()
        else:
            if not self.current_hwnd or not win32gui.IsWindow(self.current_hwnd):
                self.status.configure(text="창 감지 안됨.")
                return
            self.restore_window(self.current_hwnd)

        self.is_fullscreen = not self.is_fullscreen

    def listen_hotkey(self):
        while True:
            hotkey = self.hotkey_entry.get().strip().upper()
            if not hotkey:
                hotkey = "F5"
            try:
                keyboard.wait(hotkey)
                self.after(0, self.toggle_fullscreen)
                while keyboard.is_pressed(hotkey):
                    time.sleep(0.05)
            except:
                time.sleep(1)

    def watch_window_closure(self):
        while True:
            if self.is_fullscreen and self.current_hwnd:
                if not win32gui.IsWindow(self.current_hwnd):
                    self.status.configure(text="창 종료 감지, BFM 해제됨")
                    self.after(0, self.destroy_blackbars)  # 블랙바 제거
                    self.is_fullscreen = False
                    self.current_hwnd = None  # 상태 초기화
            time.sleep(0.5)

    def set_blackbars_topmost(self, topmost: bool):
        for bar in self.blackbars:
            if bar.winfo_exists():
                try:
                    bar.attributes("-topmost", topmost)
                except Exception as e:
                    print(f"블랙바 topmost {'설정' if topmost else '해제'} 오류 (메인 스레드): {e}")
            
    def monitor_focus_for_topmost(self):
        while True:
            if not self.current_hwnd or not win32gui.IsWindow(self.current_hwnd):
                print("게임 창 종료됨, 포커스 감시 루프 종료")
                return

            if self.is_fullscreen:
                foreground_hwnd = win32gui.GetForegroundWindow()
                if foreground_hwnd == self.current_hwnd:
                    # 포커스가 게임에 있으면 최상위 유지
                    left, top, right, bottom = win32gui.GetWindowRect(self.current_hwnd)
                    width = right - left
                    height = bottom - top
                    win32gui.SetWindowPos(self.current_hwnd, win32con.HWND_TOPMOST, left, top, width, height,
                                          win32con.SWP_NOACTIVATE)
                    self.after(0, lambda: self.set_blackbars_topmost(True))
                else:
                    win32gui.SetWindowPos(self.current_hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
                    self.after(0, lambda: self.set_blackbars_topmost(False))
            time.sleep(0.05)

if __name__ == '__main__':
    app = FullscreenApp()
    app.mainloop()
