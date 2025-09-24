import tkinter as tk
import os
from tkinter import messagebox
from tkinter import ttk
import time
import threading
from common import common_utils

from tkinter import messagebox, ttk
import ctypes



if __name__ == "__main__":

    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_int(-1))  # DPI_UNAWARE
    except Exception:
        pass
    login_info = []
    root = tk.Tk()
 
    # vi = app.viewname
    # us = app.username
    user32 = ctypes.windll.user32
    dc = user32.GetDC(0)
    LOGPIXELSX = 88
    dpi_x = ctypes.windll.gdi32.GetDeviceCaps(dc, LOGPIXELSX)
    user32.ReleaseDC(0, dc)
    print("DPI X:", dpi_x)

    user32 = ctypes.windll.user32
    screen_width = user32.GetSystemMetrics(0)  # 横幅
    screen_height = user32.GetSystemMetrics(1) # 高さ

    print(f"画面解像度: {screen_width} x {screen_height}")
