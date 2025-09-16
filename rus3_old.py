import time
import mss
import numpy as np
import cv2
import pyautogui as pg
import ctypes
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
# DPI 対応
ctypes.windll.user32.SetProcessDPIAware()

class ImageClicker:
    def __init__(self, confidence=0.8, max_attempts=5, wait_time=2, log_file="log_buffer.txt"):
        self.confidence = confidence
        self.max_attempts = max_attempts
        self.wait_time = wait_time
        self.log_file = log_file

    def log_match(self, label):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(f"{now}, 記録 {label}\n")
        print(f"ログ記録: {now}, 記録 {label}")

    def click_image_on_monitor(self, monitor, image_path, label=""):
        """指定したモニタ領域で画像を検索してクリック"""
        template = cv2.imread(image_path, cv2.IMREAD_UNCHANGED)
        if template is None:
            print(f"画像が読み込めません: {image_path}")
            return False
        if template.shape[2] == 4:
            template = cv2.cvtColor(template, cv2.COLOR_BGRA2BGR)
        template_h, template_w = template.shape[:2]

        with mss.mss() as sct:
            sct_img = sct.grab(monitor)
            img = np.array(sct_img)
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

        img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

        res = cv2.matchTemplate(img_gray, template_gray, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(res)

        if max_val >= self.confidence:
            click_x = monitor['left'] + max_loc[0] + template_w // 2
            click_y = monitor['top'] + max_loc[1] + template_h // 2

            time.sleep(0.2)
            pg.moveTo(click_x, click_y, duration=0.2)
            pg.doubleClick()
            print(f"画像 {image_path} をクリックしました: ({click_x}, {click_y})")

            # 画像名に応じてラベル設定してログ
            if label:
                self.log_match(label)
            return True
        return False

    def attempt_click_on_monitor(self, monitor, image_path, label=""):
        for attempt in range(self.max_attempts):
            if self.click_image_on_monitor(monitor, image_path, label):
                return True
            else:
                print(f"リトライ {attempt + 1}/{self.max_attempts}")
                time.sleep(self.wait_time)
        print("最大試行回数に達しました。")
        return False


# 使用例
if __name__ == "__main__":
    # 左モニタ取得
    with mss.mss() as sct:
        left_monitor = None
        for m in sct.monitors[1:]:
            if m['left'] < 0:
                left_monitor = m
                break
        if left_monitor is None:
            left_monitor = sct.monitors[1]

    root = tk.Tk()
    root.withdraw()  # メインウィンドウを非表示
    

    clicker = ImageClicker(confidence=0.8, max_attempts=5, wait_time=1)

    maxcnt = 0
    maxlimit = 20

    image_path = r'D:\RustGodess\image\buttle5.PNG'
    # image_path = r'D:\RustGodess\image\black.PNG'
    image_path = r'D:\RustGodess\image\red.PNG'
    images = [
        (r'D:\RustGodess\image\black.png', "黒"),
        (r'D:\RustGodess\image\yellow.png', "黄色"),
        (r'D:\RustGodess\image\red.png', "赤")
    ]
    # 左モニタで画像をクリック
    res= clicker.attempt_click_on_monitor(left_monitor, image_path)
    if res == True:
        while True:
            image_path = r'D:\RustGodess\image\buttle_true.PNG'
            # 左モニタで画像をクリック
            res= clicker.attempt_click_on_monitor(left_monitor, image_path)
            # 必要なら右モニタでもクリック可能
            # clicker.attempt_click_on_monitor(right_monitor)
            # 画像とログラベルを設定
            if res == False:
                messagebox.showinfo("終了", "エラー 終了しました:" + image_path)
            time.sleep(10)
            while True:
                image_path = r'D:\RustGodess\image\tajayu.PNG'
                res= clicker.attempt_click_on_monitor(left_monitor, image_path)
                if res == False:# battler
                    time.sleep(2)
                    break
                time.sleep(2)
            for img_path, label in images:
                res = clicker.attempt_click_on_monitor(left_monitor, img_path, label)
                time.sleep(1)  # 次のチェックまで1秒待機
            if res == False:
                messagebox.showinfo("終了", "エラー 終了しました:" + str(images))
            maxcnt = maxcnt + 1
            if maxcnt > maxlimit:
                messagebox.showinfo("終了", "終了しました")