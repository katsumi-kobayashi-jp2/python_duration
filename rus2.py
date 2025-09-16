import time
import mss
import numpy as np
import cv2
import pyautogui as pg
import ctypes

# DPI 対応
ctypes.windll.user32.SetProcessDPIAware()

class ImageClicker:
    def __init__(self, confidence=0.8, max_attempts=5, wait_time=2):
        self.confidence = confidence
        self.max_attempts = max_attempts
        self.wait_time = wait_time

    def click_image_on_monitor(self, monitor, image_path):
        """指定したモニタ領域で画像を検索してクリック"""
        # テンプレート画像読み込み
        template = cv2.imread(image_path, cv2.IMREAD_UNCHANGED)
        if template is None:
            print(f"画像が読み込めません: {image_path}")
            return False
        if template.shape[2] == 4:  # BGRA → BGR
            template = cv2.cvtColor(template, cv2.COLOR_BGRA2BGR)
        template_h, template_w = template.shape[:2]

        # スクリーンショット取得
        with mss.mss() as sct:
            sct_img = sct.grab(monitor)
            img = np.array(sct_img)
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

        # グレースケール化
        img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

        # テンプレートマッチング
        res = cv2.matchTemplate(img_gray, template_gray, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(res)

        if max_val >= self.confidence:
            click_x = monitor['left'] + max_loc[0] + template_w // 2
            click_y = monitor['top'] + max_loc[1] + template_h // 2

            time.sleep(0.2)
            pg.moveTo(click_x, click_y, duration=0.2)
            pg.doubleClick()
            print(f"画像 {image_path} をクリックしました: ({click_x}, {click_y})")
            return True
        return False

    def attempt_click_on_monitor(self, monitor, image_path):
        """画像パスを引数にしてクリックを試行"""
        for attempt in range(self.max_attempts):
            if self.click_image_on_monitor(monitor, image_path):
                return True
            else:
                print(f"リトライ {attempt + 1}/{self.max_attempts}")
                time.sleep(self.wait_time)
        print("最大試行回数に達しました。")
        return False


# 使用例
if __name__ == "__main__":
    image_path = r'D:\RustGodess\image\buttle5.PNG'

    with mss.mss() as sct:
        # 左モニタの座標取得（負の座標にも対応）
        left_monitor = None
        for m in sct.monitors[1:]:
            if m['left'] < 0:
                left_monitor = m
                break
        if left_monitor is None:
            left_monitor = sct.monitors[1]

    clicker = ImageClicker(confidence=0.8, max_attempts=10, wait_time=1)

    maxcnt = 0
    maxlimit = 20

    image_path = r'D:\RustGodess\image\buttle5.PNG'
    # 左モニタで画像をクリック
    res= clicker.attempt_click_on_monitor(left_monitor, image_path)
    if res == True:
        while True:
            image_path = r'D:\RustGodess\image\buttle_true.PNG'
            # 左モニタで画像をクリック
            res= clicker.attempt_click_on_monitor(left_monitor, image_path)
            # 必要なら右モニタでもクリック可能
            # clicker.attempt_click_on_monitor(right_monitor)
            