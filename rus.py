import pyautogui as pg
import time
import mss
import numpy as np
import cv2
class ImageClicker:
    def __init__(self, image_path, confidence=0.5, max_attempts=5, wait_time=2):
        self.image_path = image_path
        self.confidence = confidence
        self.max_attempts = max_attempts
        self.wait_time = wait_time

    def click_image(self):
        try:
            # 画像の位置を検出
            location = pg.locateOnScreen(self.image_path, confidence=self.confidence)
            if location:
                # 画像の中心座標を取得
                center = pg.center(location)
                # クリック
                time.sleep(3)
                pg.doubleClick(center)
                print(f"画像 {self.image_path} {location}をクリックしました。")
                return True
        except pg.ImageNotFoundException:
            print(f"画像 {self.image_path} が見つかりませんでした。")
        return False

    def attempt_click(self):
        for attempt in range(self.max_attempts):
            if self.click_image():
                return True
            else:
                print(f"リトライ {attempt + 1}/{self.max_attempts}")
                time.sleep(self.wait_time)
        print("最大試行回数に達しました。")
        return False

# 使用例
if __name__ == "__main__":
    image = r'D:\RustGodess\image\buttle5.PNG'
    # button_clicker = ImageClicker(r'D:\RustGodess\image\buttle.PNG')
    # button_clicker.attempt_click()

    # # カスタムパラメータを使用した例
    # # サブ画面②（左モニタ）
    # screenshot_sub = pg.screenshot(region=(-1920, 0, 1920, 1080))
    # screenshot_sub.save("sub.png")

    # # メイン画面①（右モニタ）
    # screenshot_main = pg.screenshot(region=(0, 0, 1920, 1080))
    # screenshot_main.save("main.png")

    # with mss.mss() as sct:
    #     for i, monitor in enumerate(sct.monitors[1:], start=1):
    #         img = np.array(sct.grab(monitor))[:, :, :3]
    #         cv2.imwrite(f"monitor_{i}.png", img)    
    custom_clicker = ImageClicker(image, confidence=0.5, max_attempts=10, wait_time=1)
    custom_clicker.click_image()
    # custom_clicker.attempt_click()
