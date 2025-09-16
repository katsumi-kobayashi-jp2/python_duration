import time
import mss
import numpy as np
import cv2
import pyautogui as pg
import ctypes
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import random

# DPI 対応
ctypes.windll.user32.SetProcessDPIAware()

class ImageClicker:
    def __init__(self, confidence=0.8, max_attempts=5, wait_time=2, log_file="log_buffer.txt"):
        self.confidence = confidence
        self.max_attempts = max_attempts
        self.wait_time = wait_time
        self.log_file = log_file

    # def log_match(self, label):
    #     now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    #     with open(self.log_file, "a", encoding="utf-8") as f:
    #         f.write(f"{now}, 記録 {label}\n")
    #     print(f"ログ記録: {now}, 記録 {label}")

    # def log_match(self, label):
    #     now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    #     log_file_in = "log_buffer.txt"
    #     log_file_out = "log_buffer2.txt"
    #     if label != '' and label != None:
    #         # 末尾に新しい記録を追加
    #         data = f"{now}, 記録 {label}\n"
    #         # log_buffer2.txt に書き込み
    #         with open(log_file_in, "a", encoding="utf-8") as f:
    #                 f.write(data)
    #         print(f"ログ記録(X): {now}, 記録 {label} + 黒・黄・赤 → {log_file_in}")        
    #         if label == "ｘ":
    #             # 元ファイルを読み込む（存在しない場合は空にする）
    #             try:
    #                 with open(log_file_in, "r", encoding="utf-8") as f:
    #                     data = f.read()
    #             except FileNotFoundError:
    #                 data = ""

    def log_match(self, label):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file_in = "log_buffer.txt"
        log_file_out = "log_buffer2.txt"
        log_file_box = "log_buffer_box.txt"
        if label not in ("", None, "40"):
            # 通常の記録
            with open(log_file_in, "a", encoding="utf-8") as f:
                f.write(f"{now}, 記録 {label}\n")
            print(f"ログ記録: {now}, 記録 {label} → {log_file_in}")

            # 特別処理: "X" のとき
            if label == 'ｘ':
                try:
                    with open(log_file_in, "r", encoding="utf-8") as f:
                        lines = f.readlines()
                except FileNotFoundError:
                    lines = []

                # "記録 X" を含む行を削除
                filtered = [line for line in lines if "記録 ｘ" not in line]

                # log_buffer2.txt に保存
                with open(log_file_out, "w", encoding="utf-8") as f:
                    f.writelines(filtered)

                print(f"→ '記録 X' を削除して {log_file_out} を作成しました")
        elif label not in ("", None):
            with open(log_file_box, "a", encoding="utf-8") as f:
                f.write(f"{now}, 記録 {label}\n")
            print(f"ログ記録: {now}, 記録 {label} → {log_file_in}")            

    def find_image_on_monitor(self, monitor, image_path):
        """クリックせず、画像の有無だけを判定"""
        template = cv2.imread(image_path, cv2.IMREAD_UNCHANGED)
        if template is None:
            print(f"画像が読み込めません: {image_path}")
            return False
        if template.shape[2] == 4:
            template = cv2.cvtColor(template, cv2.COLOR_BGRA2BGR)

        with mss.mss() as sct:
            sct_img = sct.grab(monitor)
            img = np.array(sct_img)
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

        img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

        res = cv2.matchTemplate(img_gray, template_gray, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, _ = cv2.minMaxLoc(res)

        return max_val >= self.confidence

    def click_image_on_monitor(self, monitor, image_path, label="", option=None, offset=(0,0)):
        """指定したモニタ領域で画像を検索してクリック。optionでクリック座標を優先指定可能。offsetで補正可能"""
        offset_x, offset_y = offset

        if option is not None:
            # 座標が指定された場合は、画像検索をスキップしてクリック
            click_x = monitor['left'] + option[0] + offset_x
            click_y = monitor['top'] + option[1] + offset_y
            time.sleep(0.2)
            pg.moveTo(click_x, click_y, duration=0.2)
            pg.doubleClick()
            print(f"オプション座標をクリックしました: ({click_x}, {click_y}) [補正 {offset}]")
            if label:
                self.log_match(label)
            return True

        # 画像マッチング処理
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
            click_x = monitor['left'] + max_loc[0] + template_w // 2 + offset_x
            click_y = monitor['top'] + max_loc[1] + template_h // 2 + offset_y

            time.sleep(0.2)
            pg.moveTo(click_x, click_y, duration=0.2)
            pg.doubleClick()
            print(f"画像 {image_path} をクリックしました: ({click_x}, {click_y}) [補正 {offset}]")

            if label:
                self.log_match(label)
            return True
        return False

    # def attempt_click_on_monitor(self, monitor, image_path, label="", option=None, offset=(0,0)):
    #     """リトライ付きクリック処理。optionがあれば座標優先。offsetで補正可能"""
    #     for attempt in range(self.max_attempts):
    #         if self.click_image_on_monitor(monitor, image_path, label, option, offset):
    #             return True
    #         else:
    #             print(f"リトライ {attempt + 1}/{self.max_attempts}")
    #             time.sleep(self.wait_time)
    #     print("最大試行回数に達しました。")
    #     return False

    def attempt_click_on_monitor(self, monitor, image_path, label="", option=None, offset=(0,0), retry=3):
        """リトライ付きクリック処理。optionがあれば座標優先。offsetで補正可能
        retry: リトライ回数（指定なしなら self.max_attempts を使用）
        """
        attempts = retry if retry is not None else self.max_attempts

        for attempt in range(attempts):
            if self.click_image_on_monitor(monitor, image_path, label, option, offset):
                return True
            else:
                print(f"リトライ {attempt + 1}/{attempts}")
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
    xxx = 0 #1:ON  0:OFF　空の時のみ　ふたも開ける

    zzz = 1 #すべて空あける

    yyy = 1 #買いキャンセル

    vvv = 1 #Victory

    maxlimit = 44
    timespan = random.uniform(1.1, 2.2) 
    timespan_10 = random.uniform(1.5, 5.5) 
    timespan_20 = random.uniform(5.2, 30.5) 
    time.sleep(timespan-1)
    offx = random.uniform(0.5, 7.3)
    offy = random.uniform(0.2, 5.5)


    # if 1==zzz:
    #         image_path = r'D:\RustGodess\image\open_box_black.PNG'
    #         # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
    #         res = clicker.find_image_on_monitor(left_monitor, image_path)
    #         if res == True:
    #             time.sleep(timespan_20)
    #             image_path = r'D:\RustGodess\image\open_box_black.PNG'
    #             res = clicker.attempt_click_on_monitor(left_monitor, image_path, label = "40" ,retry=2)
    #             time.sleep(timespan)  # 次のチェックまで1秒待機            
    #             if res== False:                
    #                 messagebox.showinfo("終了", "エラー終了しました")
    #                 exit()
    #             time.sleep(timespan)
    #             image_path = r'D:\RustGodess\image\box_move.PNG'
    #             res = clicker.attempt_click_on_monitor(left_monitor, image_path, label = "40",option =(500,600) ,retry=2)
    #             time.sleep(timespan)  # 次のチェックまで1秒待機            
    #             if res== False:                
    #                 messagebox.showinfo("終了", "エラー終了しました")
    #                 exit()                    




    images = [
        (r'D:\RustGodess\image\black.png', "."),
        (r'D:\RustGodess\image\yellow.png', "黄"),
        (r'D:\RustGodess\image\red.png', "赤")
    ]
    if 1==1:
        image_path = r'D:\RustGodess\image\buy_ok.PNG'
        # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
        res = clicker.find_image_on_monitor(left_monitor, image_path)
        if res == True:
            time.sleep(timespan)

            image_path = r'D:\RustGodess\image\buy_ok_X.PNG'
            res = clicker.attempt_click_on_monitor(left_monitor, image_path, retry=1)
            time.sleep(timespan)  # 次のチェックまで1秒待機            
            if res== False:                
                messagebox.showinfo("終了", "エラー終了しました")
                exit()
    image_path = r'D:\RustGodess\image\buttle5.PNG'
    # image_path = r'D:\RustGodess\image\black.PNG'
    # image_path = r'D:\RustGodess\image\red.PNG'
    # 左モニタで画像をクリック    
    res = clicker.attempt_click_on_monitor(left_monitor, image_path, offset=(offx, offy), retry=2)
    # res = clicker.attempt_click_on_monitor(left_monitor, image_path,option =(500,600))
    if res == True:
        while True:

            if 1==yyy: #買い
                time.sleep(timespan) 
                image_path = r'D:\RustGodess\image\buy_ok.PNG'
                # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
                res = clicker.find_image_on_monitor(left_monitor, image_path)
                if res == True:
                    time.sleep(timespan)

                    image_path = r'D:\RustGodess\image\buy_ok_X.PNG'
                    res = clicker.attempt_click_on_monitor(left_monitor, image_path, retry=2)
                    time.sleep(timespan)  # 次のチェックまで1秒待機            
                    if res== False:                
                        messagebox.showinfo("終了", "エラー終了しました")
                        exit()  

            if 1==1:
                time.sleep(timespan) 
                image_path = r'D:\RustGodess\image\4499.PNG'
                # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
                res = clicker.find_image_on_monitor(left_monitor, image_path)
                if res == True:
                    time.sleep(timespan)

                    image_path = r'D:\RustGodess\image\buy_ok_X.PNG'
                    res = clicker.attempt_click_on_monitor(left_monitor, image_path, retry=1)
                    time.sleep(timespan)  # 次のチェックまで1秒待機            
                    if res== False:                
                        messagebox.showinfo("終了", "エラー終了しました")
                        exit()



            # if vvv==1:
            #     time.sleep(timespan) 
            #     image_path = r'D:\RustGodess\image\toribo.PNG'
            #     # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
            #     res = clicker.find_image_on_monitor(left_monitor, image_path)
            #     if res == True:
            #         time.sleep(timespan_20)
            #         image_path = r'D:\RustGodess\image\toribo.PNG'
            #         res = clicker.attempt_click_on_monitor(left_monitor, image_path, retry=3)
            #         time.sleep(timespan)  # 次のチェックまで1秒待機            
            #         if res== False:                
            #             messagebox.showinfo("終了", "エラー終了しました")
            #             exit()
            #         time.sleep(timespan)
            #         image_path = r'D:\RustGodess\image\box_move.PNG'
            #         res = clicker.attempt_click_on_monitor(left_monitor, image_path, label = "40",option =(500,600) ,retry=2)
            #         time.sleep(timespan)  # 次のチェックまで1秒待機            
            #         if res== False:                                
            #             messagebox.showinfo("終了", "エラー終了しました")
            #             exit()
            if vvv==1:
                time.sleep(timespan) 
                image_path = r'D:\RustGodess\image\vikutori.PNG'
                # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
                res = clicker.find_image_on_monitor(left_monitor, image_path)
                if res == True:
                    time.sleep(timespan_20)
                    image_path = r'D:\RustGodess\image\vikutori.PNG'
                    res = clicker.attempt_click_on_monitor(left_monitor, image_path, offset=(offx, offy-75 ), retry=3)
                    time.sleep(timespan)  # 次のチェックまで1秒待機            
                    if res== False: 
                        messagebox.showinfo("終了", "エラー終了しました")
                        exit()

                    time.sleep(timespan) 
                    image_path = r'D:\RustGodess\image\4499.PNG'
                    # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
                    res = clicker.find_image_on_monitor(left_monitor, image_path)
                    if res == True:
                        time.sleep(timespan)
                        image_path = r'D:\RustGodess\image\buy_ok_X.PNG'
                        res = clicker.attempt_click_on_monitor(left_monitor, image_path, retry=1)
                        time.sleep(timespan)  # 次のチェックまで1秒待機            
                        if res== False:                
                            messagebox.showinfo("終了", "エラー終了しました")
                            exit()  

                        # time.sleep(timespan) 
                        # image_path = r'D:\RustGodess\image\vikutori.PNG'
                        # # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
                        # res = clicker.find_image_on_monitor(left_monitor, image_path)
                        # if res == True:
                        #     time.sleep(timespan_20)
                        #     image_path = r'D:\RustGodess\image\vikutori.PNG'
                        #     res = clicker.attempt_click_on_monitor(left_monitor, image_path, offset=(offx, offy-75 ), retry=3)
                        #     time.sleep(timespan)  # 次のチェックまで1秒待機            
                        #     if res== False: 
                        #         messagebox.showinfo("終了", "エラー終了しました")
                        #         exit()

                    time.sleep(timespan)
                    image_path = r'D:\RustGodess\image\box_move.PNG'
                    res = clicker.attempt_click_on_monitor(left_monitor, image_path, label = "40",option =(500,600) ,retry=2)
                    time.sleep(timespan)  # 次のチェックまで1秒待機            
                    if res== False:                
                        messagebox.showinfo("終了", "エラー終了しました")
                        exit()

            if zzz==1:
                time.sleep(timespan) 
                image_path = r'D:\RustGodess\image\open_box_black.PNG'
                # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
                res = clicker.find_image_on_monitor(left_monitor, image_path)
                if res == True:
                    time.sleep(timespan_20)
                    image_path = r'D:\RustGodess\image\open_box_black.PNG'
                    res = clicker.attempt_click_on_monitor(left_monitor, image_path, retry=3)
                    time.sleep(timespan)  # 次のチェックまで1秒待機            
                    if res== False:                
                        messagebox.showinfo("終了", "エラー終了しました")
                        exit()
                    time.sleep(timespan)
                    image_path = r'D:\RustGodess\image\box_move.PNG'
                    res = clicker.attempt_click_on_monitor(left_monitor, image_path, label = "40",option =(500,600) ,retry=2)
                    time.sleep(timespan)  # 次のチェックまで1秒待機            
                    if res== False:                
                        messagebox.showinfo("終了", "エラー終了しました")
                        exit()
                image_path = r'D:\RustGodess\image\open_box_red.PNG'
                # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
                res = clicker.find_image_on_monitor(left_monitor, image_path)
                if res == True:
                    time.sleep(timespan_20)
                    image_path = r'D:\RustGodess\image\open_box_red.PNG'
                    res = clicker.attempt_click_on_monitor(left_monitor, image_path, retry=3)
                    time.sleep(timespan)  # 次のチェックまで1秒待機            
                    if res== False:                
                        messagebox.showinfo("終了", "エラー終了しました")
                        exit()
                    time.sleep(timespan)
                    image_path = r'D:\RustGodess\image\box_move.PNG'
                    res = clicker.attempt_click_on_monitor(left_monitor, image_path, label = "40",option =(500,600) ,retry=2)
                    time.sleep(timespan)  # 次のチェックまで1秒待機            
                    if res== False:                
                        messagebox.showinfo("終了", "エラー終了しました")
                        exit()                        

            if 1==xxx :
                time.sleep(timespan) 
                while True: #空のビコトリBOXあり
                    image_path = r'D:\RustGodess\image\BOX_emp.PNG'
                    # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
                    res = clicker.find_image_on_monitor(left_monitor, image_path)
                    if res == True:
                        time.sleep(timespan)
                        break                    
                    if 1==1:
                        image_path = r'D:\RustGodess\image\open_box_black.PNG'
                        # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
                        res = clicker.find_image_on_monitor(left_monitor, image_path)
                        if res == True:
                            time.sleep(timespan_20)
                            image_path = r'D:\RustGodess\image\open_box_black.PNG'
                            res = clicker.attempt_click_on_monitor(left_monitor, image_path, retry=2)
                            time.sleep(timespan)  # 次のチェックまで1秒待機            
                            if res== False:                
                                messagebox.showinfo("終了", "エラー終了しました")
                                exit()
                            time.sleep(timespan)
                            image_path = r'D:\RustGodess\image\box_move.PNG'
                            res = clicker.attempt_click_on_monitor(left_monitor, image_path, label = "40",option =(500,600) ,retry=2)
                            time.sleep(timespan)  # 次のチェックまで1秒待機            
                            if res== False:                
                                messagebox.showinfo("終了", "エラー終了しました")
                                exit()


            if 1==yyy: #買い
                time.sleep(timespan) 
                image_path = r'D:\RustGodess\image\buy_ok.PNG'
                # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
                res = clicker.find_image_on_monitor(left_monitor, image_path)
                if res == True:
                    time.sleep(timespan)

                    image_path = r'D:\RustGodess\image\buy_ok_X.PNG'
                    res = clicker.attempt_click_on_monitor(left_monitor, image_path, retry=2)
                    time.sleep(timespan)  # 次のチェックまで1秒待機            
                    if res== False:                
                        messagebox.showinfo("終了", "エラー終了しました")
                        exit() 

            time.sleep(timespan)  
            image_path = r'D:\RustGodess\image\reage_down.PNG' #リーグダウン
            res = clicker.attempt_click_on_monitor(left_monitor, image_path, retry=1)
            time.sleep(timespan)  # 次のチェックまで1秒待機            
            # if res== False:                
            #     messagebox.showinfo("終了", "エラー終了しました")
            #     exit()   

            timespan = random.uniform(1.5, 2.5) 
            time.sleep(timespan)
            # timespan = 2
            image_path = r'D:\RustGodess\image\buttle_true.PNG'
            # 左モニタで画像をクリック
            res = clicker.attempt_click_on_monitor(left_monitor, image_path, offset=(offx, offy))
            if res == False:
                messagebox.showinfo("終了", "エラー 終了しました:" + image_path)
                exit()
            time.sleep(10)

            if 1==1:
                time.sleep(timespan) 
                image_path = r'D:\RustGodess\image\4499.PNG'
                # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
                res = clicker.find_image_on_monitor(left_monitor, image_path)
                if res == True:
                    time.sleep(timespan)

                    image_path = r'D:\RustGodess\image\buy_ok_X.PNG'
                    res = clicker.attempt_click_on_monitor(left_monitor, image_path, retry=1)
                    time.sleep(timespan)  # 次のチェックまで1秒待機            
                    if res== False:                
                        messagebox.showinfo("終了", "エラー終了しました")
                        exit()


            while True:
                image_path = r'D:\RustGodess\image\tajayu.PNG'
                # res = clicker.attempt_click_on_monitor(left_monitor, image_path)
                res = clicker.find_image_on_monitor(left_monitor, image_path)
                if res == False:
                    time.sleep(timespan)
                    break
                time.sleep(timespan)
            time.sleep(timespan*3)
            offx = random.uniform(0.5, 7.3)
            offy = random.uniform(0.2, 5.5)            
            image_path = r'D:\RustGodess\image\meisei_down.PNG'
            label = 'ｘ'
            res = clicker.attempt_click_on_monitor(left_monitor, image_path, label, offset=(offx, offy), retry=2)
            time.sleep(timespan)  # 次のチェックまで1秒待機            
            if res== False:
                for img_path, label in images:
                    res = clicker.attempt_click_on_monitor(left_monitor, img_path, label,option =(500,600), offset=(offx, offy), retry=2)
                    time.sleep(timespan)  # 次のチェックまで1秒待機
                    if res==True:
                        break
                if res == False:
                    messagebox.showinfo("終了", "エラー 終了しました:" + str(images))
                    exit()



            maxcnt = maxcnt + 1
            if maxcnt > maxlimit:
                messagebox.showinfo("終了", "終了しました")
                exit()
            time.sleep(timespan_10)