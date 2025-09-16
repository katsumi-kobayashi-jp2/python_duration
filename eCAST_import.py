# seleniumの必要なライブラリをインポート
# from selenium import webdriver
from selenium.webdriver.common.by import By

# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select

# tkinter（メッセージボックス表示）の必要なライブラリをインポート
import tkinter
from tkinter import messagebox
from tkinter import ttk
import time
import sys
import datetime
import os
import pyautogui as pg
import tkinter as tk
from tkinter import E, messagebox
import win32gui
import pywinauto
import ctypes
import win32gui, win32con

from pywinauto import Desktop
from pywinauto.application import Application
from functools import partial
import unicodedata
from unicodedata import east_asian_width

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.text import MIMEText  # メールを日本語で送信
from email.mime.application import MIMEApplication  # メールにファイルを添付

import socket

import webDriverUtils
from common import common_utils

import eCastUtils

import time
import tkinter as tk
from login_app_progress_extend import LoginApp

import re

import glob
import shutil
from datetime import timedelta
import logging

# import uiDriverUtils

# import os
# import datetime
# import io
# import tempfile
# import locale

# from pprint import pprint
# from pprint import PrettyPrinter

# class LoginApp:
#     credentials_file = "credentials.txt"

#     def __init__(self, root):
#         self.root = root
#         self.root.title('eCAST_Export')
#         self.root.geometry('300x320')
#         self.center_window(300, 300)
#         self.root.resizable(False, False)
#         self.root.attributes('-topmost', True)

#         # self.login_info = []
#         self.user_id = None
#         self.password = None
#         self.name = None

#         # Load saved credentials
#         saved_id, saved_password, saved_name = self.load_credentials()

#         self.progress_message = tk.Label(root, text="", font=('Arial', 12))
#         self.progress_message.pack()

#         id_frame = tk.Frame(root, bg='#F1F1F1')
#         id_frame.pack(pady=(10, 5))

#         id_label = tk.Label(id_frame, text='　　　ID:', font=('Arial', 12))
#         id_label.pack(side=tk.LEFT, padx=(10, 5))

#         self.id_entry = tk.Entry(id_frame, font=('Arial', 12), width=14)
#         self.id_entry.insert(tk.END, saved_id)
#         self.id_entry.pack(side=tk.LEFT)

#         password_frame = tk.Frame(root, bg='#F1F1F1')
#         password_frame.pack(pady=5)

#         password_label = tk.Label(password_frame, text='Password:', font=('Arial', 12))
#         password_label.pack(side=tk.LEFT, padx=(10, 5))

#         self.password_entry = tk.Entry(password_frame, font=('Arial', 12), width=14, show="*")
#         self.password_entry.insert(tk.END, saved_password)
#         self.password_entry.pack(side=tk.LEFT)

#         name_frame = tk.Frame(root, bg='#F1F1F1')
#         name_frame.pack(pady=5)

#         name_label = tk.Label(name_frame, text='　　Name:', font=('Arial', 12))
#         name_label.pack(side=tk.LEFT, padx=(10, 5))

#         self.name_entry = tk.Entry(name_frame, font=('Arial', 12), width=14)
#         self.name_entry.insert(tk.END, saved_name)
#         self.name_entry.pack(side=tk.LEFT)

#         submit_button = tk.Button(root, text='Submit', command=self.submit_credentials, bg='#4CAF50', fg='white', font=('Arial', 12))
#         submit_button.pack(pady=(10, 20))

#         self.progress_label = tk.Label(root, text='0/0', font=('Arial', 12))
#         self.progress_label.pack()

#         delete_button = tk.Button(root, text='Delete Files', command=self.delete_files_thread, bg='#F44336', fg='white', font=('Arial', 12))
#         delete_button.pack(pady=(0, 10))

#         style = ttk.Style()
#         style.theme_use('clam')  # デフォルトテーマを使用する（Windowsの場合）
#         # プログレスバーのスタイルをカスタマイズ
#         style.configure('Custom.Horizontal.TProgressbar',
#                         troughcolor='#F5F1F1',  # 背景色
#                         thickness=15,  # 高さ
#                         troughrelief='flat',
#                         troughborderwidth=0)

#         # ウィジェットを作成
#         self.progress = ttk.Progressbar(root, orient='horizontal', length=200, mode='determinate', style='Custom.Horizontal.TProgressbar')
#         self.progress.pack(pady=(0, 10))

#     def center_window(self, width, height):
#         screen_width = self.root.winfo_screenwidth()
#         screen_height = self.root.winfo_screenheight()
#         x = (screen_width // 2) - (width // 2)
#         y = (screen_height // 2) - (height // 2)
#         self.root.geometry(f'{width}x{height}+{x}+{y}')

#     def save_credentials(self, user_id, password, name):
#         with open(LoginApp.credentials_file, 'w') as file:
#             file.write(f"{user_id}\n{password}\n{name}")

#     def load_credentials(self):
#         if os.path.exists(LoginApp.credentials_file):
#             with open(LoginApp.credentials_file, 'r') as file:
#                 lines = file.readlines()
#                 if len(lines) >= 3:
#                     return lines[0].strip(), lines[1].strip(), lines[2].strip()
#         return "", "", ""

#     def submit_credentials(self):
#         self.user_id = self.id_entry.get().strip()
#         self.password = self.password_entry.get().strip()
#         self.name = self.name_entry.get().strip()

#         self.save_credentials(self.user_id, self.password, self.name)
#         login_info.append(self.user_id)
#         login_info.append(self.password)
#         login_info.append(self.name)
#         messagebox.showinfo("Submitted", f"ID: {self.user_id}\nPassword: {self.password}\nName: {self.name}")
#         # self.root.destroy()　#画面消去
#         f=f

#     def delete_files_thread(self):
#         threading.Thread(target=self.delete_files).start()

#     def delete_files(self):
#         folder_path = "C:/tmpp"
#         file_list = os.listdir(folder_path)
#         total_files = len(file_list)

#         self.progress['maximum'] = total_files
#         self.progress['value'] = 0

#         self.progress_message.config(text="処理中")
#         for i, file_name in enumerate(file_list):
#             file_path = os.path.join(folder_path, file_name)
#             for _ in  range(1,100000):
#                 for _ in  range(1,10):
#                     file_path = file_path

#             os.remove(file_path)
#             self.progress['value'] += 1
#             progress_text = f"{i+1}/{total_files}"
#             self.progress_label.config(text=progress_text)
#             self.progress_label.update()

#         messagebox.showinfo("Files Deleted", f"All files in {folder_path} have been deleted.")

tx_nowtime = datetime.datetime.today().strftime("%Y%m%d_%H%M%S.%f")[:-4]
tx_nowdate = datetime.datetime.today().strftime("%Y-%m-%d")

# wait_t = constants.wait_t          #WEB要素関数内のタイミング待ち時間
# tomer_wait = constants.tomer_wait  #タイマー関数有効時の待ち時間
# dom_wait = constants.dom_wait      #DOM待ち関数内のタイムアウト時間
# im_wait = constants.im_wait        #implicitly_wait関数設定値

import os
import glob
import threading
import json

global login_info


def run_gui():
    global root, app
    root = tk.Tk()
    app = LoginApp(root, login_info, "eCAST_import", True)
    # login_info = LoginApp.login_info
    root.mainloop()


def mail_result(login_info, list_cyouhiyou, exists_uploadFile ,cu, error_txt):
    if len(login_info) > 3 and len(list_cyouhiyou) > 0:
        if login_info[3].strip() != "":
            all_cyouhilyou = str(list_cyouhiyou)
            you_list = []
            items = login_info[3].strip().split(":")
            you_list.extend(items)
            you_list.append("katsumi.kobayashi.hg@fujifilm.com")
            # you_list = ["katsumi.kobayashi.hg@fujifilm.com"]
            me = "katsumi.kobayashi.hg@fujifilm.com"  # 送信元メールアドレス satoshi.nakano.mp@fujifilm.com
            title = "eCASTインポート結果"
            txt = all_cyouhilyou
            if error_txt != "":
                txt = all_cyouhilyou + ": " + error_txt
            file_list = []
            cu.send_mail(me, you_list, title, txt, file_list)
            if "中途失敗で、登録者名が変更されていません" in error_txt or "アップロード時に失敗でした。保存はされていません" in error_txt:
                #画面保存
                fpath= os.path.join(loggingDir, tx_nowdate)
                if os.path.exists(fpath) == False:
                    os.mkdir(fpath)
                if os.path.exists(fpath) == True:
                    fpng = os.path.join(fpath, tx_nowtime) + ".png"
                    try:
                        wD.save_scshot(fpng)
                    except Exception as e:
                        pass

loggingDir = r"\\172.22.166.106\230_契約書兼印紙管理システム\earth_データ\log"

if __name__ == "__main__":

    login_info = []
    # 元の文字列
    # original_string = "　 中野　　 智支 　 "  # 例：前後や中にも全角および半角スペースが含まれている
    # URL：https://fbecast.crm7.dynamics.com
    # ID：ecastprod@000.fujifilm.com
    # PW：Keiyaku20210304

    # 空白を取り除く
    # cleaned_string = re.sub(r'\s+', '', original_string)
    # def run_gui():
    #     global root, app
    #     root = tk.Tk()
    #     app = LoginApp(root)
    #     root.mainloop()

    # root = tk.Tk()
    # app = LoginApp(root)
    thread = threading.Thread(target=run_gui)
    thread.daemon = True
    thread.start()
    # root.mainloop()
    # thread.join()  # mainloopが終了する前にスレッドが終わっていることを確認する

 
    while True:
        while len(login_info) < 2:
            time.sleep(1)
        if (
            len(login_info[0].strip()) < 7
            or len(login_info[1].strip()) < 10
            or len(login_info[2].strip()) < 2
        ):
            app.progress_message.config(text="入力欄が間違いです")
        else:
            break

    # 設定を読み込む
    app.load_settings()  # 設定ファイルからデータをロード    
    
    args = sys.argv
    host = socket.gethostname()
    ip = socket.gethostbyname(host)

    # Chrome Webドライバー の インスタンスを生成
    dirname = os.path.dirname(__file__)

    # fxsc00568"00471Ktxaaa

    args = sys.argv
    host = socket.gethostname()
    ip = socket.gethostbyname(host)
    username = os.getlogin()

    # Chrome Webドライバー の インスタンスを生成
    dirname = os.path.dirname(__file__)

    # 使用していない　uiDriverUtils用
    # PaneName='アフター契約条件未確定一覧 -- Finance and Operations'
    # PaneName2 = PaneName + ' - Google Chrome'

    # 遅延の有効有無
    tomerOn = False
    # DOM処理クラス
    headless = app.get_headless_status()
    wD = webDriverUtils.DOMProcessor(False, headless)  #
    click_retry = wD.click_retry
    # 共通クラス
    cu = common_utils.com_utility()
    cs = cu.cs_
    loggingDir = loggingDir + "/" + username
    if os.path.exists(loggingDir) == False:
        os.mkdir(loggingDir)
    custom_logger = common_utils.CustomLogger(
        name="eCAST_logger", base_log_dir=loggingDir, level=logging.INFO
    )

    # if os.path.exists("./log") == False:
    #     os.mkdir("./log")
    # custom_logger = common_utils.CustomLogger(name="eCAST_logger", base_log_dir='log', level=logging.INFO)
    # list_cyouhiyou = ["yuujj"]
    # error_txt = "tesut"
    # mail_result(login_info, list_cyouhiyou, cu, error_txt)
    # if len(login_info) > 3:
    #     if login_info[3].strip() != "":
    #         you_list = []
    #         items = login_info[3].strip().split(":")
    #         you_list.extend(items)
    #         you_list.append("katsumi.kobayashi.hg@fujifilm.com")
    #         me = "katsumi.kobayashi.hg@fujifilm.com"  # 送信元メールアドレス satoshi.nakano.mp@fujifilm.com
    #         title = "eCASTインポート結果"
    #         txt = "[fxsc67777,fxsc88777]"
    #         file_list = []
    #         cu.send_mail(me, you_list, title, txt, file_list)

    # fpath= os.path.join(loggingDir, tx_nowdate)
    # if os.path.exists(fpath) == False:
    #     os.mkdir(fpath)
    # if os.path.exists(fpath) == True:
    #     fpng = os.path.join(fpath, tx_nowtime) + ".png"
    #     try:
    #         wD.save_scshot(fpng)
    #     except Exception as e:
    #         pass


    logger = custom_logger.get_logger()
    # エクスポートとダウンロード用フォルダ
    download_folder = os.path.expanduser("~\Downloads")

    # scaling_factor = cu.get_scaling_factor()
    # ログメッセージを記録
    # logger.debug("これはデバッグメッセージです")
    # logger.info("これは情報メッセージです")
    # logger.warning("これは警告メッセージです")
    # logger.error("これはエラーメッセージです")
    # logger.critical("これはクリティカルメッセージです")

    # データベーステスト　************************************************************************
    script_name = os.path.basename(__file__)
    pathDB = r"\\172.22.166.106\230_契約書兼印紙管理システム\appData\eCAST_usage_data.db"
    pathFolder = os.path.dirname(pathDB)
    access_rights = cu.check_access_rights(pathFolder)
    if access_rights["readable"] == False:
        txt = "アクセス件がありません：" + pathFolder
        logger.info(txt)
        app.progress_message.config(text=txt + "終了してください")
        threading.Event().wait()

    db = common_utils.AppUsageDB(pathDB, script_name, username, "", ip)
    # ', app_name ='', user_id ='', user_name ='', ip =''):
    db.connect()
    db.create_table()

    # サンプルデータを10件挿入

    # for i in range(10):
    #     db.data[2]= 20 #*20# [123.45 + j for j in range(20)]  # data1 to data20の数値データを生成
    #     db.texts[5] = "dfg" #['']*10 #[f"テキスト_{j+1}" for j in range(10)]  # text1 to text10のテキストデータを生成
    #     db.insert_data(f"アプリA_{i+1}", f"user_{i+1:03d}", f"名前_{i+1}", f"IP_{i+1}", 150 + i, db.data, db.texts)

    if "readfile" != "readfile":
        # データをJSON形式でエクスポート
        json_file_path = "app_usage_data.json"
        db.export_to_json(json_file_path)

        db.close()

        # JSONファイルを読み込み表示
        with open(json_file_path, "r", encoding="utf-8") as jsonfile:
            json_data = json.load(jsonfile)
            print(json.dumps(json_data, ensure_ascii=False, indent=4))

        sys.exit()

    # 　*****************************************************************************************

    # dest_folder = 'c:/testtmp' #\\172.22.166.100\c020_請求業務\010_請求業務G\200_請求業務G業務別\010_契約書作成\010_契約書（作成済分）\契約書作成チーム\〇eCast\☆2024LH_ＴＳＣ5年満了
    dest_folder = r"\\172.22.166.100\c020_請求業務\010_請求業務G\200_請求業務G業務別\010_契約書作成\010_契約書（作成済分）\契約書作成チーム\〇eCast\☆2024LH_ＴＳＣ5年満了"
    # dest_folder = "c:\\temp"
    access_rights = cu.check_access_rights(dest_folder)
    if access_rights["readable"] == False:
        txt = "アクセス件がありません：" + dest_folder
        logger.info(txt)
        app.progress_message.config(text=txt + "終了してください")
        threading.Event().wait()

    # download_folder = os.path.expanduser("~\Downloads")
    # elapsed_time = datetime.datetime.now()
    # result = cu.copy_and_delete_files(download_folder, dest_folder, elapsed_time, 20)

    # PaneName='アフター契約条件未確定一覧 -- Finance and Operations'
    # PaneName2 = PaneName + ' - Google Chrome'

    # txt = "正常終了しました。fxscのみ抽出しています。\n\n" + "処理予定数：" + str(5)  + "\n処理数：" + str(6) + "\n画面全数：" + "4"
    # txt = "終了:処理数が不足しています。fxscのみ抽出しています。\n\n" + "処理予定数：" + str(3)  + "\n処理数：" + str(5) + "\n画面全数：" + "7"

    # db.insert_data(db.app_name, db.user_id, db.user_name, db.ip, 6, db.data, db.texts)

    # app.progress_message.config(text=txt)
    # messagebox.showinfo("alert",txt)

    tomerOn = False

    # # 無限に待機
    # threading.Event().wait()
    txt = "処理開始"
    logger.info(txt)
    # lok = app.viewname
    app.progress_message.config(text=txt)
    wD.start_addr = "https://fbecast.crm7.dynamics.com/main.aspx?appid=e20bff65-5605-4957-bca1-f5b3bb7042a7&forceUCI=1&pagetype=dashboard&id=5ada4222-2a23-eb11-a813-00224867705c&type=system&_canOverride=true"
    kensaku_box = '//div[contains(@id, "modulesPaneOpener")]'
    # 検索窓が出るまで待機
    if wD.driver_open(kensaku_box, cs.im_wait, tomerOn) == False:
        txt = "ブラウザ起動に失敗しました。"
        logger.info(txt)
        app.progress_message.config(text=txt + "終了してください")
        threading.Event().wait()
    # send_mail('katsumi.kobayashi.hg@fujifilm.com',['katsumi.kobayashi.hg@fujifilm.com'],"エラー" + txt,txt +  " IP:"+ str(ip) ,[''])
    # LOAD待ち

    # earth処理クラス
    data = [[], []]
    eu = eCastUtils.eCAST_utility(wD, cs, data, tomerOn)
    size = wD.driver.get_window_size()
    # app.progress_message.config(text="ブラウザ起動に失敗しました。5秒後に終了します")
    # ウィンドウを最小化
    # wD.driver.minimize_window()
    # ウィンドウを最大化
    # wD.driver.maximize_window()
    # ウインドウを固定サイズにする
    # wD.driver.set_window_size(size['width'], size['height']) # サイズを設定 (任意のサイズに調整可)

    #  サイズを設定 (任意のサイズに調整可)
    wD.driver.set_window_size(1654, size["height"])  #

    # ログイン
    txt = "ログイン中"
    logger.info(txt)
    app.progress_message.config(text=txt)
    LOGIN_OK = ""
    locator = "//div[@id ='LPCPreLoadFontIconContainer']"
    elements = wD.dom_elements(locator, 5, cs.tomer_wait, tomerOn)
    if "False" in elements:
        LOGIN_OK = eu.eCAST_LOGIN_CHECK(login_info, 2, 2)
        if "Fals" in LOGIN_OK:
            txt = "タイムアウトしました:" + locator
            logger.info(txt)
            app.progress_message.config(text=txt + "終了してください")
            threading.Event().wait()
    # 情報保存

    # result = wD.dom_output(dirname +"/eCAST_All_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    # if result == 'False':
    #     result = wD.dom_output(dirname +"/eCAST_All_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #     if result == 'False':
    #         messagebox.showinfo("alert","タイムアウトしました")

    # 写真部分押下
    # locator = '//button[contains(@id, "mectrl_main_trigger")]'
    # elements = wD.dom_elements(locator,15,cs.tomer_wait,tomerOn)
    # if 'False' in elements:
    #     messagebox.showinfo("alert","タイムアウトしました")
    # elements[0].click()

    # すでにログインしていない場合
    if not ("True" in LOGIN_OK):
        while True:
            result = eu.eCAST_LOGIN(login_info, 10, 2)
            if "ユーザーIDが間違っています" in result:
                logger.info(result)
                app.progress_message.config(text=result)
            elif result != "True":
                logger.info(result)
                app.progress_message.config(text=result)
            else:
                break

    # LOAD待ち
    locator = "//div[@id ='LPCPreLoadFontIconContainer']"
    elements = wD.dom_elements(locator, 15, cs.tomer_wait, tomerOn)
    if "False" in elements:
        txt = "タイムアウトしました" + locator
        logger.info(txt)
        app.progress_message.config(text=txt + "終了してください")
        threading.Event().wait()
        # messagebox.showinfo("alert","タイムアウトしました:" + locator)

    # if 1==2 : #実行しない eCAST_LOGINの内容と同じ
    #     #すでに、ログイン済みかを判断  何故かtextで取得できない
    #     locator = '//div[contains(@id, "mectrl_currentAccount_primary")]'
    #     dom_wait_r = 5
    #     user_name = wD.dom_elements_getTextContent(locator, 1, dom_wait_r, cs.tomer_wait, tomerOn)
    #     # # wD.dom_elements(locator, dom_wait_r, cs.tomer_wait, tomerOn)[0].text
    #     # element = wD.driver.find_element(By.CSS_SELECTOR, '.mectrl_accountDetails .mectrl_name.mectrl_truncate')
    #     # name = element.text
    #     # # JavaScriptを使って `innerText` を取得
    #     # innerText_script = """
    #     # return arguments[0].innerText;
    #     # """
    #     # name_innerText = wD.driver.execute_script(innerText_script, element)

    #     # # JavaScriptを使って `textContent` を取得
    #     # textContent_script = """
    #     # return arguments[0].textContent;
    #     # """
    #     # name_textContent = wD.driver.execute_script(textContent_script, element)
    #     if 'False' in user_name:
    #         txt="起動が失敗しています:" + locator
    #         logger.info(txt)
    #         app.progress_message.config(text=txt + "終了してください")
    #         threading.Event().wait()
    #         # messagebox.showinfo("alert","起動が失敗しています:" + locator)
    #         # sys.exit()
    #     else:
    #         # 空白を取り除く
    #         cleaned_user_name = re.sub(r'\s+', '', user_name)
    #         cleaned_login_name = re.sub(r'\s+', '', login_info[2])

    #     #既にログイン済みかどうか判断
    #     if cleaned_user_name != cleaned_login_name:
    #         # 写真をクリック
    #         locator = '//button[contains(@id, "mectrl_main_trigger")]'
    #         elements = wD.safe_click_elements(locator, cs.dom_wait, cs.tomer_wait, tomerOn, click_retry)
    #         if 'False' in elements:
    #             txt="タイムアウトしました:" + locator
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             # messagebox.showinfo("alert","タイムアウトしました:" + locator)

    #         # # ログイン
    #         # locator = '//span[contains(@class, "mectrl_screen_reader_text")]'
    #         # elements = wD.safe_click_elements('//span[contains(@class, "mectrl_screen_reader_text")]', cs.dom_wait, cs.tomer_wait, tomerOn, click_retry)
    #         # if 'False' in elements:
    #         #     messagebox.showinfo("alert","タイムアウトしました:" + locator)

    #     # 情報保存
    #         # result = wD.dom_output(dirname +"/eCAST_login_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         # if result == 'False':
    #         #     result = wD.dom_output(dirname +"/eCAST_login_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         #     if result == 'False':
    #         #         messagebox.showinfo("alert","タイムアウトしました")

    #         # サインアウト
    #         locator = '//button[contains(@id, "mectrl_body_signOut")]'
    #         elements = wD.safe_click_elements(locator, cs.dom_wait, cs.tomer_wait, tomerOn, click_retry)
    #         if 'False' in elements:
    #             txt="タイムアウトしました:" + locator
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             # messagebox.showinfo("alert","タイムアウトしました:" + locator)

    #         # result = wD.dom_output(dirname +"/eCAST_acc_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         # if result == 'False':
    #         #     result = wD.dom_output(dirname +"/eCAST_acc_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         #     if result == 'False':
    #         #         messagebox.showinfo("alert","タイムアウトしました")

    #         # 別のアカウントを選択
    #         locator = '//div[contains(@id, "otherTileText")]'
    #         click_retry_temp = 30
    #         elements = wD.safe_click_elements(locator, cs.dom_wait, cs.tomer_wait, tomerOn, click_retry_temp)
    #         if 'False' in elements:
    #             txt="タイムアウトしました:" + locator
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             # messagebox.showinfo("alert","タイムアウトしました:" + locator)

    #         #サインイン
    #         result = wD.dom_output(dirname +"/eCAST_signin_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         if result == 'False':
    #             result = wD.dom_output(dirname +"/eCAST_signin_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #             if result == 'False':
    #                 txt="タイムアウトしました:"
    #                 app.progress_message.config(text=txt + "終了してください")
    #                 threading.Event().wait()
    #                 # messagebox.showinfo("alert","タイムアウトしました")

    #         # 別のアカウントを選択
    #         # locator = '//input[contains(@name, "loginfmt")]'
    #         # elements = wD.safe_click_elements('//input[contains(@name, "loginfmt")]', cs.dom_wait, cs.tomer_wait, tomerOn, click_retry)
    #         # if 'False' in elements:
    #         #     messagebox.showinfo("alert","タイムアウトしました:" + locator)

    #         # ログインID
    #         locator = '//input[contains(@name, "loginfmt")]'
    #         input_val = login_info[0] + "@000.fujifilm.com"
    #         elements =wD.safe_click_send_elements(locator ,cs.dom_wait, cs.tomer_wait, tomerOn, input_val  ,send_keys_enter=False
    #         ,send_keys_return = True,retry_count=5)
    #         if 'False' in elements:
    #             wD.check_retryCnt.append(locator + str(elements))
    #             txt="タイムアウトしました:" + locator
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             # messagebox.showinfo("alert","タイムアウトしました:" + locator)

    #         # ユーザーネームエラーへが出るとログインIDが間違っている
    #         # result = wD.dom_output(dirname +"/eCAST_next_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         # if result == 'False':
    #         #     result = wD.dom_output(dirname +"/eCAST_next_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         #     if result == 'False':
    #         #         messagebox.showinfo("alert","タイムアウトしました")

    #         #ユーザーネームエラー
    #         locator = '//div[contains(@id, "usernameError")]'
    #         dom_wait_r = 5
    #         text = wD.dom_elements_getText(locator, 1, dom_wait_r, cs.tomer_wait, tomerOn)
    #         if 'このユーザー名は間違っている可能性があります' in text:
    #             txt="ユーザーIDが間違っています:" + locator
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             # messagebox.showinfo("alert","ユーザーIDが間違っています:" + locator)
    #             # sys.exit()

    #         # dom_wait_r = 5
    #         # timeout = 5
    #         # locator = '//div[contains(@id, "usernameError")]'
    #         # result = wD.waitingForResEnd(locator, 'このユーザー名は間違っている可能性があります', timeout, dom_wait_r, cs.tomer_wait, tomerOn)
    #         # if not('False' in result):
    #         #     messagebox.showinfo("alert","ユーザーIDが間違っています:" + locator)
    #         #     sys.exit()

    #         # elements = wD.dom_elements('//div[contains(@id, "usernameError")]',5,cs.tomer_wait,tomerOn)

    #         # locator = '//div[contains(@id, "usernameError")]'
    #         # elements = wD.safe_click_elements(locator, 5, cs.tomer_wait, tomerOn, click_retry)
    #         # if not('False' in elements):
    #         #     messagebox.showinfo("alert","ユーザーIDが間違っています:" + locator)

    #         #パスワード
    #         # result = wD.dom_output(dirname +"/eCAST_pass_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         # if result == 'False':
    #         #     result = wD.dom_output(dirname +"/eCAST_pass_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         #     if result == 'False':
    #         #         messagebox.showinfo("alert","タイムアウトしました")

    #         # パスワード
    #         locator = '//input[contains(@name, "passwd")]'
    #         input_val = login_info[1] #"00471Ktx00568"
    #         #この場合はreturnキーを打たない ベリファイするため
    #         elements =wD.safe_click_send_verify_elements(locator ,cs.dom_wait, cs.tomer_wait, tomerOn, input_val  ,send_keys_enter=False
    #         ,send_keys_return = False,retry_count=30)
    #         if 'False' in elements:
    #             wD.check_retryCnt.append(locator + str(elements))
    #             txt="タイムアウトしました:" + locator
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             # messagebox.showinfo("alert","タイムアウトしました:" + locator)
    #         elements[0][0].send_keys(Keys.RETURN)

    #         # ぺージ遷移まで待つ
    #         wD.wait_for_page_transition(30)

    #         #ダミー後方要素待ち
    #         locator = '//iframe[contains(@id, "PowerAppsSharedAppHostIframe")]'
    #         elements = wD.dom_elements(locator,60,cs.tomer_wait,tomerOn)
    #         if 'False' in elements:
    #             txt="タイムアウトしました:" + locator
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             # messagebox.showinfo("alert","タイムアウトしました" + locator)
    #         iframes = wD.driver.find_elements(By.TAG_NAME, 'iframe')
    #         if len(iframes)<4:
    #             txt="タイムアウトしました:iframe" + locator
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             # messagebox.showinfo("alert","タイムアウトしました:iframe")

    #         #
    #         # result = wD.dom_output(dirname +"/eCAST_default_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         # if result == 'False':
    #         #     result = wD.dom_output(dirname +"/eCAST_default_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         #     if result == 'False':
    #         #         messagebox.showinfo("alert","タイムアウトしました")

    #         # # それぞれのiframeのIDを取得
    #         # for index, iframe in enumerate(iframes):
    #         #     iframe_id = iframe.get_attribute("id")
    #         #     iframe_name = iframe.get_attribute("name")
    #         #     print(f"iframe {index + 1}: ID = {iframe_id}, Name = {iframe_name}")
    #         #

    #         # defaultのフレーム切替
    #         wD.driver.switch_to.frame("default")  # 最初のiframeに切り替えます
    #         # #frame0の内容
    #         # result = wD.dom_output(dirname +"/eCAST_default_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         # if result == 'False':
    #         #     result = wD.dom_output(dirname +"/eCAST_default_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         #     if result == 'False':
    #         #         messagebox.showinfo("alert","タイムアウトしました")

    #         # 0番のフレーム切替
    #         wD.driver.switch_to.frame(iframes[0])  # 最初のiframeに切り替えます
    #         #frame0の内容
    #         result = wD.dom_output(dirname +"/eCAST_frame0_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         if result == 'False':
    #             result = wD.dom_output(dirname +"/eCAST_frame0_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #             if result == 'False':
    #                 txt="タイムアウトしました:"
    #                 app.progress_message.config(text=txt + "終了してください")
    #                 threading.Event().wait()
    #                 # messagebox.showinfo("alert","タイムアウトしました")
    #         locator = '//div[contains(@id, "AppDetailsSec_1_Item_1")]'
    #         elements = wD.dom_elements(locator,300,cs.tomer_wait,tomerOn)
    #         if 'False' in elements:
    #             txt="タイムアウトしました" + locator
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             messagebox.showinfo("alert","タイムアウトしました")

    #         # wD.driver.switch_to.default_content()
    #         # wD.driver.switch_to.frame(iframes[1])  # 最初のiframeに切り替えます
    #         # #frame
    #         # result = wD.dom_output(dirname +"/eCAST_frame1_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         # if result == 'False':
    #         #     result = wD.dom_output(dirname +"/eCAST_frame1_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         #     if result == 'False':
    #         #         messagebox.showinfo("alert","タイムアウトしました")

    #         # wD.driver.switch_to.default_content()
    #         # wD.driver.switch_to.frame(iframes[2])  # 最初のiframeに切り替えます
    #         # #frame
    #         # result = wD.dom_output(dirname +"/eCAST_frame2_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         # if result == 'False':
    #         #     result = wD.dom_output(dirname +"/eCAST_frame2_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         #     if result == 'False':
    #         #         messagebox.showinfo("alert","タイムアウトしました")

    #         # wD.driver.switch_to.default_content()
    #         # wD.driver.switch_to.frame(iframes[3])  # 最初のiframeに切り替えます
    #         # #frame
    #         # result = wD.dom_output(dirname +"/eCAST_frame3_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         # if result == 'False':
    #         #     result = wD.dom_output(dirname +"/eCAST_frame3_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         #     if result == 'False':
    #         #         messagebox.showinfo("alert","タイムアウトしました")

    #         # wD.driver.switch_to.default_content()
    #         # wD.driver.switch_to.frame(iframes[4])  # 最初のiframeに切り替えます
    #         # #frame
    #         # result = wD.dom_output(dirname +"/eCAST_frame4_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         # if result == 'False':
    #         #     result = wD.dom_output(dirname +"/eCAST_frame4_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         #     if result == 'False':
    #         #         messagebox.showinfo("alert","タイムアウトしました")

    #         # wD.driver.switch_to.default_content()
    #         # wD.driver.switch_to.frame(iframes[5])  # 最初のiframeに切り替えます
    #         # #frame
    #         # result = wD.dom_output(dirname +"/eCAST_frame5_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         # if result == 'False':
    #         #     result = wD.dom_output(dirname +"/eCAST_frame5_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         #     if result == 'False':
    #         #         messagebox.showinfo("alert","タイムアウトしました")

    #         # eCAST　事務向けへ遷移
    #         locator = '//div[contains(@id, "AppDetailsSec_1_Item_1")]'
    #         click_retry_temp = 30
    #         elements = wD.safe_click_elements(locator, cs.dom_wait, cs.tomer_wait, tomerOn, click_retry_temp)
    #         if 'False' in elements:
    #             txt="タイムアウトしました" + locator
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             # messagebox.showinfo("alert","タイムアウトしました:" + locator)

    #         # wD.driver.switch_to.default_content()
    #         # result = wD.dom_output(dirname +"/eCAST_apri_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         # if result == 'False':
    #         #     result = wD.dom_output(dirname +"/eCAST_apri_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         #     if result == 'False':
    #         #         messagebox.showinfo("alert","タイムアウトしました")
    #         wD.driver.switch_to.default_content()
    #         locator = "//li[contains(@aria-label, '契約帳票')]"
    #         elements = wD.dom_elements(locator,10,cs.tomer_wait,tomerOn)
    #         if 'False' in elements:
    #             txt="タイムアウトしました" + locator
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             # messagebox.showinfo("alert","タイムアウトしました")

    # ウィンドウを最小化
    # wD.driver.minimize_window()
    # wD.driver.set_window_size(10, 10) #
    # # 契約帳票をクリック
    #     txt = "契約帳票に移動"
    #     logger.info(txt)
    #     app.progress_message.config(text=txt)

    #     locator = "//li[@aria-label='契約帳票']"
    #     click_retry_temp = 5
    #     elements = wD.safe_click_elements(locator, cs.dom_wait, cs.tomer_wait, tomerOn, click_retry_temp)
    #     if 'False' in elements:
    #         txt="タイムアウトしました" + locator
    #         logger.info(txt)
    #         app.progress_message.config(text=txt + "終了してください")
    #         threading.Event().wait()
    #         # messagebox.showinfo("alert","タイムアウトしました:" + locator)

    # # 契約帳票
    #     # result = wD.dom_output(dirname +"/eCAST_契約帳票_body_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #     # if result == 'False':
    #     #     result = wD.dom_output(dirname +"/eCAST_契約帳票_menu_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #     #     if result == 'False':
    #     #         messagebox.showinfo("alert","タイムアウトしました")
    #     testF = False #テスト用
    #     dest_path = "" #エクスポートした時に生成されたフォルダ
    #     if testF!=True:

    #     # aria-label="SC02B.依頼仕分け（社員対象のみ）"
    #     #SC02B.依頼仕分け（社員対象のみ）をクリックして選択リストだす
    #     #  <span class="ms-ContextualMenu-itemText label-453">SC02B.依頼仕分け（社員対象のみ）</span>
    #         # locator = "//button[@aria-label='SC02B.依頼仕分け（社員対象のみ）']"
    #         locator = "//button[contains(@id, 'ViewSelector') and .//i[@data-icon-name='ChevronDown']]"
    #         click_retry_temp = 5
    #         elements = wD.safe_click_elements("//button[contains(@id, 'ViewSelector') and .//i[@data-icon-name='ChevronDown']]", 10, cs.tomer_wait, tomerOn, click_retry_temp)
    #         if 'False' in elements:
    #             txt="タイムアウトしました" + locator
    #             logger.info(txt)
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             # messagebox.showinfo("alert","タイムアウトしました:" + locator)
    #     # <button
    #     #                               data-value="3b1fae8a-b00a-ec11-b6e6-000d3acf3a64"
    #     #                               role="menuitemradio"
    #     #                               title="SC05.作成完了"

    #     # # ＴＳＣ5年満了　配分済
    #     #     txt = "ＴＳＣ5年満了　配分済に移動"
    #     #     logger.info(txt)
    #     #     app.progress_message.config(text=txt)
    #     #     locator = "//button[@title ='ＴＳＣ5年満了　配分済']"
    #     #     # locator = "//button[@title ='SC03C.作成未着手（手作成）']"
    #     #     click_retry_temp = 5
    #     #     elements = wD.safe_click_elements(locator, 10, cs.tomer_wait, tomerOn, click_retry_temp)
    #     #     # elements = wD.safe_click_elements("//button[@title ='SC02B.依頼仕分け（社員対象のみ）']", 10, cs.tomer_wait, tomerOn, click_retry_temp)
    #     #     if 'False' in elements:
    #     #         txt="タイムアウトしました" + locator
    #     #         logger.info(txt)
    #     #         app.progress_message.config(text=txt + "終了してください")
    #     #         threading.Event().wait()
    #     #         # messagebox.showinfo("alert","タイムアウトしました:" + locator)

    #     # ＴＳＣ5年満了　配分済
    #         # locator_txt = "ＴＳＣ5年満了　配分済"
    #         locator_txt = "修正（旧条文）　配分済"
    #         txt = locator_txt + "に移動"
    #         logger.info(txt)
    #         app.progress_message.config(text=txt)
    #         locator = f"//button[@title ='{locator_txt}']"
    #         # locator = "//button[@title ='ＴＳＣ5年満了　配分済']"
    #         # locator = "//button[@title ='SC03C.作成未着手（手作成）']"
    #         click_retry_temp = 5
    #         elements = wD.safe_click_elements(locator, 10, cs.tomer_wait, tomerOn, click_retry_temp)
    #         # elements = wD.safe_click_elements("//button[@title ='SC02B.依頼仕分け（社員対象のみ）']", 10, cs.tomer_wait, tomerOn, click_retry_temp)
    #         if 'False' in elements:
    #             txt="タイムアウトしました" + locator
    #             logger.info(txt)
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             # messagebox.showinfo("alert","タイムアウトしました:" + locator)

    #   # # SC05契約帳票
    #         # result = wD.dom_output(dirname +"/eCAST_修正（旧条文）　配分済_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         # if result == 'False':
    #         #     result = wD.dom_output(dirname +"/eCAST_修正（旧条文）　配分済_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         #     if result == 'False':
    #         #         messagebox.showinfo("alert","タイムアウトしました")

    #         #LOAD待ち
    #         locator = "//div[@id ='LPCPreLoadFontIconContainer']"
    #         elements = wD.dom_elements(locator,15,cs.tomer_wait,tomerOn)
    #         if 'False' in elements:
    #             txt="タイムアウトしました" + locator
    #             logger.info(txt)
    #             app.progress_message.config(text=txt + "終了してください")
    #             threading.Event().wait()
    #             # messagebox.showinfo("alert","タイムアウトしました:" + locator)

    #   # # SC05契約帳票
    #         result = wD.dom_output(dirname +"/eCAST_契約帳票ＴＳＣ5年満了_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #         if result == 'False':
    #             result = wD.dom_output(dirname +"/eCAST_契約帳票ＴＳＣ5年満了_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #             if result == 'False':
    #                 txt="タイムアウトしました"
    #                 app.progress_message.config(text=txt + "終了してください")
    #                 threading.Event().wait()
    #                 # messagebox.showinfo("alert","タイムアウトしました")

    # # Excel に​​エクスポートをクリック
    #     txt = "Excel に​​エクスポート"
    #     logger.info(txt)
    #     app.progress_message.config(text=txt)
    #     locator = "//li[@aria-label='Excel に​​エクスポート']"
    #     click_retry_temp = 5
    #     elapsed_time = datetime.datetime.now()
    #     elements = wD.safe_click_elements(locator, 15, cs.tomer_wait, tomerOn, click_retry_temp)
    #     if 'False' in elements:
    #         txt="タイムアウトしました" + locator
    #         logger.info(txt)
    #         app.progress_message.config(text=txt + "終了してください")
    #         threading.Event().wait()
    #         # messagebox.showinfo("alert","タイムアウトしました:" + locator)
    #     # download_folder = os.path.expanduser("~\Downloads")
    #     dest_path = cu.copy_and_delete_files(download_folder, dest_folder, "", elapsed_time, 20)

    # SC05契約帳票
    # result = wD.dom_output(dirname +"/eCAST_契約帳票ＴＳＣ5年満了_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    # if result == 'False':
    #     result = wD.dom_output(dirname +"/eCAST_契約帳票ＴＳＣ5年満了_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
    #     if result == 'False':
    #         messagebox.showinfo("alert","タイムアウトしました")

    # 枚数取得

    # ウィンドウを最小化
    # wD.driver.minimize_window()
    # wD.driver.set_window_size(10, 10) #
    # 契約帳票に移動
    eu.keiyaku_move(logger, app, wD, cs, tomerOn)

    # データ取得無限　LOOP START///////////////////////////////////////////////////////////////////////////////
    wait_loopTime = 300  # 1200 #sec
    while True:

        prev_sets = set()  # 以前のセットを保存するためのセット aria-label="継続契約書"
        list_cyouhiyou = list()
        # .find_elements(By.XPATH, "//span[contains(@class, 'statusContainer')]")
        # elements = wD.dom_elements("//span[contains(@class, 'statusContainer')]", 15, cs.tomer_wait, tomerOn)[0].text
        txt = "各契約情報を取得中:総数繰上げ表示"
        logger.info(txt)
        app.progress_message.config(text=txt)
        count = 0
        ele0_flag = False  # 先に、データ有無を確認
        exists_uploadFile = []
        # データ取得LOOP START///////////////////////////////////////////////////////////////////////////////
        while True:
            while True:
                # 先に、データ有無を確認
                locator = "//span[contains(@class, 'statusContainer')]"  # aタグ hrefを取り出す
                elements = wD.dom_elements(locator, 5, cs.tomer_wait, True)
                if elements == "False":
                    txt = "タイムアウトしました:データなし" + locator
                    logger.info(txt)
                    app.progress_message.config(text=txt + "終了してください")
                    threading.Event().wait()
                ele0_flag = False
                if "0 - 0/0" in elements[0].text:
                    # データがありません]
                    ele0_flag = True
                    break
                hrefs = []
                fxscs = []
                keiyakuIds = []
                locator = "//a[contains(@aria-label, '継続契約書')]"  # aタグ hrefを取り出す
                elements = wD.dom_elements(locator, 5, cs.tomer_wait, True)
                if elements == "False":
                    txt = "タイムアウトしました:データなし" + locator
                    logger.info(txt)
                    app.progress_message.config(text=txt + "終了してください")
                    threading.Event().wait()
                # locator = "//a[contains(@aria-label, '継続契約書')/div/span[contains(@class, 'spanWrapper-')]"
                # locator = "//a[contains(@aria-label, '継続契約書')]/div/span[contains(@class, 'spanWrapper-') and contains(text(), 'fx')]"
                # カラム２で取得
                locator = "//div[contains(@aria-colindex, '2')]//a[contains(@class, 'ms-Link linkStyles')]"
                elements2 = wD.dom_elements(locator, 5, cs.tomer_wait, True)
                if elements2 == "False":
                    txt = "タイムアウトしました:データなし" + locator
                    logger.info(txt)
                    app.progress_message.config(text=txt + "終了してください")
                    threading.Event().wait()
                    # messagebox.showinfo("alert","タイムアウトしました" + locator)
                # リンクを取得
                hrefs = [ele.get_attribute("href") for ele in elements]
                # 契約書作成依頼を取得
                fxscs = [ele.text for ele in elements2]

                # # 要素は、同じの為、番号で附番　keiyakuIds = [ele.text for ele in elements]
                # keiyakuIds = [i for i, _ in enumerate(elements)]

                # if testF==True:
                #     locator = "//a/div[div[contains(text(), '2025')]]"
                # else:
                #     locator = "//a/div[div[contains(text(), 'fxsc20')]]"
                # elements = wD.dom_elements(locator, 5, cs.tomer_wait, True)
                # keiyakuIds = [ele.text for ele in elements]

                # データベースへの保存
                txt = str(fxscs)
                app.progress_message.config(text="保存予定データ" + txt)
                # ログ保存
                txt2 = "保存予定データ" + txt
                logger.info(txt2)

                new_sets = set()  # 新しいセット

                for href, keiyakuId in zip(hrefs, fxscs):
                    set_item = (href, keiyakuId)
                    if set_item not in prev_sets:
                        new_sets.add(set_item)
                        prev_sets.add(set_item)  # 新しいセットを追加

                        app.progress["maximum"] = len(prev_sets)
                        progress_text = f"{count + 1}/{len(prev_sets)}"
                        app.progress["value"] = str(count + 1)
                        # progress_text = f"{index+1}/{maxsets}"
                        app.progress_label.config(text=progress_text)
                        app.progress_label.update()
                        count = count + 1

                # 新しいセットが旧セットにすべて含まれている場合に break
                if len(new_sets) == 0:
                    break

                # スクロール
                wait_time = 2
                # button id="CreatePersonalView-button"
                # wD.dom_elements("//button[@id ='CreatePersonalView-button']",15,cs.tomer_wait,tomerOn)

                location = "CreatePersonalView-button"  # "//button[@id ='CreatePersonalView-button']"
                if (len(prev_sets) % 50) < 30:
                    res = eu.scroll_next_page(wait_time, 1, location)
                else:  # 全数が50個の場合、何故か空回りするので、ENDキーを使用する
                    res = eu.scroll_next_pageEND(wait_time, 1, location)
                if res != "None":
                    time.sleep(5)  # スクロール後に待機

            # データない場合
            if ele0_flag == True:
                break

            # 矢印キー遷移後の位置
            locator = "//span[contains(@class, 'statusContainer')]"
            elements = wD.dom_elements(locator, 15, cs.tomer_wait, tomerOn)
            pages = elements[0].text

            # 最後まで移動か？
            locate = eu.find_pattern(pages)
            if locate[1] == locate[2]:  # 最後まで移動された
                break

            # 矢印キー　LOOPさせるので判定後に押す
            locator = "//button[@aria-label ='次のページ']"
            click_retry_temp = 5
            elements = wD.safe_click_elements(
                "//button[@aria-label ='次のページ']",
                10,
                cs.tomer_wait,
                tomerOn,
                click_retry_temp,
            )
            if "False" in elements:
                txt = "タイムアウトしました" + locator
                logger.info(txt)
                app.progress_message.config(text=txt + "終了してください")
                threading.Event().wait()
                # messagebox.showinfo("alert","タイムアウトしました:" + locator)

            time.sleep(5)  # 矢印キー後に待機

        # データある場合
        if ele0_flag != True:

            # データ取得LOOP END　////////////////////////////////////////////////////////////////////////////

            maxsets = len(prev_sets)

            app.progress["maximum"] = maxsets
            app.progress["value"] = 0
            progress_text = f"{0}/{maxsets}"
            # progress_text = f"{index+1}/{maxsets}"
            app.progress_label.config(text=progress_text)
            app.progress_label.update()

            txt = "各契約データのダウンロード中"
            logger.info(txt)
            app.progress_message.config(text=txt)
            count = 0
            
            if len(prev_sets) > 0:  # データがあれば実行
                for index, (reff, irai_No) in enumerate(prev_sets):
                    error_txt = "中途失敗で、登録者名が変更されていません:" + irai_No
                    error_txt2 = "アップロード時に失敗でした。保存はされていません:" + irai_No
                    while True:
                        # app.progress['value'] = index + 1
                        progress_text = f"{count+1}/{maxsets}"
                        # app.progress_label.config(text=progress_text)
                        # app.progress_label.update()
                        # 遷移する
                        wD.driver.get(reff)
                        # 契約書作成依頼読み込み待ち
                        locator = "//div[@id ='LPCPreLoadFontIconContainer']"
                        elements = wD.dom_elements(locator, 15, cs.tomer_wait, tomerOn)
                        if "False" in elements:
                            txt = "タイムアウトしました:データなし" + locator
                            logger.info(txt)
                            app.progress_message.config(text=txt)
                            break  # データがないので抜ける
                            threading.Event().wait()
                            # messagebox.showinfo("alert","タイムアウトしました:" + locator)

                        app.progress["value"] = count + 1
                        # progress_text = f"{index+1}/{maxsets}"
                        app.progress_label.config(text=progress_text)
                        app.progress_label.update()
                        count = count + 1
                        # 契約書作成依頼
                        # result = wD.dom_output(dirname +"/eCAST_契約書作成依頼_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
                        # if result == 'False':
                        #     result = wD.dom_output(dirname +"/eCAST4_契約書作成依頼_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
                        #     if result == 'False':
                        #         messagebox.showinfo("alert","タイムアウトしました")

                        # <button id="OverflowButton_button0$button" aria-label="契約書作成依頼 のその他のコマンド" aria-describedby="openflyout_description_OverflowButton_button0" aria-expanded="false" aria-haspopup="true" role="menuitem" title="契約書作成依頼 のその他のコマンド" tabindex="0" data-id="OverflowButton" data-lp-id="Form:fbecast_tran_contract_request-OverflowButton" type="button" class="pa-bh pa-sf pa-fp pa-ju pa-ew pa-bg pa-p pa-qc pa-si pa-sk pa-uw flexbox" style="outline-color: rgb(102, 102, 102);" data-pa-landmark-active-element="true"><span aria-hidden="true" class="pa-bx pa-be pa-a pa-gq "><span class="pa-dj pa-di pa-ca pa-a pa-p pa-v pa-sm "><span class="symbolFont MoreVertical-symbol pa-pw pa-dq pa-ca pa-dd pa-bt "></span></span><span class="pa-fv pa-ca pa-sl pa-dl pa-e pa-dm pa-sn pa-dk "></span></span></button>

                        # #画面縮小率
                        # scaling_factor = cu.get_scaling_factor()
                        # #契約関連データダウンロードが隠れている場合がある
                        # if scaling_factor > 1.2:
                        #     locator = "//button[@id ='OverflowButton_button0$button']"
                        #     elements = wD.dom_elements(locator,1.5,cs.tomer_wait,tomerOn)
                        #     if not('False' in elements):
                        #         click_retry_temp = 5
                        #         elements = wD.safe_click_elements(locator, 1, cs.tomer_wait, tomerOn, click_retry_temp)
                        #         if 'False' in elements:
                        #             txt="タイムアウトしました" + locator
                        #             logger.info(txt)
                        #             app.progress_message.config(text=txt + "終了してください")
                        #             threading.Event().wait()

                        time.sleep(1)
                        # 書類出力先情報をクリック aria-label="書類出力先情報"
                        txt = "書類出力先情報に移動"
                        logger.info(txt)
                        app.progress_message.config(text=txt)
                        locator = "//li[@aria-label='書類出力先情報']"
                        click_retry_temp = 5
                        elements = wD.safe_click_elements(
                            "//li[@aria-label='書類出力先情報']",
                            15,
                            cs.tomer_wait,
                            tomerOn,
                            click_retry_temp,
                        )
                        if "False" in elements:
                            txt = "タイムアウトしました" + locator
                            logger.info(txt)
                            app.progress_message.config(text=txt + "終了してください")
                            threading.Event().wait()
                            # messagebox.showinfo("alert","タイムアウトしました:" + locator)
                            # sys.exit()
                        # result = wD.dom_output(dirname +"/eCAST_書類出力先情報_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
                        # if result == 'False':
                        #     result = wD.dom_output(dirname +"/eCAST_書類出力先情報3_" + tx_nowtime + ".html",'/html/body',30,cs.tomer_wait,tomerOn)
                        #     if result == 'False':
                        #         messagebox.showinfo("alert","タイムアウトしました")

                        # ロード終了確認
                        txt = "書類出力先情報ロード終了確認"
                        logger.info(txt)
                        app.progress_message.config(text=txt)
                        locator = "//div[contains(@class, 'ms-Fabric brandIcon')]"
                        elements = wD.dom_elements(locator, 5, cs.tomer_wait, tomerOn)
                        if "False" in elements:
                            txt = "タイムアウトしました" + locator
                            logger.info(txt)
                            app.progress_message.config(text=txt + "終了してください")
                            threading.Event().wait()
                        
                        time.sleep(2)

                        # アップロード確認
                        txt = "アップロード確認"
                        logger.info(txt)
                        app.progress_message.config(text=txt)

                        location_txt = app.filename
                        locator = f"//a[@title ='{location_txt}']"
                        elements = wD.dom_elements(locator, 1.5, cs.tomer_wait, tomerOn)
                        #Falseであれば未設定、設定要
                        if "False" in elements:
                            # ファイルのアップロード
                            txt = "ファイルのアップロード"
                            logger.info(txt)
                            app.progress_message.config(text=txt)

                            locator = "//input[@aria-label ='ページ追加']"
                            elements = wD.dom_elements(
                                locator, 2.0, cs.tomer_wait, tomerOn
                            )
                            if "False" in elements:
                                txt = txt + ":タイムアウトしました" + locator
                                logger.info(txt)
                                app.progress_message.config(text=txt + "終了してください")
                                mail_result(login_info, list_cyouhiyou, exists_uploadFile ,cu, error_txt2)
                                # threading.Event().wait()
                            else:
                                # fn = app.filename
                                # vi = app.viewname
                                # us = app.username
                                # Windowsのダウンロードフォルダのパスを取得
                                download_folder = os.path.join(
                                    os.path.expanduser("~"), "Downloads"
                                )
                                download_file = download_folder + "\\" + app.filename
                                elements[0].send_keys(download_file)
                                time.sleep(4)

                            # アップロード済確認
                            txt = "アップロード済確認"
                            location_txt = app.filename
                            locator = f"//a[@title ='{location_txt}']"
                            elements = wD.dom_elements(
                                locator, 1.5, cs.tomer_wait, tomerOn
                            )
                            #Falseであれば未設定、設定要
                            if "False" in elements:

                            # ファイルのアップロードリトライ
                                txt = "ファイルのアップロードリトライ"
                                logger.info(txt)
                                app.progress_message.config(text=txt)

                                locator = "//input[@aria-label ='ページ追加']"
                                elements = wD.dom_elements(
                                    locator, 2.0, cs.tomer_wait, tomerOn
                                )
                                if "False" in elements:
                                    txt =  txt + ":タイムアウトしました" + locator
                                    logger.info(txt)
                                    app.progress_message.config(text=txt + "終了してください")
                                    mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt2)
                                    threading.Event().wait()

                                else:
                                    download_folder = os.path.join(
                                        os.path.expanduser("~"), "Downloads"
                                    )
                                    download_file = download_folder + "\\" + app.filename
                                    elements[0].send_keys(download_file)
                                    time.sleep(2)

                                # アップロード済確認２
                                txt = "アップロード済確認２"
                                location_txt = app.filename
                                locator = f"//a[@title ='{location_txt}']"
                                elements = wD.dom_elements(
                                    locator, 2.0, cs.tomer_wait, tomerOn
                                )
                                if "False" in elements:
                                    txt = "タイムアウトしました" + locator
                                    logger.info(txt)
                                    app.progress_message.config(text=txt + "終了してください")
                                    mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt2)
                                    threading.Event().wait()
                        else:
                            txt = "すでにファイルがセットされています" + locator
                            logger.info(txt)
                            exists_uploadFile.append(irai_No)
                            break
                            app.progress_message.config(text=txt + "終了してください")
                            threading.Event().wait()

                        while True: #エラー時にリフレッシュすると先頭タブに戻るため
                            # 事務対応をクリック
                            txt = "事務対応移動"
                            logger.info(txt)
                            app.progress_message.config(text=txt)
                            locator = "//li[@aria-label='事務対応']"
                            click_retry_temp = 5
                            elements = wD.safe_click_elements(
                                locator, 15, cs.tomer_wait, tomerOn, click_retry_temp
                            )
                            if "False" in elements:
                                txt = "タイムアウトしました" + locator
                                logger.info(txt)
                                app.progress_message.config(text=txt + "終了してください")
                                threading.Event().wait()

                            # #2次チェック担当者入力
                            # txt = '2次チェック担当者入力確認'
                            # logger.info(txt)
                            # app.progress_message.config(text=txt)
                            # locator_txt = '2次チェック担当者, 検索'
                            # locator = f"//input[@aria-label ='{locator_txt}']"
                            # elements = wD.dom_elements(locator,5,cs.tomer_wait,tomerOn)
                            # if 'False' in elements:
                            #     txt="タイムアウトしました" + locator
                            #     logger.info(txt)
                            #     app.progress_message.config(text=txt + "終了してください")
                            #     threading.Event().wait()
                            # time.sleep(0.5)
                            # elements[0].click()
                            # input_val = app.username #'fxsc00553' #app.username
                            # # send_keysで1つまたは2つの文字を送る(エラーが発生する可能性がある)
                            # elements = wD.safe_send_keys(elements[0], input_val, 2, tomerOn=True,
                            #                     enter_key=True, return_key=False)
                            # #入力結果確認　aria-label="樋口 憲弘 の削除"
                            # txt = '2次チェック担当者の入力済み欄確認'
                            # logger.info(txt)
                            # app.progress_message.config(text=txt)
                            # locator_txt = 'true'
                            # locator = f"//div[contains(@data-pa-landmark-active-element,'{locator_txt}')]"
                            # locator2 = f"//input[contains(@data-pa-landmark-active-element,'{locator_txt}')]"
                            # elements = wD.judge_twoelements(locator, locator2, 1, cs.tomer_wait, tomerOn, retry_count=5)
                            # # elements = wD.dom_elements(locator,5,cs.tomer_wait,tomerOn)
                            # if 'False' in elements:
                            #     txt="タイムアウトしました" + locator
                            #     logger.info(txt)
                            #     app.progress_message.config(text=txt + "終了してください")
                            #     threading.Event().wait()
                            # #ここでtitleが空白でなくなり、担当者名が入る
                            # time.sleep(2)
                            # ti = elements[0].get_attribute('title')
                            # if ti != '':
                            #     txt = '成功:2次チェック担当者の入力済み'
                            #     logger.info(txt)
                            #     app.progress_message.config(text=txt)
                            # else:
                            #     txt = '失敗:2次チェック担当者の入力済み'
                            #     logger.info(txt)
                            #     app.progress_message.config(text=txt)
                            #     app.progress_message.config(text=txt + "終了してください")
                            #     threading.Event().wait()

                            # 未対応を対応済にする
                            txt = "未対応を対応済にする"
                            logger.info(txt)
                            app.progress_message.config(text=txt)
                            locator_txt = "事務対応結果"
                            locator = f"//select[contains(@aria-label,'{locator_txt}')]"
                            elements = wD.dom_elements(locator, 5, cs.tomer_wait, tomerOn)
                            if "False" in elements:
                                txt = "タイムアウトしました" + locator
                                logger.info(txt)
                                app.progress_message.config(text=txt + "終了してください")
                                threading.Event().wait()

                            select_object = wD.selectelements_from_dom(elements)
                            # <select> 要素をSelectオブジェクトに変換
                            # select_object = wD.driver.Select(elements)
                            # <option> 要素を選択する (テキストで選択)
                            select_object.select_by_visible_text("対応済")

                            # 動画見ますと、対応済みに変更した後、「保存」ボタンが推されてないみたいです。

                            # 保存すると、二次チェック者がログイン者に変わると思いますので
                            # その後、二次チェック者名を樋口さんに上書きし、さらに、保存ボタンを押すことで、対応済み、かつ、樋口さんＩＤで保存が可能かと思います。

                            # 保存する　aria-label="保存. 上書き保存 (Ctrl+S)"
                            txt = "保存. 上書き保存"
                            logger.info(txt)
                            app.progress_message.config(text=txt)                            
                            locator_txt = "保存. 上書き保存"
                            locator = f"//button[contains(@aria-label,'{locator_txt}')]"
                            click_retry_temp = 5
                            elements = wD.safe_click_elements(
                                locator, 5, cs.tomer_wait, tomerOn, click_retry_temp
                            )
                            if "False" in elements:
                                txt = "タイムアウトしました" + locator
                                logger.info(txt)
                                app.progress_message.config(text=txt + "終了してください")
                                threading.Event().wait()

                            time.sleep(1)
                            # チェック担当者を消すために事務対応をクリック こうしないとチェック担当者をクリックしても反応しない
                            locator = "//li[@aria-label='事務対応']"
                            click_retry_temp = 5


                            elements = wD.safe_click_elements(locator, 15, cs.tomer_wait, tomerOn, click_retry_temp)
                            if "False" in elements:
                                txt = "タイムアウトしました" + locator
                                logger.info(txt)
                                app.progress_message.config(text=txt + "終了してください")
                                mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt)
                                threading.Event().wait()
                                                # wD.dom_output(dirname +"/eCAST_All_body_" + tx_nowtime + "_4.html",'/html/body',30,cs.tomer_wait,tomerOn)
                            time.sleep(2)


                            # # ActionChainsを使用してホバリング
                            # actions = ActionChains(wD.driver).move_to_element(elements[0]).perform()
                            # actions.move_to_element(elements[0]).perform()
                            # wD.hover_dom(elements[0])
                            # Xをクリックしてチェック担当者を消す　2回押し 
                            # JavaScriptを使ってホバー効果をトリガーし、Xボタンを表示
                            # wD.driver.execute_script("var evt = document.createEvent('MouseEvents'); evt.initMouseEvent('mouseover', true, true, window, 1, 0, 0, 0, 0, false, false, false, false, 0, null); arguments[0].dispatchEvent(evt);", elements[0])

                            time.sleep(2)
                            # Xをクリックしてチェック担当者を消す　2回押し
                            txt = "Xをクリックしてチェック担当者を消す"
                            logger.info(txt)
                            app.progress_message.config(text=txt)
                            locator = "//ul[@title='2次チェック担当者']"
                            click_retry_temp = 5
                            elements = wD.safe_click_elements(
                                locator, 15, cs.tomer_wait, tomerOn, click_retry_temp
                            )
                            if "False" in elements:
                                txt = "タイムアウトしました" + locator
                                logger.info(txt)
                                app.progress_message.config(text=txt + "終了してください")
                                mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt)
                                threading.Event().wait()
                            time.sleep(0.8)
                            # wD.hover_dom(elements[0])
                            elements[0].click()
                            # wD.driver.execute_script("arguments[0].click();", elements[0])
                            time.sleep(1)

                            # 先に消去されている可能性が高いので消去を先にチェック
                            txt = "2次チェック担当者の入力消去確認"
                            # txt = "Xをクリック" 
                            logger.info(txt)
                            app.progress_message.config(text=txt)                                
                            locator_txt = "2次チェック担当者"  # 消去
                            locator = f"//*[contains(@aria-label,'{locator_txt}')]"
                            # //*[contains(@aria-label, "2次チェック担当者")]
                            elements = wD.dom_elements(locator, 3, cs.tomer_wait, tomerOn)
                            tname = ''
                            if ("False" in elements)==False: #成功
                                tname = elements[0].tag_name

                            if "False" in elements or tname != "input": #elmenstがFalseかtagがinputなら担当者がいる
                                # 失敗した場合はデータがある可能性が高い
                                txt = "Xをクリック" 
                                logger.info(txt)
                                app.progress_message.config(text=txt)
                                locator_txt = "の削除"  # 消去
                                locator = f"//button[contains(@aria-label,'{locator_txt}')]"
                                # app.progress_message.config(text=txt)
                                # locator = "//ul[@title='2次チェック担当者']"
                                elements = wD.dom_elements(locator, 5, cs.tomer_wait, tomerOn)
                                # elements = wD.safe_click_elements(
                                #     locator, 15, cs.tomer_wait, tomerOn, click_retry_temp
                                # )
                                # if "False" in elements:
                                #     txt = "タイムアウトしました" + locator
                                #     logger.info(txt)
                                #     app.progress_message.config(text=txt + "終了してください")
                                #     mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt)
                                #     threading.Event().wait()

                                if ("False" in elements) == False: #おそらく前のクリックで先に消えている
                                    time.sleep(0.8)
                                    # wD.hover_dom(elements[0])
                                    elements[0].click() #Xボタンを押す
                                    # wD.driver.execute_script("arguments[0].click();", elements[0])
                                    time.sleep(0.8)
                            # 消えている場合下部に移動、消えていなければ失敗になる

                            retrycnt = 0
                            retryLimit = 3
                            retry_exist = False
                            #"成功:2次チェック担当者の入力消去確認"で成功にはなるが次でタイムアウトするため、次の部分でretryする
                            #構成　inputでarea-label "2次チェック担当者ありー＞　空
                            #構成　buttonでarea-label "2次チェック担当者ありー＞　あり
                            time.sleep(3)
                            while True: 
                                txt = "2次チェック担当者の入力消去確認"
                                # txt = "Xをクリック" 
                                logger.info(txt)
                                app.progress_message.config(text=txt)                                
                                locator_txt = "2次チェック担当者"  # 消去
                                locator = f"//*[contains(@aria-label,'{locator_txt}')]"
                                # //*[contains(@aria-label, "2次チェック担当者")]
                                elements = wD.dom_elements(locator, 5, cs.tomer_wait, tomerOn)
                                if "False" in elements:
                                    txt = "タイムアウトしました" + locator
                                    logger.info(txt)
                                    app.progress_message.config(text=txt + "終了してください")
                                    mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt)
                                    threading.Event().wait()                                

                                # 見つかった要素の情報を取得（例: タグ名）
                                print(f"見つかった要素のタグ名: {elements[0].tag_name}")

                                # # 親要素を取得（1つ上のレベルの要素）
                                # locator = "./parent::*"
                                # # //*[contains(@aria-label, "2次チェック担当者")]
                                # parent_element = elements[0].find_elements(By.XPATH, './parent::*')
                                # if len(parent_element) == 0:
                                #     txt = "タイムアウトしました" + locator
                                #     logger.info(txt)
                                #     app.progress_message.config(text=txt + "終了してください")
                                #     mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt)
                                #     threading.Event().wait()                                   
                                # print(f"親要素のタグ名: {parent_element.tag_name}")
                                tname = elements[0].tag_name
                                # locator_txt = "true"  # 消去
                                # locator = f"//div[contains(@data-pa-landmark-active-element,'{locator_txt}')]"
                                # locator2 = f"//input[contains(@data-pa-landmark-active-element,'{locator_txt}')]"
                                # elements = wD.judge_twoelements(
                                #     locator, locator2, 6, cs.tomer_wait, tomerOn, retry_count=5
                                # )                                
                                # if "False" in elements:
                                #     txt = "タイムアウトしました" + locator
                                #     logger.info(txt)
                                #     app.progress_message.config(text=txt + "終了してください")
                                #     mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt)
                                #     threading.Event().wait()
                                # ここでtitleが空白なら、担当者名がない
                                # time.sleep(3)
                                # ti = elements[0].get_attribute("text")
                                # ti = elements[0].get_attribute("title")                                
                                if tname == "input":
                                    txt = "成功:2次チェック担当者の入力消去確認"
                                    logger.info(txt)
                                    app.progress_message.config(text=txt)
                                else:
                                    txt = "失敗:2次チェック担当者の入力消去確認"
                                    logger.info(txt)
                                    app.progress_message.config(text=txt)
                                    app.progress_message.config(text=txt + "終了してください")
                                    mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt)
                                    threading.Event().wait()

                                txt = "2次チェック担当者いったん消しました。:"
                                app.progress_message.config(text=txt)

                                time.sleep(3)
                                # 2次チェック担当者入力
                                txt = "2次チェック担当者入力確認"
                                locator_txt = "2次チェック担当者, 検索"
                                locator = f"//input[@aria-label ='{locator_txt}']"
                                elements = wD.dom_elements(locator, 5, cs.tomer_wait, tomerOn)
                                if "False" in elements:
                                    time.sleep(10)
                                    elements = wD.dom_elements(locator, 5, cs.tomer_wait, tomerOn)
                                    retry_exist = False
                                    if "False" in elements:
                                        retrycnt = retrycnt + 1
                                        txt = "タイムアウトしました" + locator + " retry:" + str(retrycnt)
                                        logger.info(txt)
                                        if retrycnt > retryLimit:
                                            app.progress_message.config(text=txt + "終了してください")
                                            mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt)
                                            threading.Event().wait()
                                        #LOOP継続 
                                        retry_exist = True
                                        wD.driver.refresh() #先頭タブに戻る
                                        time.sleep(8)
                                        break
                                    else:
                                        break
                                else:
                                    time.sleep(0.5)
                                    break
                            if retry_exist == False: #Trueの場合は、breakせずに上のWhile文からやりなおす
                                break #正常処理の時、break

                        txt = "2次チェック担当者いったん消しました。２:"
                        app.progress_message.config(text=txt)

                        time.sleep(3)
                        # 2次チェック担当者入力
                        txt = "2次チェック担当者入力確認２"
                        locator_txt = "2次チェック担当者, 検索"
                        locator = f"//input[@aria-label ='{locator_txt}']"
                        elements = wD.dom_elements(locator, 5, cs.tomer_wait, tomerOn)
                        if "False" in elements:
                            time.sleep(10)
                            elements = wD.dom_elements(locator, 5, cs.tomer_wait, tomerOn)
                            if "False" in elements:
                                txt = "タイムアウトしました" + locator
                                logger.info(txt)
                                app.progress_message.config(text=txt + "終了してください")
                                mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt)
                                threading.Event().wait()
                                
                        time.sleep(0.5)

                        elements[0].click()
                        input_val = app.username  #'fxsc00553' #app.username
                        # send_keysで1つまたは2つの文字を送る(エラーが発生する可能性がある)
                        elements = wD.safe_send_keys(
                            elements[0],
                            input_val,
                            2,
                            tomerOn=True,
                            enter_key=True,
                            return_key=False,
                        )
                        
                        time.sleep(1)
                        # 入力結果確認　aria-label="樋口 憲弘 の削除"
                        txt = "2次チェック担当者の入力済み欄確認"
                        locator_txt = "true"
                        locator = f"//div[contains(@data-pa-landmark-active-element,'{locator_txt}')]"
                        elements = wD.dom_elements(locator, 5, cs.tomer_wait, tomerOn)
                        if "False" in elements:
                            txt = "タイムアウトしました" + locator
                            logger.info(txt)
                            app.progress_message.config(text=txt + "終了してください")
                            mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt)
                            threading.Event().wait()

                        # 保存する　aria-label="保存. 上書き保存 (Ctrl+S)"
                        locator_txt = "保存. 上書き保存"
                        locator = f"//button[contains(@aria-label,'{locator_txt}')]"
                        click_retry_temp = 5
                        elements = wD.safe_click_elements(
                            locator, 5, cs.tomer_wait, tomerOn, click_retry_temp
                        )
                        if "False" in elements:
                            txt = "タイムアウトしました" + locator
                            logger.info(txt)
                            app.progress_message.config(text=txt + "終了してください")
                            mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt)
                            threading.Event().wait()
                        time.sleep(1)

                        # データベースへの保存
                        txt = str(irai_No)
                        app.progress_message.config(text="保存済みデータ：" + txt)
                        # ログ保存
                        txt2 = "保存済みデータ：" + txt
                        logger.info(txt2)
                        list_cyouhiyou.append(txt)

                        # #保存せず、キャンセル
                        # locator_txt = '変更の破棄'
                        # locator = f"//button[contains(@aria-label,'{locator_txt}')]"
                        # click_retry_temp = 5
                        # elements = wD.safe_click_elements(locator, 5, cs.tomer_wait, tomerOn, click_retry_temp)
                        # if 'False' in elements:
                        #     txt="タイムアウトしました" + locator
                        #     logger.info(txt)
                        #     app.progress_message.config(text=txt + "終了してください")
                        #     threading.Event().wait()

                        break
        # 契約帳票に移動

        error_txt = ""
        mail_result(login_info, list_cyouhiyou, exists_uploadFile, cu, error_txt)

        # if len(login_info) > 3 and len(list_cyouhiyou) > 0:
        #     if login_info[3].strip() != "":
        #         all_cyouhilyou = str(list_cyouhiyou)
        #         you_list = []
        #         items = login_info[3].strip().split(":")
        #         you_list.extend(items)
        #         you_list.append("katsumi.kobayashi.hg@fujifilm.com")
        #         me = "katsumi.kobayashi.hg@fujifilm.com"  # 送信元メールアドレス satoshi.nakano.mp@fujifilm.com
        #         title = "eCASTインポート結果"
        #         txt = all_cyouhilyou
        #         file_list = []
        #         cu.send_mail(me, you_list, title, txt, file_list)
        # wait_loopTime = 5
        time.sleep(wait_loopTime)
        wD.driver.refresh()
        time.sleep(10)
        eu.keiyaku_move(logger, app, wD, cs, tomerOn)

    if str(maxsets) == str(count):
        txt = (
            "正常終了しました。fxscのみ抽出しています。\n\n"
            + "処理予定数："
            + str(maxsets)
            + "\n処理数："
            + str(count)
            + "\n画面全数："
            + locate[2]
        )
    else:
        txt = (
            "終了:処理数が不足しています。fxscのみ抽出しています。\n\n"
            + "処理予定数："
            + str(maxsets)
            + "\n処理数："
            + str(count)
            + "\n画面全数："
            + locate[2]
        )

    db.insert_data(
        db.app_name, db.user_id, db.user_name, db.ip, maxsets, db.data, db.texts
    )

    logger.info(txt)
    app.progress_message.config(text=txt)
    messagebox.showinfo("alert", txt)
    sys.exit()
