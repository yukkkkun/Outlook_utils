
import os # ディレクトリの操作用
import chromedriver_binary
from selenium.webdriver import Chrome, ChromeOptions # ヘッドレスブラウザ
from selenium.webdriver.chrome.options import Options
import time # 処理待機用
import pyautogui # Chrome拡張のアクションボタンクリック用
import signal

# from selenium import webdriver

##https://www.dev-dev.net/entry/2018/09/03/232436


# username = os.environ['USERNAME']
# path_extentions = 'C:/Users/{}/OneDrive/shared-Yutaka/Outlook_tools/3.2.0_0.crx'.format(username)
# path_chromedriver = 'C:/Users/{}/OneDrive/shared-Yutaka/Outlook_tools/chromedriver_win32/chromedriver.exe'.format(username)


def open_with_extensions():

    path_extentions = '3.2.0_0.crx'
    path_chromedriver = 'chromedriver_win32/chromedriver.exe'
    options = ChromeOptions()
    options.add_extension(path_extentions)

    try:
        driver = Chrome(executable_path=path_chromedriver, options=options)

        # driver = Chrome(executable_path=path_chromedriver)

        # driver = Chrome(options=options)
        driver.set_window_position(0,0) # ブラウザの位置を左上に固定
        driver.set_window_size(600,740) # ブラウザのウィンドウサイズを固定

        url = "http://tokidoki-web.com"
        
        driver.get(url)

        print("end")
        
        # # 拡張機能のアクションボタンをクリック
        # print("拡張機能ON")
        # # ディスプレイ位置から左550px、上80pxの位置で１回クリックして１秒待機
        # pyautogui.click(550, 80, 1, 1, 'left')
        
        # # ポップアップ内の要素をクリック
        # time.sleep(2)
        # print("ポップアップ内クリック")
        # # ディスプレイ位置から左450px、上430pxの位置で１回クリックして１秒待機
        # pyautogui.click(450, 430, 1, 1, 'left')
        
        time.sleep(1000)

        # ～いろいろな処理～
    finally:
        os.kill(driver.service.process.pid,signal.SIGTERM)
    # """
    # 以降のクライアント先ページへの処理
    # """