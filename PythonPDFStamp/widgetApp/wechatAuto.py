import os
import time
import datetime

import pyautogui
import pyperclip

from PyQt5.QtCore import QFileSystemWatcher


class WeChatAuto():

    def __init__(self):
        self.image_path = "K:\\GithubCode\\juntevision\\PythonPDFStamp\\image\\findFileImg.png"
        self.watch_path = "C:\\Users\\郑勋\\Documents\\WeChat Files\\q37610672\\FileStorage\\File\\2022-03\\"
        self.__file_base_path ="C:\\Users\\郑勋\\Documents\\WeChat Files\\q37610672\\FileStorage\\File\\"
    def findUser(self, user):
        # 打开微信
        pyautogui.hotkey('ctrl', 'alt', 'w')
        time.sleep(1)
        # 搜索群聊名称--君泰-海康商务群-下单
        pyautogui.hotkey("ctrl", "f")
        pyperclip.copy(user)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.hotkey('Enter')
        time.sleep(1)

    #找到文件夹图像并点击
    def findImg(self, img_path):
        pyautogui.move(200,200)
        #image_path = "K:\\GithubCode\\juntevision\\PythonPDFStamp\\image\\findFileImg.png"
        image_path = img_path
        image_loc = pyautogui.locateOnScreen(image_path, grayscale=True)
        print(image_loc)
        center_loc = pyautogui.center(image_loc)
        print(center_loc)
        pyautogui.click(center_loc)

    def sendFile(self, file_path, image_path):
        pyperclip.copy(file_path)
        self.findImg(image_path)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.hotkey('Enter')
        time.sleep(1)
        pyautogui.hotkey('Enter')

    def get_wechat_file_path(self):
        folder_path_date=datetime.datetime.now().strftime('%Y-%m')
        folder_path = self.__file_base_path + folder_path_date
        #print(folder_path)
        return folder_path
if __name__ == "__main__":
    # findUser("文件传输助手")
    # file_path ='K:\\GithubCode\\juntevision\\PythonPDFStamp\\image\\findFileImg.png'
    # sendFile(file_path,file_path)
    wechatobj = WeChatAuto()
    #wechatobj.findImg("K:\\GithubCode\\juntevision\\PythonPDFStamp\\image\\findFileImg.png")
    #判断监控目录是否存在
    print(os.path.isdir(wechatobj.watch_path))
    print(os.path.exists(wechatobj.watch_path))
    print(wechatobj.get_wechat_file_path())
    print(os.path.isdir(wechatobj.get_wechat_file_path()))
# wechat_file_path ='C:\\Users\\郑勋\\Documents\\WeChat Files\\q37610672\\FileStorage\\File\\2022-03\\'
# file_name = os.listdir(wechat_file_path)
# print(file_name)

# fileWatcher = QFileSystemWatcher()
# fileWatcher.addPath(wechat_file_path)