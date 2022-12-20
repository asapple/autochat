import win32api, win32gui, win32con
import win32clipboard as clip
import time
import pyperclip
import pandas as pd
from PIL import Image
from io import BytesIO
import random
import hashlib
import wmi
import datetime
import os
import sys

def send_m(win):
    # 以下为“CTRL+V”组合键,回车发送
    win32api.keybd_event(17, 0, 0, 0)  # 有效，按下CTRL
    time.sleep(0.5)  # 需要延时
    win32gui.SendMessage(win, win32con.WM_KEYDOWN, 86, 0)  # V
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)  # 放开CTRL
    time.sleep(1)  # 缓冲时间
    win32gui.SendMessage(win, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)  # 回车发送
    return

#粘贴文本进剪切板
def txt_ctrl_v(txt_str):
    pyperclip.copy(txt_str)
    return

#粘贴图片进剪切板
def setImage(pic_name):
    imagepath = 'images\\'+pic_name
    img = Image.open(imagepath)
    output = BytesIO()
    img.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()
    clip.OpenClipboard() #打开剪贴板
    clip.EmptyClipboard()  #先清空剪贴板
    clip.SetClipboardData(win32con.CF_DIB, data)  #将图片放入剪贴板
    clip.CloseClipboard()

def get_window(className, titleName):
    win = win32gui.FindWindow(className, titleName)

    if win != 0:
        win32gui.SetForegroundWindow(win)  # 获取控制
        win32gui.ShowWindow(win, win32con.SW_MAXIMIZE)
        time.sleep(0.5)
    else:
        print('请注意：找不到名为【%s】的窗口' % titleName)
        os.system('pause')
        exit(0)
    
    return win


def sendTaskLog():
    filename = input("input:请输入范本文件名：")
    data = pd.read_excel(filename)

    argc = len(sys.argv)
    beginLen = 0
    if argc > 1:
        beginLen = int(sys.argv[1]) - 2
    for i in data.index.values:
        if i < beginLen:
            continue
        #消息之间的间隔时间
        time.sleep(0.5+random.random()*8)
        #获取发送窗口和发送信息
        actor = data.iloc[i,0]
        win = get_window('ChatWnd', '%s'%actor)
        str = data.iloc[i,1]

        #将信息粘贴入剪切板
        if str.startswith("pic:"):    # 发送图片
            str = str[4:]
            setImage(str)
            send_m(win)
        elif str.startswith("lin:"):    # 发送链接
            str_list = str.split(':')
            group_name = str_list[1]
            win = get_window('CefWebViewWnd', '微信')
            send_link(group_name)
        else:    # 发送文字
            txt_ctrl_v(str)
            send_m(win)
        print("info:【%s】发送【%s】成功"%(actor,str))

def yanzheng():
    seral = ""
    c = wmi.WMI()
    
    # 获取硬盘序列号
    for physical_disk in c.Win32_DiskDrive():
        hard_seral=physical_disk.SerialNumber   
        seral += hard_seral
    # 获取CPU序列号
    for cpu in c.Win32_Processor():
        cpu_seral=cpu.ProcessorId.strip()
        seral += cpu_seral
    # 获取主板序列号
    for board_id in c.Win32_BaseBoard():
        board_id=board_id.SerialNumber
        seral += board_id
    print("info:注册码使用的硬件信息为：",seral,sep="")

    # 获取当前时间
    now = datetime.datetime.now()
    # 获取年份和月份
    year = now.year
    month = now.month

    seral= seral + "Chenxuan" + str(year) + str(month)
    key = hashlib.sha256(seral.encode('utf-8')).hexdigest()
    
    keyConf = ''
    with open('config.txt', "r") as f:
        text = f.read()
        keyConf = text
    if keyConf == key:
        print("info:验证成功！")
        return
    else:
        print("info:配置文件config.txt注册码无效！")

    inKey = input("input:请输入注册码：")
    if key!=inKey:
        print("info:验证失败！")
        os.system('pause')
        exit(0)
    else:
        print("info:验证成功！")

def send_link(group_name):
    
    screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    
    # 将鼠标移动到分享的按键位置,并点击鼠标
    win32api.SetCursorPos((int(screen_width*0.9505),int(screen_height*0.0417)))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.3)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

    time.sleep(1)

    # 复制群名并粘贴进搜索框
    pyperclip.copy(group_name)
    win32api.keybd_event(17, 0, 0, 0)
    win32api.keybd_event(86, 0, 0, 0)
    time.sleep(0.3)
    win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)

    time.sleep(1)

    #按下并释放回车键，选定群聊
    win32api.keybd_event(0x0D, 0, 0, 0)
    win32api.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)

    time.sleep(1)

    # 将鼠标移动到发送的按键位置，并点击发送
    win32api.SetCursorPos((int(screen_width*0.5417),int(screen_height*0.6667)))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.3)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

yanzheng()

os.system('pause')
print("info:程序开始执行；")
sendTaskLog()
print("info:程序执行结束；")
os.system('pause')
# scheduler = BlockingScheduler()
# scheduler.add_job(sendTaskLog, 'interval', seconds=3)
# scheduler.add_job(sendTaskLog, 'cron',day_of_week='mon-fri', hour=7,minute=31,second='10',misfire_grace_time=30)
# scheduler.add_job(sendTaskLog, 'cron', day_of_week='mon-fri', hour=6, minute=55, second='10', misfire_grace_time=30)
# try:
# scheduler.start()
# except (KeyboardInterrupt, SystemExit):
    # pass