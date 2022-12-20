import pyperclip
import win32gui
import win32con
import win32api
import time
import os
import ctypes
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
def send_link(group_name):
    
    screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    
    # 将鼠标移动到分享的按键位置
    win32api.SetCursorPos((int(screen_width*0.9505),int(screen_height*0.0417)))
    # 鼠标点击
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

win = get_window('CefWebViewWnd', '微信')
send_link('test群聊')