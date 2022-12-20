import keyboard
import time
import threading

running = 1
def pause():
    global running 
    running = 0
    print("暂停键已按下，即将暂停（按c继续）...")
def conti():
    global running 
    running = 1
    print("继续键已按下，即将继续...")
keyboard.add_hotkey('p', pause)
keyboard.add_hotkey('c', conti)
# 在新线程中执行程序的主体逻辑
def main_thread():
    while True:
        if running:
            print(1)
            time.sleep(1)
            pass
        else:
            # 等待快捷键被触发
            time.sleep(0.1)

thread = threading.Thread(target=main_thread)
thread.start()