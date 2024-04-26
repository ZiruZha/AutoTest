# *************************************
# ******此文件为myGui_V102.py库文件******
# *************************************
# Date:June,13th,2023
# Introduction:
# update on July,26th,2023
# 增加自动截屏记录功能
# Update on Dec,6th,2024
# 修改说明
# 一次测完跳闸、自检、遥信传动
# Update on Apr,26th,2024
# Version-200_lib
# 修改说明
# 1. 删除主程序，改为def main()，以供GUI调用
# 2. 增加逻辑————若图片识别5次未找到，报错跳出main()，避免GUI按钮进入死循环
# 3. 所用的图片、模板全部打包放入dataFiles文件中
# 4. 全部改用相对路径


import pyautogui
import time
import xlrd
import pyperclip
import os

# 鼠标左右键设置
left = "right"
right = "left"
coefficient = 0.5
# 每X条记录一次
RecordPerTimes = 10


# 定义鼠标事件

# pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159


class noPicFound(ValueError):
    pass


def mouseClick(clickTimes, lOrR, img, reTry):
    if reTry == 1:
        i = 0
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                break
            print("未找到匹配图片,5秒后重试")
            print(img)
            i = i + 1
            time.sleep(5)
            if i > 5:
                raise noPicFound('noPicFound')
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                print("重复")
                i += 1
            time.sleep(0.1)


def sel_word():
    img = 'dataFiles\\' + 'ifWord.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pyautogui.click(location.x, location.y, clicks=1, interval=0.2, duration=0.2, button=left)
    else:
        img = 'dataFiles\\' + '3.png'
        mouseClick(1, left, img, 1)
    time.sleep(0.1)


def Record():
    pyautogui.hotkey('ctrl', 'prntscrn')
    time.sleep(1)
    sel_word()
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.hotkey('alt', 'esc')
    time.sleep(1)
    print('record')
    pass


def logIn():
    pyautogui.press('left')
    time.sleep(1 * coefficient)
    pyautogui.press('=')
    time.sleep(1 * coefficient)
    pyautogui.press('=')
    time.sleep(1 * coefficient)
    pyautogui.press('down')
    time.sleep(1 * coefficient)
    pyautogui.press('left')
    time.sleep(1 * coefficient)
    pyautogui.press('=')
    time.sleep(1 * coefficient)
    pyautogui.press('=')
    time.sleep(1 * coefficient)
    pyautogui.press('down')
    time.sleep(1 * coefficient)
    pyautogui.press('enter')
    time.sleep(1 * coefficient)
    pass


def main(input):
    global RecordPerTimes
    RecordPerTimes = input
    print('Welcome to AutoTest-61850')
    # key = input('input times\n')
    # key = int(key)
    # coefficient = input('input time coefficient\n')
    # coefficient = float(coefficient)
    i = 0
    s = ''
    while i < 6:
        print('Countdown to begin---->' + str(5 - i))
        time.sleep(1)
        i = i + 1
        pass
    # if key == -1:
    #     i = -2
    # else:
    #     i = 0
    logIn()
    times = 0
    flg = 1  # flg = 1：跳闸传动；2：自检传动；3：遥信传动
    while flg < 4:
        pyautogui.press('enter')
        time.sleep(1 * coefficient)
        pyautogui.press('enter')
        time.sleep(2 * coefficient)
        times = times + 1
        if times % RecordPerTimes == 0:
            Record()
        pyautogui.press('down')
        time.sleep(1 * coefficient)
        pyautogui.press('enter')
        time.sleep(2 * coefficient)
        times = times + 1
        if times % RecordPerTimes == 0:
            Record()
        pyautogui.press('delete')
        time.sleep(1 * coefficient)
        pyautogui.press('down')
        time.sleep(1 * coefficient)
        chuanDongPngDic = {1: 'TJChuanDongEnd.png', 2: 'ZJChuanDongEnd.png', 3: 'YXChuanDongEnd.png'}
        img = 'dataFiles\\' + chuanDongPngDic.get(flg)
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            chuanDongDic = {1: '跳闸', 2: '自检', 3: '遥信'}
            s = s + chuanDongDic.get(flg) + '传动信号---->' + str(times / 2) + '个\n'
            times = 0
            Record()
            flg = flg + 1
            pyautogui.press('delete')
            time.sleep(1 * coefficient)
            pyautogui.press('down')
            time.sleep(1 * coefficient)
            pyautogui.press('enter')
            time.sleep(1 * coefficient)
            logIn()
            continue
        # if key == -1:
        #     continue
        # i = i + 1
        # print('Tested items---->' + str(i))
    # if times % RecordPerTimes != 0:
    #     Record()
    # print(s)
    return s
