# *************************************
# ******此文件为myGui_V102.py库文件******
# *************************************
# Date:June,12th,2023
# Introduction:
# 1.通过识别图片选取鼠标操作对象，因此需设定屏幕分辨率设定（设置->屏幕->缩放与布局：100%；显示器分辨率：1920*1080）
# 2.只适配博电PowerTest ANA_GOOSE软件（常规采样软件，即只能测试G、DG装置。若要测试DA装置，需配合合并单元使用）
# 3.使用前需做好准备工作，包括但不限于：
# 3.1将Arptool、Word、PowerTest ANA_GOOSE软件打开
# 3.2设置装置通道识别码、软压板（所有保护定值不需要修改）
# 3.3投待测试保护功能硬压板
# 3.4...
# 4.目前无法测试PTCT断线、远跳、三相不一致保护、远方跳闸保护
# 5.目前测试记录通过截图保存，导致测试文件过大
# update on June,26th,2023
# version--V1.01
# 1.将装置类型、scd文件目录改为输入参数
# 2.DA、DG装置只导入一次iec文件
# 3.添加word标题
# update on July,14th,2023
# version--V1.02
# 1.增加按键等待时间
# 2.将部分鼠标操作替换为键盘操作
# update on August,22nd,2023
# version--V1.03
# 1.将部分鼠标操作替换为键盘操作
# version--V2.00
# V1版本使用博电测试仪V2版本使用昂立测试仪
# 图片识别、操作步骤改变
# Update on Apr,26th,2024
# Version-200_lib
# 修改说明
# 1. 删除主程序，改为def main()，以供GUI调用
# 2. 增加逻辑————若图片识别5次未找到，报错跳出main()，避免GUI按钮进入死循环
# 3. 所用的图片、模板全部打包放入dataFiles文件中
# 4. 全部改用相对路径

import pyautogui
import time
# import xlrd
import pyperclip
import os
import pandas as pd
import re
import sys

# 全局变量
# 测试等待时间
TestWaitingTime = 30
# 按键反应时间
ButtonReactTime = 1
# SCD文件名称
SCDFileName = '1.scd'
# 鼠标左右键设置
left = "right"
right = "left"
firstTime = 1
ifG = '0'
excelpath = "dataFiles\\TestItems.xlsx"
sheetDif = pd.read_excel(excelpath, sheet_name="dif")
sheetZone = pd.read_excel(excelpath, sheet_name="zone")
sheetRocRelay = pd.read_excel(excelpath, sheet_name="rocRelay")
sheetRocInv = pd.read_excel(excelpath, sheet_name="rocInv")
sheetOverload = pd.read_excel(excelpath, sheet_name="overload")
sheetOvervoltage = pd.read_excel(excelpath, sheet_name="overvoltage")
sheetAR = pd.read_excel(excelpath, sheet_name="AR")
sheetFload = pd.read_excel(excelpath, sheet_name="fload")
SCDAddr = ''


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


# 任务
def mainFunc():
    # 纵差
    # 投纵差，退距离，退零序过流
    para_RuanYaBan = 1100
    GongNengRuanYaBan(para_RuanYaBan)
    Dif()
    # 零序过流
    # 退纵差，退距离，投零序过流
    para_RuanYaBan = 1
    GongNengRuanYaBan(para_RuanYaBan)
    OC()
    # 距离
    # 退纵差，投距离，退零序过流
    para_RuanYaBan = 10
    GongNengRuanYaBan(para_RuanYaBan)
    Z()


def sel_arp():
    img = 'dataFiles\\' + 'ifArp.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pyautogui.click(location.x, location.y, clicks=1, interval=0.2, duration=0.2, button=left)
    else:
        img = 'dataFiles\\' + '1.png'
        mouseClick(1, left, img, 1)
    time.sleep(0.1)


def sel_powertest():
    img = 'dataFiles\\' + 'ifPowerTest.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pyautogui.click(location.x, location.y, clicks=1, interval=0.2, duration=0.2, button=left)
    else:
        img = 'dataFiles\\' + '2.png'
        mouseClick(1, left, img, 1)
    time.sleep(0.1)


def sel_onlly():
    img = 'dataFiles\\' + 'ifOnlly.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pyautogui.click(location.x, location.y, clicks=1, interval=0.2, duration=0.2, button=left)
    else:
        img = 'dataFiles\\' + '4.png'
        mouseClick(1, left, img, 1)
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


def GongNengRuanYaBan(para_RuanYaBan):
    sel_arp()
    img = 'dataFiles\\' + 'gongnengruanyabanSelected.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pass
    else:
        img = 'dataFiles\\' + 'gongnengruanyaban.png'
        mouseClick(1, left, img, 1)
    ch1 = int(para_RuanYaBan / 1000)
    ch2 = int(para_RuanYaBan % 1000 / 100)
    z = int(para_RuanYaBan % 100 / 10)
    oc = int(para_RuanYaBan % 10 / 1)
    img = 'dataFiles\\' + 'ch1.png'
    mouseClick(1, left, img, 1)
    down(ch1)
    img = 'dataFiles\\' + 'ch1.png'
    mouseClick(1, left, img, 1)
    i = 0
    while i < 6:
        pyautogui.press('tab')
        time.sleep(ButtonReactTime / 10)
        i += 1
    down(ch2)
    img = 'dataFiles\\' + 'ch1.png'
    mouseClick(1, left, img, 1)
    i = 0
    while i < 12:
        pyautogui.press('tab')
        time.sleep(ButtonReactTime / 10)
        i += 1
    down(z)
    img = 'dataFiles\\' + 'ch1.png'
    mouseClick(1, left, img, 1)
    i = 0
    while i < 18:
        pyautogui.press('tab')
        time.sleep(ButtonReactTime / 10)
        i += 1
    down(oc)


def SettingModify(type):
    sel_arp()
    img = 'dataFiles\\' + 'settingSelected.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pass
    else:
        img = 'dataFiles\\' + 'setting.png'
        mouseClick(1, left, img, 1)
    # type = 1---->纵差
    # type = 2---->零序II段
    # type = 3---->零序III段
    # type = 4---->距离I段
    # type = 5---->距离III段
    # type = 6---->距离III段
    # type = 7---->零序反时限
    # type = 8---->三相不一致
    # type = 9---->过流过负荷
    # type = 10---->电铁、钢厂等冲击性负荷
    # type = 11---->过电压及远方跳闸保护
    # type = 12---->3/2断路器接线————————————不测重合闸
    if type == 1:
        return
    elif type == 2:
        return
    elif type == 3:
        return
    elif type == 4:
        # 距离I段使能
        img = 'dataFiles\\' + 'z1En.png'
        searchSetting(img)
        down(1)
        # 距离II段退出
        searchSetting(img)
        TabDown(6, 0)
        # 距离III段退出
        searchSetting(img)
        TabDown(12, 0)
        return
    elif type == 5:
        # 距离I段退出
        img = 'dataFiles\\' + 'z1En.png'
        searchSetting(img)
        down(0)
        # 距离II段使能
        searchSetting(img)
        TabDown(6, 1)
        # 距离III段退出
        searchSetting(img)
        TabDown(12, 0)
        return
    elif type == 6:
        # 距离I段退出
        img = 'dataFiles\\' + 'z1En.png'
        searchSetting(img)
        down(0)
        # 距离II段退出
        searchSetting(img)
        TabDown(6, 0)
        # 距离III段使能
        searchSetting(img)
        TabDown(12, 1)
    elif type == 7:
        # 零序反时限使能
        img = 'dataFiles\\' + 'OCInverseEnableSetting.png'
        searchSetting(img)
        down(1)
        # 零序反时限电流定值
        img = 'dataFiles\\' + 'OCInverseValueSetting.png'
        searchSetting(img)
        down(0.04)
        # 零序反时限时间
        searchSetting(img)
        TabDown(6, 0.1)
        # 零序反时限配合时间
        searchSetting(img)
        TabDown(12, 0.1)
        # 零序反时限最小时间
        searchSetting(img)
        TabDown(18, 0)
    elif type == 8:
        pass
    elif type == 9:
        # 过负荷跳闸使能
        img = 'dataFiles\\' + 'OverloadTripEnable.png'
        searchSetting(img)
        down(1)
        # 过负荷跳闸定值
        img = 'dataFiles\\' + 'OverloadTripValue.png'
        searchSetting(img)
        down(0.04)
        # 过负荷跳闸时间
        searchSetting(img)
        TabDown(6, 0)
    elif type == 10:
        # 冲击性负荷使能
        img = 'dataFiles\\' + 'ImpactLoadEnable.png'
        searchSetting(img)
        down(1)
        # 差动动作电流定值
        img = 'dataFiles\\' + 'DifTripValue.png'
        searchSetting(img)
        down(10)
    elif type == 11:
        # 过电压跳本侧使能
        img = 'dataFiles\\' + 'OverVoltageTripEnable.png'
        searchSetting(img)
        down(1)
        searchSetting(img)
        TabDown(12, 1)
    elif type == 12:
        # 距离I段退出
        img = 'dataFiles\\' + 'z1En.png'
        searchSetting(img)
        down(0)
        # 距离II段使能
        searchSetting(img)
        TabDown(6, 1)
        # 距离III段使能
        searchSetting(img)
        TabDown(12, 1)
        # 单相重合闸使能
        img = 'dataFiles\\' + 'SingleRecEnable.png'
        searchSetting(img)
        down(1)
        searchSetting(img)
        TabDown(6, 0)
        searchSetting(img)
        TabDown(12, 0)
        searchSetting(img)
        TabDown(18, 0)
        # 2段保护闭锁重合闸
        img = 'dataFiles\\' + '2BlockRec.png'
        searchSetting(img)
        down(0)
        pass
        return


def PTSettingModify(type):
    sel_powertest()
    if type == 1:
        return
    elif type == 2:
        return
    elif type == 3:
        return


# 差动测试
def Dif():
    sel_word()
    pyperclip.copy('纵联差动保护')
    pyautogui.hotkey('alt', '2')
    time.sleep(ButtonReactTime)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    print('纵联差动保护测试\n')
    # ClearAllCheck()
    # SelectCheck(1100)
    SettingModify(1)
    for i in range(len(sheetDif['item'])):
        itemName = str(sheetDif.loc[i, 'item'])
        templateName = str(sheetDif.loc[i, 'template'])
        Start(templateName, itemName)
    # # 故障类型：AN；测时间；高定值
    # fileName = 'ANDifTimeHigh.tpl'
    # Start(fileName)
    # fileName = 'ANDifTimeLow.tpl'
    # Start(fileName)
    # fileName = 'BNDifTimeHigh.tpl'
    # Start(fileName)
    # fileName = 'BNDifTimeLow.tpl'
    # Start(fileName)
    # fileName = 'CNDifTimeHigh.tpl'
    # Start(fileName)
    # fileName = 'CNDifTimeLow.tpl'
    # Start(fileName)
    # # 故障类型：AN；测定值；低定值
    # fileName = 'ANDifValueLow.tpl'
    # Start(fileName)
    # fileName = 'BNDifValueLow.tpl'
    # Start(fileName)
    # fileName = 'CNDifValueLow.tpl'
    # Start(fileName)


def Start(templateName, itemName):
    global firstTime
    global ifG
    print('\0')
    print(itemName)
    # 打开测试模板
    OpenTpl(templateName)
    if (int(ifG) != 1) & (firstTime == 1):
        importSCD()
        firstTime = 0
    onllyStart()
    # 等待测试完成，时间固定
    time.sleep(TestWaitingTime)
    # 记录
    pyperclip.copy(itemName)
    sel_word()
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.hotkey('alt', '4')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    onllyRecord()
    ArpRecord()


# 零序过流测试
def OC():
    sel_word()
    pyperclip.copy('零序过流保护')
    pyautogui.hotkey('alt', '2')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    pyperclip.copy('II段')
    pyautogui.hotkey('alt', '3')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    print('零序过流保护测试\n')
    # 零序II段
    SettingModify(2)
    # 测时间
    # fileName = 'ANOC2Time.tpl'
    # Start(fileName)
    # fileName = 'BNOC2Time.tpl'
    # Start(fileName)
    # fileName = 'CNOC2Time.tpl'
    # Start(fileName)
    for i in range(len(sheetRocRelay['item'])):
        itemName = str(sheetRocRelay.loc[i, 'item'])
        templateName = str(sheetRocRelay.loc[i, 'template'])
        if re.search(r'ROC2', templateName) is not None:
            if re.search(r'Time', templateName) is not None:
                Start(templateName, itemName)
    # 测定值
    # 动作时间改为最小值
    ValueTestTimeSet()
    # 零序III段无单独控制字，将动作时间改为最大值
    img = 'dataFiles\\' + 'OC2TimeSetting.png'
    searchSetting(img)
    TabDown(12, 10)
    # 测试
    # fileName = 'ANOC2Value.tpl'
    # Start(fileName)
    # fileName = 'BNOC2Value.tpl'
    # Start(fileName)
    # fileName = 'CNOC2Value.tpl'
    # Start(fileName)
    for i in range(len(sheetRocRelay['item'])):
        itemName = str(sheetRocRelay.loc[i, 'item'])
        templateName = str(sheetRocRelay.loc[i, 'template'])
        if re.search(r'ROC2', templateName) is not None:
            if re.search(r'Value', templateName) is not None:
                Start(templateName, itemName)
    # 动作时间恢复
    ValueTestTimeSetRec()
    # 零序III段
    sel_word()
    pyperclip.copy('III段')
    pyautogui.hotkey('alt', '3')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    print('零序过流保护测试\n')
    SettingModify(3)
    # 测时间
    # fileName = 'ANOC3Time.tpl'
    # Start(fileName)
    # fileName = 'BNOC3Time.tpl'
    # Start(fileName)
    # fileName = 'CNOC3Time.tpl'
    # Start(fileName)
    for i in range(len(sheetRocRelay['item'])):
        itemName = str(sheetRocRelay.loc[i, 'item'])
        templateName = str(sheetRocRelay.loc[i, 'template'])
        if re.search(r'ROC3', templateName) is not None:
            if re.search(r'Time', templateName) is not None:
                Start(templateName, itemName)
    # 测定值
    ValueTestTimeSet()
    # 测试
    # fileName = 'ANOC3Value.tpl'
    # Start(fileName)
    # fileName = 'BNOC3Value.tpl'
    # Start(fileName)
    # fileName = 'CNOC3Value.tpl'
    # Start(fileName)
    for i in range(len(sheetRocRelay['item'])):
        itemName = str(sheetRocRelay.loc[i, 'item'])
        templateName = str(sheetRocRelay.loc[i, 'template'])
        if re.search(r'ROC3', templateName) is not None:
            if re.search(r'Value', templateName) is not None:
                Start(templateName, itemName)
    # 动作时间恢复
    ValueTestTimeSetRec()
    return


def ValueTestTimeSet():
    sel_arp()
    img = 'dataFiles\\' + 'settingSelected.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pass
    else:
        img = 'dataFiles\\' + 'setting.png'
        mouseClick(1, left, img, 1)
    # 零序II段时间
    img = 'dataFiles\\' + 'OC2TimeSetting.png'
    searchSetting(img)
    down(0.01)
    # 零序III段时间
    searchSetting(img)
    TabDown(12, 0.01)
    # 接地距离II段时间
    img = 'dataFiles\\' + 'Z2TimeSetting.png'
    searchSetting(img)
    down(0.01)
    # 接地距离III段时间
    searchSetting(img)
    TabDown(12, 0.01)
    # 相间距离II段时间
    searchSetting(img)
    TabDown(30, 0.01)
    # 相间距离III段时间
    searchSetting(img)
    TabDown(42, 0.01)
    return


def TabDown(TabTimes, value):
    i = 0
    while i < TabTimes:
        pyautogui.press('tab')
        i += 1
    down(value)


def ValueTestTimeSetRec():
    sel_arp()
    img = 'dataFiles\\' + 'settingSelected.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pass
    else:
        img = 'dataFiles\\' + 'setting.png'
        mouseClick(1, left, img, 1)
    # 零序II段时间
    img = 'dataFiles\\' + 'OC2TimeSetting.png'
    searchSetting(img)
    down(0.3)
    # 零序III段时间
    searchSetting(img)
    TabDown(12, 3)
    # 接地距离II段时间
    img = 'dataFiles\\' + 'Z2TimeSetting.png'
    searchSetting(img)
    down(0.5)
    # 接地距离III段时间
    searchSetting(img)
    TabDown(12, 1)
    # 相间距离II段时间
    searchSetting(img)
    TabDown(30, 0.5)
    # 相间距离III段时间
    searchSetting(img)
    TabDown(42, 1)
    return


# 距离测试
def Z():
    sel_word()
    pyperclip.copy('距离保护')
    pyautogui.hotkey('alt', '2')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    pyperclip.copy('I段')
    pyautogui.hotkey('alt', '3')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    print('零序过流保护测试\n')
    print('距离保护测试\n')
    i = 0
    while i < 3:
        if i == 1:
            sel_word()
            pyperclip.copy('II段')
            pyautogui.hotkey('alt', '3')
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.press('enter')
        elif i == 2:
            sel_word()
            pyperclip.copy('III段')
            pyautogui.hotkey('alt', '3')
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.press('enter')
        SettingModify(i + 4)
        # 测时间
        # fileName = 'ANZ' + str((i + 1)) + 'Time.tpl'
        # Start(fileName)
        # fileName = 'BNZ' + str((i + 1)) + 'Time.tpl'
        # Start(fileName)
        # fileName = 'CNZ' + str((i + 1)) + 'Time.tpl'
        # Start(fileName)
        # fileName = 'ABZ' + str((i + 1)) + 'Time.tpl'
        # Start(fileName)
        # fileName = 'BCZ' + str((i + 1)) + 'Time.tpl'
        # Start(fileName)
        # fileName = 'CAZ' + str((i + 1)) + 'Time.tpl'
        # Start(fileName)
        for j in range(len(sheetZone['item'])):
            itemName = str(sheetZone.loc[j, 'item'])
            templateName = str(sheetZone.loc[j, 'template'])
            if re.search(str(i + 1), templateName) is not None:
                if re.search(r'Time', templateName) is not None:
                    Start(templateName, itemName)
        ValueTestTimeSet()
        # 测定值
        # fileName = 'ANZ' + str((i + 1)) + 'Value.tpl'
        # Start(fileName)
        # fileName = 'BNZ' + str((i + 1)) + 'Value.tpl'
        # Start(fileName)
        # fileName = 'CNZ' + str((i + 1)) + 'Value.tpl'
        # Start(fileName)
        # fileName = 'ABCZ' + str((i + 1)) + 'Value.tpl'
        # Start(fileName)
        for j in range(len(sheetZone['item'])):
            itemName = str(sheetZone.loc[j, 'item'])
            templateName = str(sheetZone.loc[j, 'template'])
            if re.search(str(i + 1), templateName) is not None:
                if re.search(r'Value', templateName) is not None:
                    Start(templateName, itemName)
        ValueTestTimeSetRec()
        i += 1

    # # 距离I段
    # SettingModify(4)
    # fileName = 'ANZ1Time.tpl'
    # Start(fileName)
    # fileName = 'BNZ1Time.tpl'
    # Start(fileName)
    # fileName = 'CNZ1Time.tpl'
    # Start(fileName)
    # fileName = 'ABZ1Time.tpl'
    # Start(fileName)
    # fileName = 'BCZ1Time.tpl'
    # Start(fileName)
    # fileName = 'CAZ1Time.tpl'
    # Start(fileName)
    # # 距离II段
    # SettingModify(5)
    # fileName = 'ANZ2Time.tpl'
    # Start(fileName)
    # fileName = 'BNZ2Time.tpl'
    # Start(fileName)
    # fileName = 'CNZ2Time.tpl'
    # Start(fileName)
    # fileName = 'ABZ2Time.tpl'
    # Start(fileName)
    # fileName = 'BCZ2Time.tpl'
    # Start(fileName)
    # fileName = 'CAZ2Time.tpl'
    # Start(fileName)
    # # 距离III段
    # SettingModify(6)
    # fileName = 'ANZ3Time.tpl'
    # Start(fileName)
    # fileName = 'BNZ3Time.tpl'
    # Start(fileName)
    # fileName = 'CNZ3Time.tpl'
    # Start(fileName)
    # fileName = 'ABZ3Time.tpl'
    # Start(fileName)
    # fileName = 'BCZ3Time.tpl'
    # Start(fileName)
    # fileName = 'CAZ3Time.tpl'
    # Start(fileName)

    return


def ClearAllCheck():
    sel_powertest()
    i = 0
    while i < 6:
        img = 'dataFiles\\' + 'check.png'
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=1, interval=0.2, duration=0.2, button=left)
            pyautogui.press('enter')
        img = 'dataFiles\\' + 'check1.png'
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=1, interval=0.2, duration=0.2, button=left)
            pyautogui.press('enter')
        time.sleep(0.1)
        i += 1


def SelectCheck(para):
    i = int(para / 1000)
    j = int(para % 1000 / 100)
    k = int(para % 100 / 10)
    l = int(para % 10)
    img = 'dataFiles\\' + 'state1.png'
    mouseClick(1, left, img, 1)
    pyautogui.press('tab')
    if i:
        pyautogui.press('enter')
    i = 1
    while i < 5:
        pyautogui.press('tab')
        i += 1
    if j:
        pyautogui.press('enter')
    i = 1
    while i < 5:
        pyautogui.press('tab')
        i += 1
    if k:
        pyautogui.press('enter')
    i = 1
    while i < 5:
        pyautogui.press('tab')
        i += 1
    if l:
        pyautogui.press('enter')


def multiPress(button, times):
    i = 0
    while i < times:
        pyautogui.press(str(button))
        time.sleep(ButtonReactTime)
        i += 1


def OpenTpl(filename):
    """
    博电测试仪软件打开测试测试模板
    sel_powertest()
    pyautogui.press('alt')
    time.sleep(ButtonReactTime)
    pyautogui.press('f')
    time.sleep(ButtonReactTime)
    pyautogui.press('o')
    time.sleep(ButtonReactTime)
    """
    sel_onlly()
    time.sleep(ButtonReactTime)
    img = 'dataFiles\\' + 'onllyTaskbarFile.png'
    mouseClick(1, left, img, 1)
    multiPress('down', 2)
    multiPress('enter', 2)
    pyperclip.copy(filename)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.hotkey('alt', 'o')
    time.sleep(3)
    '''img =  'dataFiles\\' + 'OpenTPLConfirm.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    while location is not None:
        pyautogui.press('n')
        time.sleep(10)
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)'''


def importSCD():
    img = 'dataFiles\\' + 'gooseSub.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    while location is None:
        img = 'dataFiles\\' + 'PowerTestIEC.png'
        mouseClick(1, left, img, 1)
        time.sleep(5)
        img = 'dataFiles\\' + 'gooseSub.png'
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    mouseClick(1, left, img, 1)
    time.sleep(5)
    img = 'dataFiles\\' + 'importSCL.png'
    mouseClick(1, left, img, 1)
    time.sleep(5)
    # 选择文件目录，删除当前目录，输入目录地址
    pyautogui.press('f4')
    time.sleep(ButtonReactTime)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(ButtonReactTime)
    pyautogui.press('backspace')
    time.sleep(ButtonReactTime)
    pyperclip.copy(SCDAddr)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    # 选择文件名称
    pyautogui.hotkey('alt', 'n')
    time.sleep(ButtonReactTime)
    i = 0
    while i < 50:
        pyautogui.press('backspace')
        i += 1
    pyperclip.copy(SCDFileName)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    img = 'dataFiles\\' + 'PL2201.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pyautogui.click(location.x, location.y, clicks=1, interval=0.2, duration=0.2, button=left)
    # mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'emptyBlank.png'
    mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'gooseSub1.png'
    mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'GooseSubEmptyBlank.png'
    mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'importKey.png'
    mouseClick(1, left, img, 1)
    time.sleep(5)
    img = 'dataFiles\\' + 'importSCDConfirmBlue.png'
    mouseClick(1, left, img, 1)
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    img = 'dataFiles\\' + 'gooseTripA.png'
    mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'gooseTripA1.png'
    mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'gooseTripB.png'
    mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'gooseTripB1.png'
    mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'gooseTripC.png'
    mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'gooseTripC1.png'
    mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'gooseRec.png'
    mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'gooseRec1.png'
    mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'goosePub.png'
    mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'ChineseInput.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pyautogui.click(location.x, location.y, clicks=1, interval=0.2, duration=0.2, button=left)
    else:
        pass
    img = 'dataFiles\\' + 'gooseTWJ.png'
    while True:
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=1, interval=0.2, duration=0.2, button=left)
            pyautogui.press('[')
            pyautogui.press('1')
            pyautogui.press('0')
            pyautogui.press(']')
            time.sleep(ButtonReactTime)
        else:
            break
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    img = 'dataFiles\\' + 'importSCDConfirm.png'
    mouseClick(1, left, img, 1)
    time.sleep(ButtonReactTime)
    return


def PowerTestStart():
    pyautogui.press('f2')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    pyautogui.press('f2')
    time.sleep(ButtonReactTime)


def onllyStart():
    pyautogui.press('f3')
    time.sleep(ButtonReactTime)


def PowerTestRecord():
    sel_powertest()
    pyautogui.hotkey('ctrl', 'prntscrn')
    time.sleep(ButtonReactTime)
    sel_word()
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)


def onllyRecord():
    sel_onlly()
    pyautogui.hotkey('ctrl', 'prntscrn')
    time.sleep(ButtonReactTime)
    # pyautogui.hotkey('alt', 'space', 'c')
    # time.sleep(ButtonReactTime)
    sel_word()
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)


def ArpRecord():
    sel_arp()
    img = 'dataFiles\\' + 'tripLogSelected.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pass
    else:
        img = 'dataFiles\\' + 'tripLog.png'
        mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'ReportRefresh.png'
    mouseClick(30, left, img, 1)
    pyautogui.hotkey('ctrl', 'prntscrn')
    sel_word()
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')


def settingDownload():
    img = 'dataFiles\\' + 'settingDownload.png'
    mouseClick(1, left, img, 1)
    time.sleep(2)
    img = 'dataFiles\\' + 'ok.png'
    mouseClick(1, left, img, 1)
    # pyautogui.press('enter')
    # time.sleep(2)


def down(value):
    pyautogui.press('tab')
    time.sleep(ButtonReactTime)
    value = str(value)
    length = len(value)
    i = 0
    while i < length:
        if value[i] == "0":
            pyautogui.press('0')
        elif value[i] == "1":
            pyautogui.press('1')
        elif value[i] == "2":
            pyautogui.press('2')
        elif value[i] == "3":
            pyautogui.press('3')
        elif value[i] == "4":
            pyautogui.press('4')
        elif value[i] == "5":
            pyautogui.press('5')
        elif value[i] == "6":
            pyautogui.press('6')
        elif value[i] == "7":
            pyautogui.press('7')
        elif value[i] == "8":
            pyautogui.press('8')
        elif value[i] == "9":
            pyautogui.press('9')
        elif value[i] == ".":
            pyautogui.press('.')
        time.sleep(ButtonReactTime)
        i += 1
    settingDownload()


def searchSetting(target):
    img = 'dataFiles\\' + 'searchSettingInit.png'
    mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + target
    while True:
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=1, interval=0.2, duration=0.2, button=left)
            break
        else:
            i = 0
            while i < (6 * 13):
                pyautogui.press('tab')
                i += 1


def func_r_test():
    sel_word()
    pyperclip.copy('零序反时限过流保护')
    pyautogui.hotkey('alt', '2')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    print('零序反时限过流保护\n')
    GongNengRuanYaBan(1)
    SettingModify(7)
    # 测试开始
    # fileName = 'ANOCInverseTime.tpl'
    # Start(fileName)
    # fileName = 'BNOCInverseTime.tpl'
    # Start(fileName)
    # fileName = 'CNOCInverseTime.tpl'
    # Start(fileName)
    for i in range(len(sheetRocInv['item'])):
        itemName = str(sheetRocInv.loc[i, 'item'])
        templateName = str(sheetRocInv.loc[i, 'template'])
        Start(templateName, itemName)
    # 测试结束
    # 零序反时限使能关闭
    sel_arp()
    img = 'dataFiles\\' + 'settingSelected.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pass
    else:
        img = 'dataFiles\\' + 'setting.png'
        mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'OCInverseEnableSetting.png'
    searchSetting(img)
    down(0)

    return 0


def func_p_test():
    sel_word()
    pyperclip.copy('三相不一致保护')
    pyautogui.hotkey('alt', '2')
    time.sleep(ButtonReactTime)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    print('三相不一致保护\n')
    print('无法测试该功能\n')
    # GongNengRuanYaBan(0)
    SettingModify(8)
    return 0


def func_l_test():
    sel_word()
    pyperclip.copy('过流过负荷功能')
    pyautogui.hotkey('alt', '2')
    time.sleep(ButtonReactTime)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    print('过流过负荷功能\n')
    GongNengRuanYaBan(0)
    SettingModify(9)
    # 测试开始
    # fileName = 'ANOLTime.tpl'
    # Start(fileName)
    # fileName = 'BNOLTime.tpl'
    # Start(fileName)
    # fileName = 'CNOLTime.tpl'
    # Start(fileName)
    for i in range(len(sheetOverload['item'])):
        itemName = str(sheetOverload.loc[i, 'item'])
        templateName = str(sheetOverload.loc[i, 'template'])
        Start(templateName, itemName)
    # 测试结束
    # 过负荷使能关闭
    sel_arp()
    img = 'dataFiles\\' + 'settingSelected.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pass
    else:
        img = 'dataFiles\\' + 'setting.png'
        mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'OverloadTripEnable.png'
    searchSetting(img)
    down(0)
    return 0


def func_d_test():
    sel_word()
    pyperclip.copy('电铁、钢厂等冲击性负荷')
    pyautogui.hotkey('alt', '2')
    time.sleep(ButtonReactTime)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    print('电铁、钢厂等冲击性负荷\n')
    GongNengRuanYaBan(1110)
    SettingModify(10)
    # 测试开始
    # fileName = 'ANIL.tpl'
    # Start(fileName)
    # fileName = 'ANILCtrl.tpl'
    # Start(fileName)
    # fileName = 'BNIL.tpl'
    # Start(fileName)
    # fileName = 'BNILCtrl.tpl'
    # Start(fileName)
    # fileName = 'CNIL.tpl'
    # Start(fileName)
    # fileName = 'CNILCtrl.tpl'
    # Start(fileName)
    for i in range(len(sheetFload['item'])):
        itemName = str(sheetFload.loc[i, 'item'])
        templateName = str(sheetFload.loc[i, 'template'])
        Start(templateName, itemName)
    # 测试结束
    # 冲击性负荷使能关闭
    sel_arp()
    img = 'dataFiles\\' + 'settingSelected.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pass
    else:
        img = 'dataFiles\\' + 'setting.png'
        mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'ImpactLoadEnable.png'
    searchSetting(img)
    down(0)
    img = 'dataFiles\\' + 'DifTripValue.png'
    searchSetting(img)
    down(0.4)
    return 0


def func_y_test():
    sel_word()
    pyperclip.copy('远方跳闸保护')
    pyautogui.hotkey('alt', '2')
    time.sleep(ButtonReactTime)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    pyperclip.copy('过电压保护')
    pyautogui.hotkey('alt', '2')
    time.sleep(ButtonReactTime)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    print('过电压及远方跳闸保护\n')
    print('y型号只能测试过电压功能，不测试远方跳闸保护功能\n')
    GongNengRuanYaBan(0)
    SettingModify(11)
    # 测试
    # fileName = 'ANOVTime.tpl'
    # Start(fileName)
    # fileName = 'BNOVTime.tpl'
    # Start(fileName)
    # fileName = 'CNOVTime.tpl'
    # Start(fileName)
    for i in range(len(sheetOvervoltage['item'])):
        itemName = str(sheetOvervoltage.loc[i, 'item'])
        templateName = str(sheetOvervoltage.loc[i, 'template'])
        Start(templateName, itemName)
    # 测试结束
    # 过电压跳本侧使能关闭
    sel_arp()
    img = 'dataFiles\\' + 'settingSelected.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pass
    else:
        img = 'dataFiles\\' + 'setting.png'
        mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'OverVoltageTripEnable.png'
    searchSetting(img)
    down(0)
    return 0


def func_k_test():
    print('重合闸、零序加速及距离加速\n')
    sel_word()
    pyperclip.copy('零序加速')
    pyautogui.hotkey('alt', '2')
    time.sleep(ButtonReactTime)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    GongNengRuanYaBan(11)
    SettingModify(12)
    # 测重合闸
    # 零序后加速
    # fileName = 'ANRecOCAcc.tpl'
    # Start(fileName)
    # fileName = 'BNRecOCAcc.tpl'
    # Start(fileName)
    # fileName = 'CNRecOCAcc.tpl'
    # Start(fileName)
    for i in range(len(sheetAR['item'])):
        itemName = str(sheetAR.loc[i, 'item'])
        templateName = str(sheetAR.loc[i, 'template'])
        if re.search(r'Roc', templateName) is not None:
            Start(templateName, itemName)
    sel_word()
    pyperclip.copy('距离加速')
    pyautogui.hotkey('alt', '2')
    time.sleep(ButtonReactTime)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(ButtonReactTime)
    pyautogui.press('enter')
    time.sleep(ButtonReactTime)
    # 距离2段加速
    # fileName = 'ANRecZ2Acc.tpl'
    # Start(fileName)
    # fileName = 'BNRecZ2Acc.tpl'
    # Start(fileName)
    # fileName = 'CNRecZ2Acc.tpl'
    # Start(fileName)
    for i in range(len(sheetAR['item'])):
        itemName = str(sheetAR.loc[i, 'item'])
        templateName = str(sheetAR.loc[i, 'template'])
        if re.search(r'Zone', templateName) is not None:
            Start(templateName, itemName)
    # 距离3段加速
    # fileName = 'ANRecZ3Acc.tpl'
    # Start(fileName)
    # fileName = 'BNRecZ3Acc.tpl'
    # Start(fileName)
    # fileName = 'CNRecZ3Acc.tpl'
    # Start(fileName)
    # 测试结束
    # 禁止重合闸使能
    sel_arp()
    img = 'dataFiles\\' + 'settingSelected.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pass
    else:
        img = 'dataFiles\\' + 'setting.png'
        mouseClick(1, left, img, 1)
    img = 'dataFiles\\' + 'SingleRecEnable.png'
    searchSetting(img)
    down(0)
    searchSetting(img)
    TabDown(6, 0)
    searchSetting(img)
    TabDown(12, 1)
    searchSetting(img)
    TabDown(18, 0)
    return 0


def main(parameter):  # 参数说明---选配型号RPLDYK---例:100011即为-RYK
    # print(os.getcwd())
    # file = 'cmd.xls'
    # 打开文件
    # wb = xlrd.open_workbook(filename=file)
    # 通过索引获取表格sheet页
    # sheet1 = wb.sheet_by_index(0)
    print('welcome to AutoTest')
    # 装置类型
    # 1——G;2——DA；3——DG
    # a = input('输入装置类型:\n1——G;2——DA；3——DG\n')
    ifG = 1
    if int(ifG) != 1:
        # SCD文件目录
        SCDAddr = input('输入SCD文件地址:\n注意：SCD文件名称固定为"1.scd"\n')
    # key = input('输入选配型号RPLDYK:\n例:100011即为-RYK\n')
    key = parameter
    # SaveType = input('输入报告保存类型:\n1——图片保存;2——文字保存;3——图片加文字')
    key = int(key)
    func_r = int(key / 100000)
    func_p = int(key % 100000 / 10000)
    func_l = int(key % 10000 / 1000)
    func_d = int(key % 1000 / 100)
    func_y = int(key % 100 / 10)
    func_k = int(key % 10 / 1)
    print('Tests begin!\n')
    print('主保护测试\n')
    mainFunc()
    print('重合闸\n')
    if func_k == 0:
        TestWaitingTime = 50
        func_k_test()
        TestWaitingTime = 30
    print('选配功能测试\n')
    if func_r == 1:
        func_r_test()
    if func_p == 1:
        func_p_test()
    if func_l == 1:
        func_l_test()
    if func_d == 1:
        func_d_test()
    if func_y == 1:
        func_y_test()
