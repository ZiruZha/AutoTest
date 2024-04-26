# *************************************
# ******此文件为myGui_V102.py库文件******
# *************************************
# Date:July,12th,2023
# 功能
# 1.覆盖原有cid文件
# 2.修改config.txt中管理序号、装置型号、选配功能
# 3.文件加密
# 4.打包
# JoiMaker使用说明
# 1.输入待制作joi包文件夹地址，文件夹名称需为标准格式，文件夹包含:
# 1.1同名文件夹，文件夹包含:
# 1.1.1未加密的原始文件
# 1.2icd文件
# 2.
# Update on Dec,6th,2023
# 修改说明
# DA型号装置joi包加入ccd文件，打包前需要先修改ccd文件中装置type
# Update on Mar,26th,2024
# Version-2.00
# 1.增加新功能:批量打包功能
# 功能说明，读excel，excel中包括1、装置型号2、版本号3、时间4、cq号
# Update on Apr,26th,2024
# Version-200_lib
# 修改说明
# 1. 删除主程序，改为def main()，以供GUI调用
# 2. 增加逻辑————若图片识别5次未找到，报错跳出main()，避免GUI按钮进入死循环
# 3. 所用的图片、模板全部打包放入dataFiles文件中
# 4. 所用的软件放入dataFiles文件中
# 5. 全部改用相对路径


import pyautogui
import time
import xlrd
import pyperclip
import re
from subprocess import run
import os
import pandas as pd

# import myGui

# 全局变量
# 测试等待时间
TestWaitingTime = 0.5
# joi文件Info
DeviceType = '0'
DeviceFunc = '0'
DeviceName = '0'
CQNumber = '0'
FileName = '0'
IfFuncR = '-1'
IfFuncP = '-1'
IfFuncL = '-1'
IfFuncD = '-1'
IfFuncY = '-1'
IfFuncK = '-1'
# 装置类型
# 1——G;2——DA；3——DG
ifG = 2
# 鼠标左右键设置
left = "right"
right = "left"
JoiAddr = ''


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


def sel_word():
    img = 'dataFiles\\' + 'ifWord.png'
    location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
    if location is not None:
        pyautogui.click(location.x, location.y, clicks=1, interval=0.2, duration=0.2, button=left)
    else:
        img = 'dataFiles\\' + '3.png'
        mouseClick(1, left, img, 1)
    time.sleep(0.1)


def run_cmd(cmd_str='', echo_print=1):
    """
    执行cmd命令，不显示执行过程中弹出的黑框
    备注：subprocess.run()函数会将本来打印到cmd上的内容打印到python执行界面上，所以避免了出现cmd弹出框的问题
    :param cmd_str: 执行的cmd命令
    :return:
    """

    if echo_print == 1:
        print('\n执行cmd指令="{}"'.format(cmd_str))
    run(cmd_str, shell=True)


def run_cmd_Popen_fileno(cmd_string):
    """
    执行cmd命令，并得到执行后的返回值，python调试界面输出返回值
    :param cmd_string: cmd命令，如：'adb devices'
    :return:
    """
    import subprocess

    print('运行cmd指令：{}'.format(cmd_string))
    return subprocess.Popen(cmd_string, shell=True, stdout=None, stderr=None).wait()


def ExtractInfo(Addr):
    global DeviceType, DeviceFunc, CQNumber, FileName, IfFuncR, IfFuncP, IfFuncL, IfFuncD, IfFuncY, IfFuncK, DeviceName
    findDeviceTypeStr = re.compile(r'303[A,C]-(.*)G[A-Z]{2}')
    findDeviceFuncStr = re.compile(r'G[A-Z]{2}-[A-Z]*')
    findCQNumberStr = re.compile(r'.....$')
    findFileNameStr = re.compile(r'\\N(.*)$')
    DeviceTypeStr = re.search(findDeviceTypeStr, Addr)
    DeviceFuncStr = re.search(findDeviceFuncStr, Addr)
    CQNumberStr = re.search(findCQNumberStr, Addr)
    FileNameStr = re.search(findFileNameStr, Addr)
    if DeviceTypeStr.group()[5:7:1] == 'DG':
        DeviceType = 'DG'
    elif DeviceTypeStr.group()[5:7:1] == 'DA':
        DeviceType = 'DA'
    else:
        DeviceType = 'G'
    if DeviceFuncStr is None:
        DeviceFunc = ''
        DeviceName = 'NSR-' + DeviceTypeStr.group()
    else:
        DeviceFunc = DeviceFuncStr.group()[4::1]
        DeviceName = 'NSR-' + DeviceTypeStr.group() + '-' + DeviceFuncStr.group()[4::1]
    findIfFuncR = re.compile(r'R')
    findIfFuncP = re.compile(r'P')
    findIfFuncL = re.compile(r'L')
    findIfFuncD = re.compile(r'D')
    findIfFuncY = re.compile(r'Y')
    findIfFuncK = re.compile(r'K')
    IfFuncR = re.search(findIfFuncR, DeviceFunc)
    IfFuncP = re.search(findIfFuncP, DeviceFunc)
    IfFuncL = re.search(findIfFuncL, DeviceFunc)
    IfFuncD = re.search(findIfFuncD, DeviceFunc)
    IfFuncY = re.search(findIfFuncY, DeviceFunc)
    IfFuncK = re.search(findIfFuncK, DeviceFunc)
    CQNumber = CQNumberStr.group()
    # if IfFuncR is None:
    #     IfFuncR = 0
    # else:
    #     IfFuncR = 1
    # if IfFuncP is None:
    #     IfFuncP = 0
    # else:
    #     IfFuncP = 1
    # if IfFuncL is None:
    #     IfFuncL = 0
    # else:
    #     IfFuncL = 1
    # if IfFuncD is None:
    #     IfFuncD = 0
    # else:
    #     IfFuncD = 1
    # if IfFuncY is None:
    #     IfFuncY = 0
    # else:
    #     IfFuncY = 1
    # if IfFuncK is None:
    #     IfFuncK = 0
    # else:
    #     IfFuncK = 1
    FileName = FileNameStr.group()[1::1]
    return 0


def CidReplace(Addr):
    # 删除原cid、icd文件
    TempStr = 'del /q ' + Addr + '\\' + FileName + '\\*.cid'
    os.system(TempStr)
    TempStr = 'del /q ' + Addr + '\\' + FileName + '\\*.icd'
    os.system(TempStr)
    # 将icd文件复制到目录中
    TempStr = 'copy ' + Addr + '\\' + '*.icd ' + Addr + '\\' + FileName
    os.system(TempStr)
    # 将icd文件重命名
    TempStr = 'rename ' + Addr + '\\' + FileName + '\\*.icd ' + 'configured.cid'
    os.system(TempStr)
    pass


def ConfigEdit(Addr):
    TempStr = Addr + '\\' + FileName + '\\' + 'config.txt'
    file = open(TempStr)
    ConfigText = file.read()
    replaceSUBQ = re.compile(r'SUBQ=30000\d{5}')
    ConfigText = re.sub(replaceSUBQ, ('SUBQ=30000' + CQNumber), ConfigText)
    replaceDeviceType = re.compile(r'DEVICE TYPE=.* DEVICE NAME')
    # print(ConfigText)
    # print(re.search(replaceDeviceType, ConfigText))
    ConfigText = re.sub(replaceDeviceType, ('DEVICE TYPE=' + DeviceName + ' DEVICE NAME'), ConfigText)

    replaceFuncR = re.compile(
        r'setswit_mod_rocinv                0x3f      %d   0      0      1       1       xxxx   xxxx   \d')
    replaceFuncP = re.compile(
        r'setswit_mod_pdr                   0x3f      %d   0      0      1       1       xxxx   xxxx   \d')
    replaceFuncL = re.compile(
        r'setswit_mod_overload              0x3f      %d   0      0      1       1       xxxx   xxxx   \d')
    replaceFuncD = re.compile(
        r'setswit_mod_fload                 0x3f      %d   0      0      1       1       xxxx   xxxx   \d')
    replaceFuncY = re.compile(
        r'setswit_mod_ov_yt                 0x3f      %d   0      0      1       1       xxxx   xxxx   \d')
    replaceFuncK = re.compile(
        r'setswit_multi_brk                 0x3f      %d   0      0      1       1       xxxx   xxxx   \d')
    if IfFuncR is None:
        ConfigText = re.sub(replaceFuncR,
                            'setswit_mod_rocinv                0x3f      %d   0      0      1       1       xxxx   xxxx   0',
                            ConfigText)
    else:
        ConfigText = re.sub(replaceFuncR,
                            'setswit_mod_rocinv                0x3f      %d   0      0      1       1       xxxx   xxxx   1',
                            ConfigText)
    if IfFuncP is None:
        ConfigText = re.sub(replaceFuncP,
                            'setswit_mod_pdr                   0x3f      %d   0      0      1       1       xxxx   xxxx   0',
                            ConfigText)
    else:
        ConfigText = re.sub(replaceFuncP,
                            'setswit_mod_pdr                   0x3f      %d   0      0      1       1       xxxx   xxxx   1',
                            ConfigText)
    if IfFuncL is None:
        ConfigText = re.sub(replaceFuncL,
                            'setswit_mod_overload              0x3f      %d   0      0      1       1       xxxx   xxxx   0',
                            ConfigText)
    else:
        ConfigText = re.sub(replaceFuncL,
                            'setswit_mod_overload              0x3f      %d   0      0      1       1       xxxx   xxxx   1',
                            ConfigText)
    if IfFuncD is None:
        ConfigText = re.sub(replaceFuncD,
                            'setswit_mod_fload                 0x3f      %d   0      0      1       1       xxxx   xxxx   0',
                            ConfigText)
    else:
        ConfigText = re.sub(replaceFuncD,
                            'setswit_mod_fload                 0x3f      %d   0      0      1       1       xxxx   xxxx   1',
                            ConfigText)
    if IfFuncY is None:
        ConfigText = re.sub(replaceFuncY,
                            'setswit_mod_ov_yt                 0x3f      %d   0      0      1       1       xxxx   xxxx   0',
                            ConfigText)
    else:
        ConfigText = re.sub(replaceFuncY,
                            'setswit_mod_ov_yt                 0x3f      %d   0      0      1       1       xxxx   xxxx   1',
                            ConfigText)
    if IfFuncK is None:
        ConfigText = re.sub(replaceFuncK,
                            'setswit_multi_brk                 0x3f      %d   0      0      1       1       xxxx   xxxx   0',
                            ConfigText)
    else:
        ConfigText = re.sub(replaceFuncK,
                            'setswit_multi_brk                 0x3f      %d   0      0      1       1       xxxx   xxxx   1',
                            ConfigText)
    # print(ConfigText)
    TempStr = JoiAddr + '\\' + FileName + '\\config.txt'
    with open(TempStr, 'w') as f:
        f.write(ConfigText)
    pass


def CcdEdit(Addr):
    # 修改type
    TempStr = Addr + '\\' + FileName + '\\' + 'configured.ccd'
    file = open(TempStr, encoding='UTF-8')
    CcdText = file.read()
    replaceSUBQ = re.compile(r'name="TEMPLATE" type=".*"')
    CcdText = re.sub(replaceSUBQ, ('name="TEMPLATE" type="' + DeviceName + '"'), CcdText)
    TempStr = JoiAddr + '\\' + FileName + '\\configured.ccd'
    with open(TempStr, 'w', encoding='UTF-8') as f:
        f.write(CcdText)
    pass


def Encryption(Addr):
    print('文件加密开始')
    # # 关闭encryption.exe目录
    # # os.system('taskkill /f /t /im encryption.exe')
    # # 打开encryption.exe目录
    # os.popen('start D:\\encrypt_security-V2.00')
    # time.sleep(TestWaitingTime)
    # # 运行encryption.exe
    # # 注意
    # # 不能使用cmd命令start D:\encrypt_security-V2.00\encryption.exe直接打开
    # # 控制台直接运行encryption.exe
    # pyautogui.hotkey('ctrl', 'a')
    # time.sleep(TestWaitingTime)
    # i = 0
    # while i < 6:
    #     pyautogui.press('down')
    #     time.sleep(TestWaitingTime)
    #     i += 1
    # pyautogui.press('enter')
    # time.sleep(5)
    # “加密文件夹选择”
    os.popen('encryption-new.exe')
    time.sleep(5)
    while 1:
        img = 'dataFiles\\' + 'encryptionDisable.png'
        location = pyautogui.locateCenterOnScreen(img, confidence=0.99)
        print(location)
        if location is None:
            print('encryptionDisable is None')
            img = 'dataFiles\\' + 'encryptionEnable.png'
            location = pyautogui.locateCenterOnScreen(img, confidence=0.99)
            if location is not None:
                print('encryptionEnable is not None')
                break
        os.system('taskkill /f /t /im encryption-new.exe')
        time.sleep(1)
        os.popen('encryption-new.exe')
        time.sleep(5)
    i = 0
    while i < 2:
        pyautogui.press('tab')
        time.sleep(TestWaitingTime)
        i += 1
    pyautogui.press('enter')
    time.sleep(TestWaitingTime)
    # 加密文件目录选择
    # i = 0
    # while i < 5:
    #     pyautogui.press('tab')
    #     time.sleep(TestWaitingTime)
    #     i += 1
    # pyautogui.press('enter')
    # time.sleep(TestWaitingTime)
    pyautogui.press('f4')
    time.sleep(TestWaitingTime)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(TestWaitingTime)
    TempStr = Addr + '/' + FileName
    pyperclip.copy(TempStr)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(TestWaitingTime)
    pyautogui.press('enter')
    time.sleep(TestWaitingTime)
    # 文件选择
    i = 0
    while i < 4:
        pyautogui.press('tab')
        time.sleep(TestWaitingTime)
        i += 1
    # 全选
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(TestWaitingTime)
    # 确定
    pyautogui.press('enter')
    time.sleep(5)
    pyautogui.press('enter')
    time.sleep(TestWaitingTime)
    # pyautogui.press('esc')
    # 关闭encryption.exe目录
    os.system('taskkill /f /t /im encryption-new.exe')
    time.sleep(1)
    # 在joi文件目录下新建encrytion文件夹
    TempStr = 'md ' + JoiAddr + '\\encryption'
    os.system(TempStr)
    TempStr = 'copy dataFiles\\encrypt_security-V2.00\\encryption\\*.* ' + JoiAddr + '\\encryption'
    os.system(TempStr)
    os.system('del /q dataFiles\\encrypt_security-V2.00\\encryption\\*.*')
    print('文件加密结束')
    pass


def Pack():
    print('joi打包开始')
    # pyautogui.hotkey('alt', 'tab')
    # time.sleep(TestWaitingTime)
    # pyautogui.press('esc')
    # time.sleep(TestWaitingTime)
    # os.system('taskkill /f /t /im ArpFilePack.exe')
    # 打开打包工具
    os.popen('ArpFilePack.exe')
    time.sleep(3)
    # 选择打包joi包
    i = 0
    while i < 2:
        pyautogui.press('tab')
        time.sleep(TestWaitingTime)
        i += 1
    pyautogui.press('enter')
    time.sleep(3)
    # 选择目录
    # i = 0
    # while i < 5:
    #     pyautogui.press('tab')
    #     time.sleep(TestWaitingTime)
    #     i += 1
    # pyautogui.press('enter')
    # time.sleep(TestWaitingTime)
    pyautogui.press('f4')
    time.sleep(TestWaitingTime)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(TestWaitingTime)
    # 输入加密文件地址
    TempStr = JoiAddr + '\\encryption'
    pyperclip.copy(TempStr)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(TestWaitingTime)
    pyautogui.press('enter')
    time.sleep(TestWaitingTime)
    # 选择打包文件
    i = 0
    while i < 4:
        pyautogui.press('tab')
        time.sleep(TestWaitingTime)
        i += 1
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(TestWaitingTime)
    pyautogui.press('enter')
    time.sleep(TestWaitingTime)
    # 选择下载版号
    i = 0
    while i < 6:
        pyautogui.press('tab')
        time.sleep(TestWaitingTime)
        i += 1
    if DeviceType == 'G':
        i = 0
        while i < 19:
            pyautogui.press('1')
            time.sleep(TestWaitingTime)
            pyautogui.press('down')
            time.sleep(TestWaitingTime)
            i += 1
        pyautogui.press('enter')
        time.sleep(TestWaitingTime)
    elif DeviceType == 'DG':
        i = 0
        while i < 17:
            pyautogui.press('1')
            time.sleep(TestWaitingTime)
            pyautogui.press('down')
            time.sleep(TestWaitingTime)
            i += 1
        i = 0
        while i < 2:
            pyautogui.press('1')
            time.sleep(TestWaitingTime)
            pyautogui.press('3')
            time.sleep(TestWaitingTime)
            pyautogui.press('down')
            time.sleep(TestWaitingTime)
            i += 1
        i = 0
        while i < 3:
            pyautogui.press('1')
            time.sleep(TestWaitingTime)
            pyautogui.press('down')
            time.sleep(TestWaitingTime)
            i += 1
        pyautogui.press('enter')
        time.sleep(TestWaitingTime)
    elif DeviceType == 'DA':
        i = 0
        while i < 17:
            pyautogui.press('1')
            time.sleep(TestWaitingTime)
            pyautogui.press('down')
            time.sleep(TestWaitingTime)
            i += 1
        i = 0
        while i < 2:
            pyautogui.press('4')
            time.sleep(TestWaitingTime)
            pyautogui.press('down')
            time.sleep(TestWaitingTime)
            i += 1
        i = 0
        while i < 4:
            pyautogui.press('1')
            time.sleep(TestWaitingTime)
            pyautogui.press('down')
            time.sleep(TestWaitingTime)
            i += 1
        pyautogui.press('enter')
        time.sleep(TestWaitingTime)
    # 打包
    pyautogui.hotkey('alt', 'o')
    time.sleep(3)
    # 选择保存目录
    # i = 0
    # while i < 6:
    #     pyautogui.press('tab')
    #     time.sleep(TestWaitingTime)
    #     i += 1
    # pyautogui.press('enter')
    # time.sleep(TestWaitingTime)
    pyautogui.press('f4')
    time.sleep(TestWaitingTime)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(TestWaitingTime)
    # 输入joi保存地址
    pyperclip.copy(JoiAddr)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(TestWaitingTime)
    pyautogui.press('enter')
    time.sleep(TestWaitingTime)
    # 选择保存名称
    # i = 0
    # while i < 6:
    #     pyautogui.press('tab')
    #     time.sleep(TestWaitingTime)
    #     i += 1
    pyautogui.press('tab')
    time.sleep(TestWaitingTime)
    pyautogui.hotkey('alt', 'n')
    time.sleep(TestWaitingTime)
    pyperclip.copy(FileName + '.joi')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(TestWaitingTime)
    pyautogui.press('enter')
    time.sleep(5)
    os.system('taskkill /f /t /im ArpFilePack.exe')
    print('joi打包结束')
    pass


def main(startNo, endNo):
    print('welcome to JoiMaker')
    # startNo = int(input('输入工程程序需求统计和交付管理.xlsx中待打包文件起始行号:\n'))
    # endNo = int(input('输入工程程序需求统计和交付管理.xlsx中待打包文件结束行号:\n'))
    startNo = int(startNo)
    endNo = int(endNo)
    # startNo = 82
    # endNo = 82
    # JoiAddr = input('输入joi文件夹地址:\n')
    # JoiAddr = 'C:\\Users\\admin\Desktop\\NSR-303A-GCN-RPLD_230711_V2.01_50425'
    excelPath = 'dataFiles\\' + '工程程序需求统计和交付管理.xlsx'
    df_1 = pd.read_excel(excelPath, sheet_name='Sheet1')  # 打开表格中第一个sheet，为df_1
    for i in range(startNo - 2, endNo - 1):
        if df_1.loc[i, '型号'] is not None and df_1.loc[i, '时间'] is not None and df_1.loc[
            i, 'configVersion'] is not None and df_1.loc[
            i, '管理序号'] is not None:
            # 读取excel中需要新生成的joi包信息
            fileName = str(df_1.loc[i, '型号']) + '_' + str(int(df_1.loc[i, '时间'])) + '_' + str(
                df_1.loc[i, 'configVersion']) + '_' + str(int(df_1.loc[i, '管理序号']))
            global JoiAddr
            desk = os.path.join(os.path.expanduser("~"), 'Desktop') + '\\'
            JoiAddr = desk + fileName
            ExtractInfo(JoiAddr)
            # 调试用
            print('Device Type:', DeviceType)
            print('Device Name:', DeviceName)
            print('Function:', DeviceFunc)
            print('IfFuncR:', IfFuncR)
            print('IfFuncP:', IfFuncP)
            print('IfFuncL:', IfFuncL)
            print('IfFuncD:', IfFuncD)
            print('IfFuncY:', IfFuncY)
            print('IfFuncK:', IfFuncK)
            print('CQ Number:', CQNumber)
            print('File Name:', FileName)
            # 在桌面新建根目录
            makeNewDir = 'mkdir ' + JoiAddr
            os.system(makeNewDir)
            # 在根目录下新建同名文件夹
            makeNewDir = 'mkdir ' + JoiAddr + '\\' + fileName
            os.system(makeNewDir)
            # 根据Device Type和IfFuncK选择程序包模板
            DAKTemplateAddr = 'dataFiles\\' + 'JoiMakerPrototype\\NSR-303A-DA-GCN-RLDYK'
            DATemplateAddr = 'dataFiles\\' + 'JoiMakerPrototype\\NSR-303A-DA-GCN-RPLDY'
            DGKTemplateAddr = 'dataFiles\\' + 'JoiMakerPrototype\\NSR-303A-DG-GCN-RLDYK'
            DGTemplateAddr = 'dataFiles\\' + 'JoiMakerPrototype\\NSR-303A-DG-GCN-RPLDY'
            GKTemplateAddr = 'dataFiles\\' + 'JoiMakerPrototype\\NSR-303A-GCN-RLDYK'
            GTemplateAddr = 'dataFiles\\' + 'JoiMakerPrototype\\NSR-303A-GCN-RPLDY'
            if DeviceType == 'DA':
                if IfFuncK is not None:
                    tempStr = DAKTemplateAddr
                else:
                    tempStr = DATemplateAddr
            elif DeviceType == 'DG':
                if IfFuncK is not None:
                    tempStr = DGKTemplateAddr
                else:
                    tempStr = DGTemplateAddr
            elif DeviceType == 'G':
                if IfFuncK is not None:
                    tempStr = GKTemplateAddr
                else:
                    tempStr = GTemplateAddr
            xcopyCommand = 'xcopy ' + tempStr + ' ' + JoiAddr + '\\' + fileName + '/y'
            os.system(xcopyCommand)
            # 复制icd文件
            DAKIcdInfoAddr = 'dataFiles\\' + '国网国产化ICD文件\\NSR-303A-DA-GCN-RLDYK'
            DAIcdInfoAddr = 'dataFiles\\' + '国网国产化ICD文件\\NSR-303A-DA-GCN-RPLDY'
            DGKIcdInfoAddr = 'dataFiles\\' + '国网国产化ICD文件\\NSR-303A-DG-GCN-RLDYK'
            DGIcdInfoAddr = 'dataFiles\\' + '国网国产化ICD文件\\NSR-303A-DG-GCN-RPLDY'
            GKIcdInfoAddr = 'dataFiles\\' + '国网国产化ICD文件\\NSR-303A-GCN-RLDYK'
            GIcdInfoAddr = 'dataFiles\\' + '国网国产化ICD文件\\NSR-303A-GCN-RPLDY'
            if DeviceType == 'DA':
                if IfFuncK is not None:
                    df_2 = pd.read_csv(DAKIcdInfoAddr + '\\NSR-303A-DA-GCN一个半开关.csv', encoding='GBK')
                else:
                    df_2 = pd.read_csv(DAIcdInfoAddr + '\\NSR-303A-DA-GCN双母接线.csv', encoding='GBK')
            elif DeviceType == 'DG':
                if IfFuncK is not None:
                    df_2 = pd.read_csv(DGKIcdInfoAddr + '\\20230203.csv', encoding='GBK')
                else:
                    df_2 = pd.read_csv(DGIcdInfoAddr + '\\20230203.csv', encoding='GBK')
            elif DeviceType == 'G':
                if IfFuncK is not None:
                    df_2 = pd.read_csv(GKIcdInfoAddr + '\\NSR-303A-GCN一个半接线.csv', encoding='GBK')
                else:
                    df_2 = pd.read_csv(GIcdInfoAddr + '\\NSR-303A-GCN双母接线.csv', encoding='GBK')
            for j in range(len(df_2['选配功能'])):
                if DeviceFunc == '':
                    if df_2['选配功能'].isna()[j]:
                        icdName = df_2.loc[j, 'ICD文件名称']
                        break
                if df_2.loc[j, '选配功能'] == DeviceFunc:
                    icdName = df_2.loc[j, 'ICD文件名称']
                    break
            if DeviceType == 'DA':
                if IfFuncK is not None:
                    xcopyCommand = 'copy ' + DAKIcdInfoAddr + '\\' + icdName + ' ' + JoiAddr
                else:
                    xcopyCommand = 'copy ' + DAIcdInfoAddr + '\\' + icdName + ' ' + JoiAddr
            elif DeviceType == 'DG':
                if IfFuncK is not None:
                    xcopyCommand = 'copy ' + DGKIcdInfoAddr + '\\' + icdName + ' ' + JoiAddr
                else:
                    xcopyCommand = 'copy ' + DGIcdInfoAddr + '\\' + icdName + ' ' + JoiAddr
            elif DeviceType == 'G':
                if IfFuncK is not None:
                    xcopyCommand = 'copy ' + GKIcdInfoAddr + '\\' + icdName + ' ' + JoiAddr
                else:
                    xcopyCommand = 'copy ' + GIcdInfoAddr + '\\' + icdName + ' ' + JoiAddr
            os.system(xcopyCommand)
            # 打包
            CidReplace(JoiAddr)
            ConfigEdit(JoiAddr)
            if DeviceType == 'DA' or DeviceType == 'DG':
                CcdEdit(JoiAddr)
            Encryption(JoiAddr)
            Pack()
        else:
            print('第', i, '行数据有问题')
