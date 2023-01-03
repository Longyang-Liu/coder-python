import tkinter as tk
import pyautogui
import time
import xlrd2
import math
from tkinter import messagebox


def getExcel(excelPath):
    # excel文件路径
    excelFilePath = excelPath
    # 指定sheet下标，从0开始
    sheetIndex = 0
    # 获取文本
    wb = xlrd2.open_workbook(excelFilePath)
    sheet = wb.sheet_by_index(sheetIndex)
    return sheet

def addFunc(addBotton, sleepTime):
    list = spliceStr(addBotton)
    # 点击加好友
    pyautogui.moveTo(int(list[0]), int(list[1]))
    pyautogui.click()
    time.sleep(sleepTime)

def inputQQ(qq, searchPosition, sleepTime):
    list = spliceStr(searchPosition)
    #清除输入框内
    pyautogui.press("backspace")
    #填写QQ号码
    pyautogui.typewrite(message=qq, interval=0.2)
    #点击搜索
    pyautogui.moveTo(int(list[0]), int(list[1]))
    pyautogui.click()
    time.sleep(sleepTime)

def clickAdd(addtion, sleepTime):
    list = spliceStr(addtion)
    #点击加好友按钮
    pyautogui.moveTo(int(list[0]), int(list[1]))
    pyautogui.click()
    time.sleep(sleepTime)

def nextStep(nextStepPosition, sleepTime):
    list = spliceStr(nextStepPosition)
    #点击下一步
    pyautogui.moveTo(int(list[0]), int(list[1]))
    pyautogui.click()
    time.sleep(sleepTime)
    pyautogui.click()
    pyautogui.moveTo(int(list[0]), int(list[1]))
    pyautogui.click()
    time.sleep(sleepTime)

def close(closeSmall, closeBig, sleepTime):
    listSmall = spliceStr(closeSmall)
    listBig = spliceStr(closeBig)
    # 关闭查找好友列表
    pyautogui.moveTo(int(listSmall[0]), int(listSmall[1]))
    pyautogui.click()
    pyautogui.moveTo(int(listBig[0]), int(listBig[1]))
    pyautogui.click()
    time.sleep(sleepTime)

def addLinker(excelPath, addBotton, searchPosition, addtion, nextStepPosition, closeSmall, closeBig, sleepTime):
    sheet = getExcel(excelPath)
    rows = sheet.nrows
    # 指定列，从0开始
    colIndex = 0
    for i in range(0, rows):
        if (type(sheet.cell(i, colIndex).value) == float):
            value = math.trunc(sheet.cell(i, colIndex).value)
        else:
            value = sheet.cell(i, colIndex).value
        qq = str(value)
        print(qq)
        addFunc(addBotton, sleepTime)
        inputQQ(qq, searchPosition, sleepTime)
        clickAdd(addtion, sleepTime)
        nextStep(nextStepPosition, sleepTime)
        close(closeSmall, closeBig, sleepTime)

def spliceStr(str):
    return str.split(',')

def guiPage():
    def LoginButton():
        excelPath = rt.filePath.get()
        addBotton = rt.addBotton.get()
        searchPosition = rt.searchPosition.get()
        addtion = rt.addtion.get()
        nextStepPosition = rt.nextStepPosition.get()
        closeSmall = rt.closeSmall.get()
        closeBig = rt.closeBig.get()
        sleepTime = rt.sleepTime.get()
        try:
            addLinker(excelPath, addBotton, searchPosition, addtion, nextStepPosition, closeSmall, closeBig, int(sleepTime))
        except Exception as e:
            messagebox.showerror('提示:', e)
    # 主窗口
    rt = tk.Tk("")
    rt.geometry('500x500')
    rt.title("AddLinker(by：lly)")
    # 变量
    rt.filePath = tk.StringVar()
    rt.addBotton = tk.StringVar()
    rt.searchPosition = tk.StringVar()
    rt.addtion = tk.StringVar()
    rt.nextStepPosition = tk.StringVar()
    rt.closeSmall = tk.StringVar()
    rt.closeBig = tk.StringVar()
    rt.sleepTime = tk.StringVar()
    # 账号
    f1 = tk.Frame(rt)
    tk.Label(f1, text='文件路径:  ').grid(row=0, column=0, padx=50)
    tk.Entry(f1, textvariable=rt.filePath).grid(row=0, column=1)
    f1.grid(pady=10)
    # 加好友按钮
    f2 = tk.Frame(rt)
    tk.Label(f2, text='加好友按钮坐标:  ').grid(row=1, column=0, padx=50)
    tk.Entry(f2, textvariable=rt.addBotton).grid(row=1, column=1)
    f2.grid(pady=10)
    # 查找按钮
    f3 = tk.Frame(rt)
    tk.Label(f3, text='查找按钮坐标:  ').grid(row=2, column=0, padx=50)
    tk.Entry(f3, textvariable=rt.searchPosition).grid(row=2, column=1)
    f3.grid(pady=10)
    # 查找按钮
    f3 = tk.Frame(rt)
    tk.Label(f3, text='添加按钮坐标:  ').grid(row=3, column=0, padx=50)
    tk.Entry(f3, textvariable=rt.addtion).grid(row=3, column=1)
    f3.grid(pady=10)
    # 下一步按钮
    f4 = tk.Frame(rt)
    tk.Label(f4, text='下一步按钮坐标:  ').grid(row=4, column=0, padx=50)
    tk.Entry(f4, textvariable=rt.nextStepPosition).grid(row=4, column=1)
    f4.grid(pady=10)
    # 下一步按钮
    f5 = tk.Frame(rt)
    tk.Label(f5, text='关闭小框按钮坐标:  ').grid(row=5, column=0, padx=50)
    tk.Entry(f5, textvariable=rt.closeSmall).grid(row=5, column=1)
    f5.grid(pady=10)
    # 下一步按钮
    f6 = tk.Frame(rt)
    tk.Label(f6, text='关闭大框按钮坐标:  ').grid(row=6, column=0, padx=50)
    tk.Entry(f6, textvariable=rt.closeBig).grid(row=6, column=1)
    f6.grid(pady=10)
    # 下一步按钮
    f7 = tk.Frame(rt)
    tk.Label(f7, text='睡眠时间:  ').grid(row=7, column=0, padx=50)
    tk.Entry(f7, textvariable=rt.sleepTime).grid(row=7, column=1)
    f7.grid(pady=10)
    # 登录按钮
    tk.Button(rt, text='执行', command=LoginButton).grid(pady=30)
    rt.mainloop()


def main():
    guiPage()


if __name__ == '__main__':
    main()