# -*- coding: utf-8 -*-
import sys
import os
import subprocess

curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)

import time

import wx
import wx.adv
import wx.lib.agw.aui as aui
import wx.lib.mixins.listctrl
import csv
from datetime import datetime

APP_TITLE = "精策工厂打印"
APP_ICON = "res/invest.ico"
destFileFolder = '''c:\\temp'''
btwFilepath = destFileFolder + "\\test.btw";
paramFilePath = destFileFolder + "\\param.txt"

logFilePath = destFileFolder + "\\print_log.csv"
version = '1.0.20040705'


class MainFrame(wx.Frame):

    textName = None
    textType = None
    comboBoxPihao = None
    datePickerCtrl = None
    textBoxCount = None
    textBoxId = None
    textScanInfo = []

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, -1, APP_TITLE)
        self.SetBackgroundColour(wx.Colour(204, 223, 254))
        self.SetSizeHints((1200, 800))
        self.btnSaveCalcInvest : wx.Button = None

        self.textName = None
        self.textType = None
        self.comboBoxPihao = None
        self.datePickerCtrl = None
        self.textBoxCount = None
        self.textBoxId = None
        self.textScanInfo = []

        icon = wx.Icon(APP_ICON, wx.BITMAP_TYPE_ICO)
        self.SetIcon(icon)

        # add panes, clock
        self.addPanes()

        # bind key-down event
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
        self.panelLeft.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)
        self.panelRight.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)


    def addPanes(self):
        self.panelLeft = wx.Panel(self, -1)
        self.panelRight = wx.Panel(self, -1)

        btnPrint= wx.Button(self.panelLeft, -1, u'打印标签', pos=(30, 150), size=(150, 60))
        btnPrint.Bind(wx.EVT_BUTTON, self.OnPrint)

        self._mgr = aui.AuiManager()
        self._mgr.SetManagedWindow(self)
        self._mgr.AddPane(self.panelLeft, aui.AuiPaneInfo().Name("LeftPanel").
                          Left().Layer(1).MinSize((200, -1)).Caption(u"操作区").MinimizeButton(True).MaximizeButton(True).CloseButton(True))

        self._mgr.AddPane(self.panelRight, aui.AuiPaneInfo().Name("CenterPanelInvest").
                          CenterPane().Show())

        STATIC_TEXT_SIZE = (150, 35)
        TEXT_SIZE = (250, 35)
        TEXT_GAP = 100
        font = wx.Font(18, wx.DECORATIVE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD, underline = False, faceName = "Monaco")

        '''
        产品名称     产品型号
        产品批号     生产日期
        本箱数量     本箱箱号
        '''
        # 设置 panelRightInvest 的布局
        sizer = wx.BoxSizer(wx.VERTICAL)  # 创建垂直布局管理器

        self.AddGap(sizer, 50)

        row_sizer = wx.BoxSizer(wx.HORIZONTAL)  # 水平布局，用于每行
        staticTextName  = wx.StaticText(self.panelRight, -1, label="产品名称：", size=STATIC_TEXT_SIZE)
        self.textName        = wx.TextCtrl(self.panelRight, -1, size=TEXT_SIZE)
        staticTextType  = wx.StaticText(self.panelRight, -1, label="产品型号：", size=STATIC_TEXT_SIZE)
        self.textType        = wx.TextCtrl(self.panelRight, -1, size=TEXT_SIZE)
        staticTextName.SetFont(font)
        self.textName.SetFont(font)
        staticTextType.SetFont(font)
        self.textType.SetFont(font)
        row_sizer.Add(staticTextName,   0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        row_sizer.Add(self.textName,         1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, TEXT_GAP)  # 右对齐并留出间隔
        row_sizer.Add(staticTextType,   2, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, TEXT_GAP)  # 右对齐并留出间隔
        row_sizer.Add(self.textType,         3, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        # 将行添加到垂直布局中
        sizer.Add(row_sizer, 0, wx.ALL | wx.EXPAND, 5)  # 留出间隔并允许拉伸


        row_sizer = wx.BoxSizer(wx.HORIZONTAL)  # 水平布局，用于每行

        staticTextPihao  = wx.StaticText(self.panelRight, -1, label="产品批号：", size=STATIC_TEXT_SIZE)
        # 创建下拉列表，列表项为['111', '222', '333']，默认选择第一个
        self.comboBoxPihao = wx.ComboBox(self.panelRight, value="111", choices=["111", "222", "333"],
                                         style=wx.CB_READONLY, size=TEXT_SIZE)
        staticTextDate  = wx.StaticText(self.panelRight, -1, label="生产日期：", size=STATIC_TEXT_SIZE)
        # 创建日期选择控件，默认选择今天
        self.datePickerCtrl = wx.adv.DatePickerCtrl(self.panelRight, dt=wx.DateTime.Now())

        staticTextPihao.SetFont(font)
        self.comboBoxPihao.SetFont(font)
        staticTextDate.SetFont(font)
        self.datePickerCtrl.SetFont(font)
        row_sizer.Add(staticTextPihao,   0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        row_sizer.Add(self.comboBoxPihao,1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, TEXT_GAP)  # 右对齐并留出间隔
        row_sizer.Add(staticTextDate,    2, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, TEXT_GAP)  # 右对齐并留出间隔
        row_sizer.Add(self.datePickerCtrl,3, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        # 将行添加到垂直布局中
        sizer.Add(row_sizer, 0, wx.ALL | wx.EXPAND, 5)  # 留出间隔并允许拉伸

        row_sizer = wx.BoxSizer(wx.HORIZONTAL)  # 水平布局，用于每行
        staticTextBoxCount = wx.StaticText(self.panelRight, -1, label="本箱数量：", size=STATIC_TEXT_SIZE)
        self.textBoxCount = wx.TextCtrl(self.panelRight, -1, size=TEXT_SIZE)
        staticTextBoxId = wx.StaticText(self.panelRight, -1, label="本箱箱号：", size=STATIC_TEXT_SIZE)
        self.textBoxId = wx.TextCtrl(self.panelRight, -1, size=TEXT_SIZE)
        staticTextBoxCount.SetFont(font)
        self.textBoxCount.SetFont(font)
        staticTextBoxId.SetFont(font)
        self.textBoxId.SetFont(font)
        row_sizer.Add(staticTextBoxCount,   0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        row_sizer.Add(self.textBoxCount,         1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, TEXT_GAP)  # 右对齐并留出间隔
        row_sizer.Add(staticTextBoxId,   2, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, TEXT_GAP)  # 右对齐并留出间隔
        row_sizer.Add(self.textBoxId,         3, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        # 将行添加到垂直布局中
        sizer.Add(row_sizer, 0, wx.ALL | wx.EXPAND, 5)  # 留出间隔并允许拉伸

        self.AddGap(sizer, 50)

        # 每个设备的扫码信息
        for i in range(10):  # 遍历两次以创建两行
            row_sizer = wx.BoxSizer(wx.HORIZONTAL)  # 水平布局，用于每行

            # 静态文本
            static_text = wx.StaticText(self.panelRight, -1, label=f"扫码信息 {i+1}：", size=STATIC_TEXT_SIZE)
            row_sizer.Add(static_text, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
            static_text.SetFont(font)

            # 输入框
            text_ctrl = wx.TextCtrl(self.panelRight, -1, size=TEXT_SIZE)
            text_ctrl.SetFont(font)
            self.textScanInfo.append(text_ctrl)
            row_sizer.Add(text_ctrl, 1, wx.FIXED_MINSIZE | wx.RIGHT, 5)  # 拉伸输入框并留出间隔

            # 将行添加到垂直布局中
            sizer.Add(row_sizer, 0, wx.ALL | wx.EXPAND, 5)  # 留出间隔并允许拉伸

        self.panelRight.SetSizer(sizer)

        self._mgr.Update()

    def GetSelectedDate(self):
        # 将wx.DateTime对象转换为datetime.date对象
        date = self.datePickerCtrl.GetValue()
        formatted_date = f"{date.year}年{date.month + 1:02d}月{date.day:02d}日"
        return formatted_date

    def AddGap(self, sizer, gapHeight):
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)  # 水平布局，用于每行
        staticGap = wx.StaticText(self.panelRight, -1, label="          ", size=(100, gapHeight))
        row_sizer.Add(staticGap,         0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        sizer.Add(row_sizer, 0, wx.ALL | wx.EXPAND, 5)  # 留出间隔并允许拉伸

    def OnKeyDown(self, evt):
        keyCode = evt.GetKeyCode()
        # print(keyCode)
        if (evt.GetKeyCode() == wx.WXK_F5):
            print("press F5")
        if (evt.GetKeyCode() == wx.WXK_F9):
            print("press F9")
        evt.Skip()

    def printInfo(self):
        index = 0
        for textScanInfo in self.textScanInfo:
            print(f"scanInfo {index + 1}: {textScanInfo.GetValue()}")
            index += 1

    def OnPrint(self, evt):
        print("main: OnPrint")
        self.printInfo()

        printText = f"{self.textName.GetValue()}, {self.textType.GetValue()}, " +\
                    f"{self.comboBoxPihao.GetValue()}, {self.GetSelectedDate()}, " +\
                    f"{self.textBoxCount.GetValue()}, {self.textBoxId.GetValue()}"
        index = 0
        for textScanInfo in self.textScanInfo:
            printText = printText + f", {textScanInfo.GetValue()}"
            index += 1

        with open(paramFilePath, "w") as sr1:  # 使用with语句自动管理文件打开和关闭
            sr1.write(printText)

        self.print_method(btwFilepath, printText)



    def log_method(self, printText):
        # 这里假设我们只是打印一个日志消息，实际中可以根据需要修改
        print("打印成功，写入日志记录")
        current_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        flag = '打印成功'

        # 打开文件，追加模式
        with open(logFilePath, 'a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            # 写入一行
            writer.writerow([current_date, version, flag, printText])

    def print_method(self, paramFilePath, printText, copyCount = 1):
        try:
            for i in range(copyCount):
                # 使用subprocess运行bartend.exe，并传递参数
                # 注意：Python中的字符串格式化与C#略有不同
                cmd = [
                    "bartend.exe",
                    f"/AF={paramFilePath}",
                    "/P",
                    "/min=SystemTray"
                ]
                subprocess.Popen(cmd)  # 使用Popen来启动进程，不等待它完成
                self.log_method(printText)  # 调用日志方法
        except Exception as ex:
            # 使用tkinter的messagebox显示错误信息
            print("Print Error")
            # root = tk.Tk()
            # root.withdraw()  # 隐藏主窗口
            # messagebox.showerror("Error", "打印异常，请检查打印机连接及打印模板文件", icon="warning")

class mainApp(wx.App):
    def OnInit(self):
        self.SetAppName(APP_TITLE)
        self.Frame = MainFrame(None)
        self.Frame.Show()
        self.Frame.Maximize(maximize=True)
        return True

if __name__ == "__main__":
    app = mainApp()
    app.MainLoop()


