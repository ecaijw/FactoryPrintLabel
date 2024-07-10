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
btwFilepath = destFileFolder + "\\print.btw";
paramFilePath = destFileFolder + "\\print.txt"

logFilePath = destFileFolder + "\\print_log.csv"
version = '1.0.20040707'

PRODUCT_NUMBER_COUNT = 10


class MainFrame(wx.Frame):

    comboxName = None
    comboBoxType = None
    textPihao = None
    datePickerProduction = None
    textBoxCount = None
    datePickerValid = None
    staticTextStatus = None
    textProductNumber = []

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, -1, APP_TITLE)
        self.SetBackgroundColour(wx.Colour(204, 223, 254))
        self.SetSizeHints((1200, 800))
        self.btnSaveCalcInvest : wx.Button = None

        self.comboxName = None
        self.comboBoxType = None
        self.textPihao = None
        self.datePickerProduction = None
        self.textBoxCount = None
        self.datePickerValid = None
        self.staticTextStatus = None
        self.textTest = None
        self.textProductNumber = []

        # icon = wx.Icon(APP_ICON, wx.BITMAP_TYPE_ICO)
        # self.SetIcon(icon)

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
        self.comboxName = wx.ComboBox(self.panelRight, value="组合血糖仪", choices=["组合血糖仪"],
                                         style=wx.CB_READONLY, size=TEXT_SIZE)
        # HAPPY-100/300/500型组合血糖仪
        staticTextType  = wx.StaticText(self.panelRight, -1, label="产品型号：", size=STATIC_TEXT_SIZE)
        self.comboBoxType = wx.ComboBox(self.panelRight, value="HAPPY-100型", choices=["HAPPY-100型", "HAPPY-300型", "HAPPY-500型"],
                                         style=wx.CB_READONLY, size=TEXT_SIZE)
        staticTextName.SetFont(font)
        self.comboxName.SetFont(font)
        staticTextType.SetFont(font)
        self.comboBoxType.SetFont(font)
        row_sizer.Add(staticTextName,   0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        row_sizer.Add(self.comboxName, 1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, TEXT_GAP)  # 右对齐并留出间隔
        row_sizer.Add(staticTextType,   0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        row_sizer.Add(self.comboBoxType, 1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        # 将行添加到垂直布局中
        sizer.Add(row_sizer, 0, wx.ALL | wx.EXPAND, 5)  # 留出间隔并允许拉伸


        row_sizer = wx.BoxSizer(wx.HORIZONTAL)  # 水平布局，用于每行

        staticTextPihao  = wx.StaticText(self.panelRight, -1, label="产品批号：", size=STATIC_TEXT_SIZE)
        # 创建下拉列表，默认选择第一个
        self.textPihao        = wx.TextCtrl(self.panelRight, -1, size=TEXT_SIZE)
        self.textPihao.SetEditable(False)
        staticTextDate  = wx.StaticText(self.panelRight, -1, label="生产日期：", size=STATIC_TEXT_SIZE)
        # 创建日期选择控件，默认选择今天
        self.datePickerProduction = wx.adv.DatePickerCtrl(self.panelRight, dt=wx.DateTime.Now())

        staticTextPihao.SetFont(font)
        self.textPihao.SetFont(font)
        staticTextDate.SetFont(font)
        self.datePickerProduction.SetFont(font)
        row_sizer.Add(staticTextPihao,   0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        row_sizer.Add(self.textPihao, 1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, TEXT_GAP)  # 右对齐并留出间隔
        row_sizer.Add(staticTextDate,    0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        row_sizer.Add(self.datePickerProduction, 1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        # 将行添加到垂直布局中
        sizer.Add(row_sizer, 0, wx.ALL | wx.EXPAND, 5)  # 留出间隔并允许拉伸

        row_sizer = wx.BoxSizer(wx.HORIZONTAL)  # 水平布局，用于每行
        staticTextBoxCount = wx.StaticText(self.panelRight, -1, label="本箱数量：", size=STATIC_TEXT_SIZE)
        self.textBoxCount = wx.TextCtrl(self.panelRight, -1, size=TEXT_SIZE)
        staticTextDatePickerValid = wx.StaticText(self.panelRight, -1, label="有效日期：", size=STATIC_TEXT_SIZE)
        self.datePickerValid = self.createDateValid()
        staticTextBoxCount.SetFont(font)
        self.textBoxCount.SetFont(font)
        staticTextDatePickerValid.SetFont(font)
        self.datePickerValid.SetFont(font)
        row_sizer.Add(staticTextBoxCount,   0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        row_sizer.Add(self.textBoxCount,         1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, TEXT_GAP)  # 右对齐并留出间隔
        row_sizer.Add(staticTextDatePickerValid,   0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        row_sizer.Add(self.datePickerValid, 1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        # 将行添加到垂直布局中
        sizer.Add(row_sizer, 0, wx.ALL | wx.EXPAND, 5)  # 留出间隔并允许拉伸

        self.AddGap(sizer, 50)

        for i in range(PRODUCT_NUMBER_COUNT):
            self.textProductNumber.append(wx.TextCtrl(self.panelRight, -1, size=TEXT_SIZE))

        halfProductNumberCount = int(PRODUCT_NUMBER_COUNT / 2)
        # 每个设备的扫码信息
        for i in range(halfProductNumberCount):  # 遍历两次以创建两行
            row_sizer = wx.BoxSizer(wx.HORIZONTAL)  # 水平布局，用于每行

            # 静态文本
            static_text = wx.StaticText(self.panelRight, -1, label=f"产品编号 {i+1}：", size=STATIC_TEXT_SIZE)
            row_sizer.Add(static_text, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
            static_text.SetFont(font)

            # 输入框
            self.textProductNumber[i].SetFont(font)
            self.textProductNumber[i].Bind(wx.EVT_TEXT, self.onScanInfoChanged)
            row_sizer.Add(self.textProductNumber[i], 1, wx.FIXED_MINSIZE | wx.RIGHT, TEXT_GAP)  # 拉伸输入框并留出间隔

            # 静态文本
            static_text = wx.StaticText(self.panelRight, -1, label=f"产品编号 {i+1 + halfProductNumberCount}：", size=STATIC_TEXT_SIZE)
            row_sizer.Add(static_text, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
            static_text.SetFont(font)

            # 输入框
            self.textProductNumber[i + halfProductNumberCount].SetFont(font)
            self.textProductNumber[i + halfProductNumberCount].Bind(wx.EVT_TEXT, self.onScanInfoChanged)
            row_sizer.Add(self.textProductNumber[i + halfProductNumberCount], 1, wx.FIXED_MINSIZE | wx.RIGHT, 5)  # 拉伸输入框并留出间隔

            # 将行添加到垂直布局中
            sizer.Add(row_sizer, 0, wx.ALL | wx.EXPAND, 5)  # 留出间隔并允许拉伸

        # 状态文字
        self.AddGap(sizer, 50)
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)  # 水平布局，用于每行
        # 静态文本
        self.staticTextStatus = wx.StaticText(self.panelRight, -1, label=f"请进行扫码", size=STATIC_TEXT_SIZE)
        row_sizer.Add(self.staticTextStatus, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        self.staticTextStatus.SetFont(font)
        sizer.Add(row_sizer, 0, wx.ALL | wx.EXPAND, 5)  # 留出间隔并允许拉伸

        # 状态文字
        self.AddGap(sizer, 50)
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)  # 水平布局，用于每行
        # 静态文本
        self.textTest = wx.TextCtrl(self.panelRight, -1, value=f"测试信息", size=(800, 500), style=wx.TE_MULTILINE | wx.TE_READONLY)
        row_sizer.Add(self.textTest, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)  # 右对齐并留出间隔
        self.textTest.SetFont(font)
        sizer.Add(row_sizer, 0, wx.ALL | wx.EXPAND, 5)  # 留出间隔并允许拉伸

        # UI布局完成
        self.panelRight.SetSizer(sizer)

        self._mgr.Update()

    def onScanInfoChanged(self, event):
        triggered_control = event.GetEventObject()
        textScanInfo = event.GetString()

        if (triggered_control == self.textProductNumber[0]):
            t = datetime.now()

            print(f"onScanInfoChanged {t.microsecond}: ", textScanInfo)
            labelTest = f"({t.microsecond}): {event.GetString()}\n{self.textTest.GetValue()}"
            self.textTest.SetValue(labelTest)

        # 扫描枪的信息，可能分多次进入输入框
        if (len(textScanInfo) == 39):
            pihao = textScanInfo[26:34]
            lastThree = textScanInfo[-3:]
            self.staticTextStatus.SetLabel("扫码信息：" + textScanInfo)

            if (triggered_control == self.textProductNumber[0]):
                print("Text changed:", textScanInfo)
                print("pihao:", pihao)
                self.textPihao.SetValue(pihao)

            # use changeValue() to set the text, which does not trigger EVT_TEXT event
            triggered_control.ChangeValue(pihao + lastThree)

            print(f"onScanInfoChanged: SetFocus()")
            # 设置下一个产品编号接收输入
            textCtrlIndex = 0
            for textCtrlIndex in range(PRODUCT_NUMBER_COUNT - 1):
                if (triggered_control == self.textProductNumber[textCtrlIndex]):
                    self.textProductNumber[textCtrlIndex + 1].SetFocus()
            if (triggered_control == self.textProductNumber[PRODUCT_NUMBER_COUNT - 1]):
                self.textBoxCount.SetFocus()

    def createDateValid(self):

        # 获取今天的日期
        today = wx.DateTime.Now()

        # 计算5年后的日期（仅改变年份）
        # 注意：这里假设月份和日期在当前年份的5年后仍然有效
        # 如果不是，你可能需要添加额外的逻辑来调整月份和日期
        future_year = today.GetYear() + 5
        # print(f"{future_year}, {today.GetMonth() + 1}, {today.GetDay()}")
        try:
            future_date = wx.DateTime.FromDMY(today.GetDay(), today.GetMonth(), future_year)

        except wx._core.wxAssertionError as e:
            # 捕获 wxAssertionError 并处理它
            # 例如，你可以打印一条错误消息或执行一些恢复操作
            print(f"Caught wxAssertionError: {e}")
            # 检查日期是否有效（例如，2月29日在非闰年无效）
            future_date = wx.DateTime.FromDMY(1, today.GetMonth(), future_year)
            future_date.SetToLastMonthDay()  # 调整为该月的最后一天
        return wx.adv.DatePickerCtrl(self.panelRight, dt=future_date)

    def GetSelectedDate(self, datePicker):
        # 将wx.DateTime对象转换为datetime.date对象
        date = datePicker.GetValue()
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
        for textScanInfo in self.textProductNumber:
            print(f"产品编号 {index + 1}: {textScanInfo.GetValue()}")
            index += 1

    def checkBeforePrint(self):
        productNumberSet = set()

        for i in range(PRODUCT_NUMBER_COUNT):
            productNumber = self.textProductNumber[i].GetValue()
            if productNumber == "": # 产品编号空：没有进行扫描，不用进行检查
                continue
            if productNumber in productNumberSet:
                self.staticTextStatus.SetLabel(f"无法进行打印：产品编号不能相同。")
                return False
            productNumberSet.add(productNumber)
        return True

    def OnPrint(self, evt):
        print("main: OnPrint")
        self.printInfo()

        if not self.checkBeforePrint():
            return

        printText = f"{self.comboxName.GetValue()}, {self.comboBoxType.GetValue()}, " +\
                    f"{self.textPihao.GetValue()}, {self.GetSelectedDate(self.datePickerProduction)}, " +\
                    f"{self.textBoxCount.GetValue()}, {self.GetSelectedDate(self.datePickerValid)}"
        # 10个扫描信息。如果是空，保留逗号
        barCodeText= ""
        index = 0
        for textScanInfo in self.textProductNumber:
            printText = printText + f", {textScanInfo.GetValue()}"
            barCodeText = barCodeText + f"{textScanInfo.GetValue()} "
            index += 1

        printText = printText + "," + barCodeText
        with open(paramFilePath, "w", encoding="utf-8") as sr1:  # 使用with语句自动管理文件打开和关闭
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

    def print_method(self, btwFilePath, printText, copyCount = 1):
        try:
            for i in range(copyCount):
                # 使用subprocess运行bartend.exe，并传递参数
                # 注意：Python中的字符串格式化与C#略有不同
                cmd = [
                    f"C:\\Program Files (x86)\\Seagull\\BarTender Suite\\bartend.exe",
                    f"/AF={btwFilePath}",
                    "/P",
                    "/min=SystemTray"
                ]
                subprocess.Popen(cmd)  # 使用Popen来启动进程，不等待它完成
                self.log_method(printText)  # 调用日志方法
        except Exception as ex:
            print(f"Caught Exception: {ex}")
            print("Print Error")

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


