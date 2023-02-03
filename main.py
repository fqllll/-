import wx
import os
import pynput.keyboard
from pynput.keyboard import Key
import time
import threading
import pandas as pd
import sys

if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']
    
class MyFrame(wx.Frame):
    def __init__(self, superior):
        wx.Frame.__init__(self, parent=superior, title="微信消息自动发送", size=(460, 600), pos=(100, 100))
        # 面板
        self.panel = wx.Panel(self)
        # 更换背景颜色
        self.panel.SetBackgroundColour('white')

        ##################################################

        # 操作时间间隔默认值
        self.INTERVAL_TIME = '0.5'

        ##################################################

        # 标题的设计
        title = wx.StaticText(parent=self.panel, label='微信消息自动发送', pos=(100, 20))
        title.SetForegroundColour('black')
        title.SetFont(wx.Font(20, wx.SWISS, wx.NORMAL, wx.NORMAL))

        ##################################################

        # 定时区域设计
        wx.StaticText(parent=self.panel, label='定时：', pos=(50, 60))

        self.combobox_chose_hour = wx.ComboBox(parent=self.panel, size=(40, 20),
                                               choices=['空'] + [str(i) for i in range(0, 24)],
                                               pos=(90, 60))
        wx.StaticText(parent=self.panel, label=':', pos=(135, 60))
        self.combobox_chose_min = wx.ComboBox(parent=self.panel, size=(40, 20),
                                              choices=['空'] + [
                                                  '0' + str(i) if (i in [j for j in range(0, 10)]) else str(i) for i in
                                                  range(0, 61)],
                                              pos=(145, 60))

        wx.StaticText(parent=self.panel, label='*没有设置定时或任一设置为空时默认立即发送', pos=(50, 80)).SetForegroundColour('red')

        ##################################################

        # 操作时间间隔区域设计
        wx.StaticText(parent=self.panel, label='操作时间间隔：', pos=(288, 60))
        self.textCtrl_interval_time = wx.TextCtrl(parent=self.panel, value=self.INTERVAL_TIME, size=(30, 20),
                                                  pos=(372, 60),
                                                  style=wx.TE_CENTER)
        wx.StaticText(parent=self.panel, label='s', pos=(405, 60))

        ##################################################

        # 状态栏区域设计
        wx.StaticText(parent=self.panel, label='状态栏：', pos=(50, 100))
        self.status_bar_text = wx.TextCtrl(parent=self.panel, value='''
                                                                    #########################################
                                                                    欢迎使用微信消息自动发送v1.0~~~
                                                                    #########################################
                                                                    操作提示：
                                                                    ①检查相关的配置
                                                                    （建议点击下方的’使用说明‘按钮查看详细的配置和操作过程）
                                                                    ②点击’编辑发送‘按钮，在打开的excel表格编辑发送内容
                                                                    ③点击’发送‘按钮
                                                                    #########################################
                                                                    注意：
                                                                    ①微信是否已经登录
                                                                    ②微信的快捷键是否修改过
                                                                    ③保证微信发送框内使用英文输入法（避免中英文混合时造成文本不一致）
                                                                    ④注意excel文档的编辑格式
                                                                    ⑤按下Esc按钮可以终止正在执行的发送操作
                                                                    #########################################
                                                                    '''.replace(' ',''),
                                           size=(360, 320),
                                           pos=(50, 120),
                                           style=wx.TE_MULTILINE)
        self.button_clear_status_bar = wx.Button(parent=self.panel, label='清空状态栏', size=(360, 30),
                                                 pos=(50, 440))  # 清空状态栏按钮

        ##################################################

        # 按钮区域设计
        self.button_open_instructions = wx.Button(parent=self.panel, label='使用说明', size=(100, 40),
                                                  pos=(50, 480))  # 使用说明 按钮
        self.button_edit_content = wx.Button(parent=self.panel, label='编辑发送内容', size=(100, 40),
                                             pos=(180, 480))  # 编辑发送发送 按钮
        self.button_send = wx.Button(parent=self.panel, label='发送', size=(100, 40), pos=(310, 480))  # 发送 按钮

        ##################################################

        # 建立按钮与事件的联系
        self.Bind(wx.EVT_BUTTON, self.OnButton_clear_bar_text, self.button_clear_status_bar)  # 清空状态栏按钮
        self.Bind(wx.EVT_BUTTON, self.OnButton_open_instructions, self.button_open_instructions)  # 打开使用说明
        self.Bind(wx.EVT_BUTTON, self.OnButton_edit_content, self.button_edit_content)  # 编辑发送
        self.Bind(wx.EVT_BUTTON, self.OnButton_send, self.button_send)  # 编辑发送

    ##################################################

    # 状态栏内分隔输出
    def separate_output(self, content):
        self.status_bar_text.AppendText('#########################################\n')
        self.status_bar_text.AppendText(content)
        self.status_bar_text.AppendText('#########################################\n\n')

    # 清空状态栏文本
    def OnButton_clear_bar_text(self, even):
        self.status_bar_text.SetValue('')  # 清空文本

    # 打开使用说明
    def OnButton_open_instructions(self, even):
        file_path = os.getcwd() + r'\data\使用说明.pdf'
        if os.path.exists(file_path):
            os.system(file_path)
            self.separate_output('成功打开使用说明的pdf文件！\n')
        else:
            self.separate_output("路径中的文件不存在，打开使用说明的相关文件失败！\n请检查当前目录下的data文件夹中使用说明.pdf是否存在！\n")

    # 打开编辑发送
    def OnButton_edit_content(self, even):
        file_path = os.getcwd() + r'\data\send_list.xls'
        if os.path.exists(file_path):
            os.system(file_path)
            self.separate_output('成功打开发送内容的xls文件！\n编辑完成后请记得保存~\n')
        else:
            self.separate_output("路径中的文件不存在，打开发送内容的相关文件失败！\n请检查当前目录下的data文件夹中send_list.xls是否存在！\n")

    # 发送信息
    def OnButton_send(self, even):
        file_path = os.getcwd() + r'\data\send_list.xls'
        if not os.path.exists(file_path):
            self.separate_output('发送内容的相关文件不存在！无法进行发送操作！\n')
            return
        # 获取操作时间间隔
        try:
            interval_time = float(self.textCtrl_interval_time.GetValue())
        except ValueError:
            self.separate_output('操作时间间隔的输入值需要数字，已将其恢复为默认值！\n')
            self.textCtrl_interval_time.SetValue(self.INTERVAL_TIME)
            return
        self.separate_output('当前设置的操作时间间隔为：' + str(interval_time) + 's\n')
        # 获取用户选择的发送时间点
        value_h = self.combobox_chose_hour.GetValue()
        value_min = self.combobox_chose_min.GetValue()
        # 实例化发送类
        send = Send(interval_time)
        # 对用户的选择的时间进行校验，对发送时刻作出判断
        if value_h in ['', '空'] or value_min in ['', '空']:
            self.separate_output('未设置定时，请您确保配置正确！3秒后立即为您发送消息！\n')
            for i in range(3, 0, -1):
                self.separate_output(str(i) + '！按下Esc键可终止发送操作！\n')
                time.sleep(1)
            self.separate_output('正在发送！按下Esc键可终止发送操作！\n')
            # 发送
            send.send()
            self.separate_output('发送完成！\n')
        else:
            # 目标时间
            target = value_h + ':' + value_min
            self.separate_output('您设置了定时发送！将于' + target + '为您发送！\n' + '请检查相关配置，并在接近发送时刻暂时停下对电脑的操作！\n')
            while True:
                # 获取当前时间
                now = time.strftime("%H:%M", time.localtime())
                if target == now:
                    self.separate_output('已到达目标时间：' + target + '！\n开始发送！\n')
                    # 发送
                    send.send()
                    break
                time.sleep(1)
            self.separate_output('发送完成！\n')


class Send:
    def __init__(self, interval_time):
        self.keyboard = pynput.keyboard.Controller()

        # 控制操作间隔时间
        self.interval_time = interval_time
        # 控制主线程是否结束
        self.isEnd = False
        # 存放发送人昵称和发送内容的列表
        self.nickname_list = []
        self.content_list = []

    # 按下和释放按键方法，定义接受多值参数以方便不同长度组合键的使用
    def __press_release_key(self, *input_key):
        for i in range(len(input_key)):
            self.keyboard.press(input_key[i])
        for i in range(len(input_key) - 1, -1, -1):
            self.keyboard.release(input_key[i])
        time.sleep(self.interval_time)

    # 检测键盘是否按下Esc的线程方法
    def __keyboard_esc(self, key):
        if key == Key.esc:
            self.isEnd = True
            return False

    # 发送微信信息的线程方法
    def __send_message_thread(self):
        # 读取
        flag = self.__read_file()
        # 读取的发送内容文件没有问题才执行发送操作
        if flag:
            # 利用快捷键打开微信
            self.__press_release_key(Key.ctrl, Key.alt, 'w')
            for i in range(len(self.nickname_list)):
                # 利用快捷键搜索发送人
                self.__press_release_key(Key.ctrl, 'f')
                self.keyboard.type(self.nickname_list[i])
                time.sleep(self.interval_time)
                self.__press_release_key(Key.enter)
                # 将内容输入到微信输入框并发送
                self.keyboard.type(self.content_list[i])
                time.sleep(self.interval_time)
                self.__press_release_key(Key.enter)
        self.isEnd = True
        return False

    # 读取发送内容文件
    def __read_file(self):
        try:
            df = pd.read_excel(os.getcwd() + r'\data\send_list.xls', header=None)
            self.nickname_list = df[0].tolist()
            self.content_list = df[1].tolist()
            # 返回是否存在nan值
            return not df.isnull().values.any()
        except:
            # 存在错误
            return False

    # 启动线程
    def __start_thread(self):
        # 创建键盘监听线程
        listener = pynput.keyboard.Listener(on_press=self.__keyboard_esc)
        # 设置线程为“不重要”，随随主进程结束而结束
        listener.daemon = 1

        # 创建发送信息线程
        t2 = threading.Thread(target=self.__send_message_thread)
        # 设置线程为“不重要”，随随主进程结束而结束
        t2.daemon = 1

        # 启动线程
        listener.start()
        t2.start()
        while True:
            if self.isEnd:
                return

    def send(self):
        self.__start_thread()

app = wx.App()
frame = MyFrame(None)
frame.Show(True)
app.MainLoop()
