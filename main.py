from PyQt5.QtNetwork import QAbstractSocket, QTcpSocket, QHostAddress
from PyQt5.QtWebSockets import QWebSocket, QWebSocketServer
from PyQt5.QtCore import QUrl, QObject, QTimer, pyqtSignal
import sys, json
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QApplication,QSpacerItem, QSizePolicy, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QTextBrowser, QCheckBox, QHBoxLayout, QTextEdit, QPlainTextEdit
import xlwings as xw
import datetime

WEBSOCKET_PORT = 8080


def get_local_ip():
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.connect(("8.8.8.8", 80))
    if "192.168" in s.getsockname()[0]:
        return s.getsockname()[0]
    return ""

class WebSocketModel(QObject):
    update_signal = pyqtSignal()
    def __init__(self, parent=None):
        super().__init__(parent)
        self.config_data = {
            "boss_name": 'unknown',
            "max_times": 9999,
            "check_timeout": 6,
            "friend": '',
            "is_multiple": True,
            "is_draw": True,
        }
        self.excel = xw.App(visible=True,add_book=False)
        self.book = self.excel.books.add()
        self.sheet = None
        self.sheets = []
        self.logview = None
        self.count = 0
        self.lastrow = 0
        self.server = QWebSocketServer("WebSocket Server", QWebSocketServer.NonSecureMode, parent)
        self.server.newConnection.connect(self.handle_new_connection)
        self.clients = []

    def listen(self, port):
        address = QHostAddress.Any
        if not self.server.listen(address, port):
            print("Failed to listen on port", port)
        else:
            print("WebSocket server listening on port", port)
    
    def handle_new_connection(self):
        # if len(self.clients) >= 1:
        #     print("额外客户端连接.")
        #     return
        client = self.server.nextPendingConnection()
        client.binaryMessageReceived.connect(self.handle_message)
        client.textMessageReceived.connect(self.handle_text_message)
        client.disconnected.connect(self.handle_disconnect)
        self.clients.append(client)
        print("Client connected:", client.peerAddress().toString())
    
    def handle_text_message(self, message):
        sender = self.sender()
        print("Text message received from", sender.peerAddress().toString(), ":", message)
        c_data = json.loads(message)
        if c_data == '' or c_data == None:
            return
        if c_data['code'] == 1:
            if c_data['call'] == 'SyncHomeUI':
                self.config_data.update(c_data['data'])
                self.update_signal.emit()
            if c_data['call'] == 'getGameDataCurrent':
                self.sheet = self.book.sheets.add()
                self.sheets.append(self.sheet)
                header_data = c_data['data']['header']
                body_data = c_data['data']['body']
                self.sheet.range('A1:A3').merge()
                self.sheet.range('A1').value = header_data['golds']
                if header_data['golds'] >= 2000:
                    self.sheet.range('A1').color = (0,238,0)
                elif header_data['golds'] <= -2000:
                    self.sheet.range('A1').color = (238,0,0)

                self.sheet.range('B1:C1').merge()
                self.sheet.range('B2:C2').merge()
                self.sheet.range('B3:C3').merge()
                self.sheet.range('B1').value = '吃: ' + str(header_data['win'])
                self.sheet.range('B1').api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter 

                self.sheet.range('B2').value = '赔: ' + str(header_data['lose'])
                self.sheet.range('B2').api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter 

                self.sheet.range('B3').value = '平: ' + str(header_data['draw'])
                self.sheet.range('B3').api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter 

                self.sheet.range('B1:B3').color = (125,170,249)

                self.sheet.range('D1:E1').merge()
                self.sheet.range('D2:E2').merge()
                self.sheet.range('D3:E3').merge()
                self.sheet.range('D1').value = '本局次数: ' + str(header_data['current_count'])
                self.sheet.range('D1').color = (238,212,236)
                self.sheet.range('D2').value = '本局押注总计: ' + str(header_data['stake_count'])
                self.sheet.range('D2').color = (254,247,207)
                self.sheet.range('D3').value = '庄家本局结算: ' + str(header_data['stake_golds'])
                self.sheet.range('D3').color = (228,252,200)
                
                self.sheet.range('F1:G1').merge()
                self.sheet.range('F2:G2').merge()
                self.sheet.range('F3:G3').merge()
                self.sheet.range('F1').value = '押注上限: ' + str(header_data['max_times'])
                self.sheet.range('F1').color = (169,225,170)

                self.sheet.range('F2').value = '庄家姓名: ' + str(header_data['name'])
                self.sheet.range('F2').color = (212,186,227)
                self.sheet.range('F3').value = '庄家红包: ' + str(header_data['redp'])
                self.sheet.range('F3').color = (199,161,237)

                self.sheet.range('A4').value = '总单'
                self.sheet.range('A4:G4').color = (178, 178, 178)

                self.sheet.range('B4:C4').merge()
                self.sheet.range('B4').value = '参与人名'
                self.sheet.range('D4').value = '押注/收'
                self.sheet.range('E4').value = '红包点数'
                self.sheet.range('F4').value = '本局结算'
                self.sheet.range('G4').value = '收付款'

                # header A1:G4
                # header A5:??

                
                index = 0
                temp_num = header_data['golds']
                print("temp_num: %d", temp_num)
                while index < len(body_data):
                    row = index+5
                    self.sheet.range('A'+str(row)).value = body_data[index]['golds']
                    if body_data[index]['golds'] >= 2000:
                        self.sheet.range('A'+str(row)).color = (0,238,0)
                    elif body_data[index]['golds'] <= -2000:
                        self.sheet.range('A'+str(row)).color = (238,0,0)
                    self.sheet.range('B'+str(row)+':C'+str(row)).merge()
                    self.sheet.range('B'+str(row)).value = body_data[index]['name']

                    # self.sheet.range('B'+str(row)).api.Font.Color = (142, 124, 195)
                    if body_data[index]['isfake']:
                        self.sheet.range('B'+str(row)).api.Font.Color = (255, 153, 0)
                    if body_data[index]['islazy']:
                        self.sheet.range('B'+str(row)).api.Font.Color = (74, 134, 232)
                    if body_data[index]['isbug']:
                        self.sheet.range('B'+str(row)).api.Font.Color = (153, 0, 255)
                    
                    self.sheet.range('D'+str(row)).value = body_data[index]['chat']
                    self.sheet.range('E'+str(row)).value = body_data[index]['redp']
                    self.sheet.range('F'+str(row)).value = body_data[index]['current_golds']
                    # temp_num -= body_data[index]['current_golds']
                    index += 1
                temp_num -= header_data['stake_golds']
                print("temp_num_end: %s", str(temp_num))
                self.lastrow = index+5
                # 更新公式
                tmp = "=SUM(F5:F"+str(index+5)+")+"+str(temp_num)
                print("公式: %s" %tmp)
                self.sheet.range('A1').formula=tmp
                self.sheet.range('A5:'+str(index+5)).api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter 
            if c_data['call'] == 'sendMessage':
                print(c_data['data'])
                
                pass
        elif c_data['code'] == 2:
            
            pass
        print(c_data['call'])
    
    def makeExcelData(self):
        if self.sheet is None:
            return ""
        return {
            "header": self.sheet.range("A1:G4").value,
            "list": self.sheet.range("A5:G"+str(self.lastrow)).value,
        }

    def handle_message(self, message):
        sender = self.sender()
        print("Message received from", sender.peerAddress().toString(), ":", message)
        for client in self.clients:
            if client != sender:
                client.sendBinaryMessage(message)
    
    def handle_disconnect(self):
        sender = self.sender()
        print("Client disconnected:", sender.peerAddress().toString())
        
        # self.clients.remove(sender)
    
    def quit_all(self):
        print("全部退出")
        self.server.close()
        self.book.close()
        self.excel.quit()
        pass

class Controller():
    def __init__(self, model) -> None:
        self.model = model

    def clearConfig(self):
        # cleanDataForWeb
        message = {"code": 1, "call": "cleanDataForWeb", "data": ""}
        self.sendMessageAll(message)
        self.logv("清空配置")
        pass
    
    def saveConfig(self):
        message = {"code": 1, "call": "setHomeUIForWeb", "data": self.model.config_data}
        self.sendMessageAll(message)
        self.logv("保存配置")
        pass

    def sendMessageAll(self, data):
        if len(self.model.clients) <= 0:
            self.logv("没有客户端连接")
            return
        for client in self.model.clients:
            client.sendTextMessage(json.dumps(data))

    def startRunning(self, status):
        message = {"code": 1, "call": "changeCheckRunscript", "data": status}
        self.sendMessageAll(message)
        if status == True:
            self.logv("开始采集")
        else:
            self.logv("关闭采集")    
        
    def roobotRunning(self, status):
        message = {"code": 1, "call": "changeCheckRunrobot", "data": status}
        self.sendMessageAll(message)
        if status == True:
            self.logv("开始机器人")
        else:
            self.logv("关闭机器人")
        pass

    def openwx(self):
        self.logv("打开wx")
        message = {"code": 1, "call": "openWeChat", "data": ""}
        self.sendMessageAll(message)

    def openredp(self):
        self.logv("打开红包")
        message = {"code": 1, "call": "openRedp", "data": ""}
        self.sendMessageAll(message)
    
    def pushdata(self):
        self.logv("推送数据")
        pdata = self.model.makeExcelData()
        print(pdata)
        message = {"code": 1, "call": "pushData", "data": pdata}
        self.sendMessageAll(message)
        
    def logv(self, msg):
        if self.logview is not None:
            self.logview.append("[{}] {}".format(
                                datetime.datetime.now().strftime('%m-%d %H:%M'),
                                msg))
    
    def clearDataOne(self):
        self.logv("清空一把")
        message = {"code": 1, "call": "clearDataOne", "data": ""}
        self.model.sheet.delete()
        self.sendMessageAll(message)

    def clearDataAll(self):
        self.logv("清空所有")
        message = {"code": 1, "call": "clearDataAll", "data": ""}
        for sh in self.model.sheets:
            if sh is not None:
                sh.delete()
        self.sendMessageAll(message)

    def setLogView(self, view):
        self.logview = view

class MainView(QWidget):
    def create_views(self):
        # Create input widgets
        self.boss_name_edit = QLineEdit(self)
        self.max_times_edit = QLineEdit(self)
        self.check_timeout_edit = QLineEdit(self)
        self.friend_edit = QLineEdit(self)
        self.status_text = QLabel("v 1.0.0")
        self.myip = QLabel(get_local_ip()+":" + str(WEBSOCKET_PORT))

        self.is_multiple_checkbox = QCheckBox("翻倍模式", self)
        self.is_draw_checkbox = QCheckBox("一点庄吃", self)

        # Create buttons
        self.clear_button = QPushButton("清空配置", self)
        self.clear_button.clicked.connect(self.controller.clearConfig)

        self.save_button = QPushButton("保存配置", self)
        self.save_button.clicked.connect(self.controller.saveConfig)


        self.clear_data_one = QPushButton("清空一把", self)
        self.clear_data_one.clicked.connect(self.controller.clearDataOne)

        self.clear_data_all = QPushButton("清空所有", self)
        self.clear_data_all.clicked.connect(self.controller.clearDataAll)

        self.open_wx = QPushButton("打开wx", self)
        self.open_wx.clicked.connect(self.controller.openwx)

        self.open_redp = QPushButton("打开红包", self)
        self.open_redp.clicked.connect(self.controller.openredp)

        self.push_data = QPushButton("推送数据", self)
        self.push_data.clicked.connect(self.controller.pushdata)

        self.run_checkbox = QCheckBox("手动采集", self)
        self.run_checkbox.clicked.connect(self.controller.startRunning)

        self.roobot_checkbox = QCheckBox("自动核账", self)
        self.roobot_checkbox.clicked.connect(self.controller.roobotRunning)

        # Create text browser to display WebSocket data
        self.data_browser = QTextBrowser(self)
        self.controller.setLogView(self.data_browser)

    def create_layouts(self):
        # Set up layout
        main_layout = QVBoxLayout(self)
        child_layout = QHBoxLayout(self)
        main_layout.addLayout(child_layout)

        right_layout = QVBoxLayout(self)
        left_layout = QVBoxLayout(self)
        child_layout.addLayout(right_layout)
        child_layout.addLayout(left_layout)

        right_layout.addWidget(QLabel("庄家名字:"))
        right_layout.addWidget(self.boss_name_edit)
        right_layout.addWidget(QLabel("最大下注:"))
        right_layout.addWidget(self.max_times_edit)
        right_layout.addWidget(QLabel("机器人超时:"))
        right_layout.addWidget(self.check_timeout_edit)
        right_layout.addWidget(QLabel("自己人:"))
        right_layout.addWidget(self.friend_edit)
        right_check_layout = QHBoxLayout(self)
        
        right_check_layout.addWidget(self.is_multiple_checkbox)
        right_check_layout.addWidget(self.is_draw_checkbox)
        right_layout.addLayout(right_check_layout)

        left_layout.addWidget(self.data_browser)
        

        hbox1 = QHBoxLayout(self)
        hbox1.addWidget(self.clear_button)
        hbox1.addWidget(self.save_button)
        left_layout.addLayout(hbox1)

        hbox2 = QHBoxLayout(self)
        hbox2.addWidget(self.open_wx)
        hbox2.addWidget(self.open_redp)
        left_layout.addLayout(hbox2)
        left_layout.addWidget(self.push_data)

        hbox3 = QHBoxLayout(self)
        hbox3.addWidget(self.clear_data_one)
        hbox3.addWidget(self.clear_data_all)
        left_layout.addLayout(hbox3)
        
        left_check_layout = QHBoxLayout(self)
        left_check_layout.addWidget(self.run_checkbox)
        left_check_layout.addWidget(self.roobot_checkbox)
        left_layout.addLayout(left_check_layout)

        foolt_layout = QHBoxLayout(self)
        foolt_layout.addWidget(self.status_text)

        spacer_item = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        foolt_layout.addItem(spacer_item)
        foolt_layout.addWidget(self.myip)
        main_layout.addLayout(foolt_layout)

    def bind_view_module(self):
        self.controller.model.update_signal.connect(self.update_ui)
        self.boss_name_edit.textChanged.connect(lambda: self.update_config_data("boss_name", self.boss_name_edit.text()))
        self.max_times_edit.textChanged.connect(lambda: self.update_config_data("max_times", self.max_times_edit.text()))
        self.check_timeout_edit.textChanged.connect(lambda: self.update_config_data("check_timeout", self.check_timeout_edit.text()))
        self.friend_edit.textChanged.connect(lambda: self.update_config_data("friend", self.friend_edit.text()))
        self.is_multiple_checkbox.stateChanged.connect(lambda: self.update_config_data("is_multiple", self.is_multiple_checkbox.isChecked()))
        self.is_draw_checkbox.stateChanged.connect(lambda: self.update_config_data("is_draw", self.is_draw_checkbox.isChecked()))
        
    def update_ui(self):
        # Update UI with the latest config data
        self.boss_name_edit.setText(str(self.controller.model.config_data["boss_name"]))
        self.max_times_edit.setText(str(self.controller.model.config_data["max_times"]))
        self.check_timeout_edit.setText(str(self.controller.model.config_data["check_timeout"]))
        self.friend_edit.setText(str(self.controller.model.config_data["friend"]))
        self.is_multiple_checkbox.setChecked(self.controller.model.config_data["is_multiple"])
        self.is_draw_checkbox.setChecked(self.controller.model.config_data["is_draw"])
    
    def update_config_data(self, key, value):
        # Update the corresponding value in config_data
        self.controller.model.config_data[key] = value
        # Emit the update signal to trigger UI update
        self.controller.model.update_signal.emit()
    def initall(self):
        pass

    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.setWindowTitle("wx红包助手")
        self.create_views()
        self.create_layouts()
        self.bind_view_module()
        
        self.initall()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    server = WebSocketModel()
    controller = Controller(server)
    view = MainView(controller)
    server.listen(WEBSOCKET_PORT)
    view.show()
    print("websocket启动")
    app.aboutToQuit.connect(server.quit_all)
    sys.exit(app.exec_())
