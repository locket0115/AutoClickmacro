import sys
import pyautogui
import time
import logging
import pickle

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import QColor
from PyQt5 import QtCore, QtWidgets
from pynput import keyboard


from excel import *
from data import *


# 루트 로거 생성
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)  # 모든 로그 메시지를 처리하기 위해 DEBUG 레벨로 설정

# 로그 파일 (logger.log) 핸들러 설정
file_handler = logging.FileHandler('logger.log')
file_handler.setLevel(logging.INFO)  # 모든 로그를 기록

# 에러 로그 파일 (error.log) 핸들러 설정
error_handler = logging.FileHandler('error.log')
error_handler.setLevel(logging.ERROR)  # ERROR 레벨 이상만 기록

# 콘솔 핸들러 설정
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)  # 모든 로그를 콘솔에 출력

# 로그 포맷 설정
formatter = logging.Formatter('%(asctime)s  %(levelname)s: %(message)s')
file_handler.setFormatter(formatter)
error_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# 핸들러 추가
logger.addHandler(file_handler)
logger.addHandler(error_handler)
logger.addHandler(console_handler)


# 예외처리
def exception_handler(exc_type, exc_value, exc_traceback):
    logger.critical("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))

    sys.__excepthook__(exc_type, exc_value, exc_traceback)



class OrderThread(QThread):
    progress_signal = pyqtSignal(str)  # 주문 진행 상태 전달 신호
    finished_signal = pyqtSignal()  # 주문 완료 신호

    highlightExcelTable_signal = pyqtSignal(int, QColor) # 주문중인 행 강조 신호
    highlightCommandWidget_signal = pyqtSignal(QWidget, str) # 진행중인 커맨드 강조 신호

    def __init__(self, showExcelData, showCommand):
        super().__init__()
        self.showExcelData = showExcelData
        self.showCommand = showCommand
        
        self.running = True
        
        self.mutex = QMutex()
        self.condition = QWaitCondition()
        self.paused = False

    def run(self):
        if self.showExcelData.count() == 0: # 엑셀 파일이 불러오지 않은 경우
            self.progress_signal.emit("엑셀 파일을 불러와주세요.")
            logger.error('No Excel tab')
            return

        if self.showCommand.count() == 0: # 커맨드 탭이 없는 경우
            self.progress_signal.emit("커맨드 탭을 추가해주세요.")
            logger.error('No Command Tab')
            return

        # 데이터 가져오기
        current_table = self.showExcelData.currentWidget()

        stock = self.showExcelData.tabText(self.showExcelData.currentIndex())
        logger.debug(f"Stock: {stock}")


        for row in range(current_table.rowCount()): # 주문 row에 대하여
            self.mutex.lock()
            if self.paused: # 일시정지 체크
                self.progress_signal.emit('주문 일시정지됨.')
                self.condition.wait(self.mutex) 

            if not self.running: # 정지 체크
                self.progress_signal.emit("주문 정지됨.")
                return
            
            self.mutex.unlock()

            self.highlightExcelTable_signal.emit(row, QColor(128, 128, 128)) # 강조



            bs = True if current_table.item(row, 0).text() == '매수' else False
            price = float(current_table.item(row, 1).text())
            method = current_table.item(row, 2).text()
            quantity = int(current_table.item(row, 3).text())

            logger.debug(f"Order: {'매수' if bs else '매도'}, {price}, {method}, {quantity}")

            # 커맨트 탭 선택
            command_tab = None
    
            for idx in range(self.showCommand.count()): # 구매 방식 탭
                if self.showCommand.tabText(idx) == f"{'매수' if bs else '매도'}_{method}":
                    command_tab = self.showCommand.widget(idx).layout().itemAt(0).widget()
                    self.showCommand.setCurrentIndex(idx)

                    logger.debug(f"tab {self.showCommand.tabText(idx)} used.")

                    break
            
            if command_tab == None:
                for idx in range(self.showCommand.count()): # 구매 방식이 없을 경우 기본 탭 사용
                    if self.showCommand.tabText(idx) == f"{'매수' if bs else '매도'}":
                        command_tab = self.showCommand.widget(idx).layout().itemAt(0).widget()
                        self.showCommand.setCurrentIndex(idx)

                        logger.debug(f"tab {self.showCommand.tabText(idx)} used.")

                        break

            if command_tab == None: #기본 탭도 없을 경우 현재 탭 사용
                command_tab = self.showCommand.currentWidget().layout().itemAt(0).widget()
                logger.debug(f"tab {self.showCommand.tabText(self.showCommand.currentIndex())} used.")


            current_command = command_tab.widget()


            for command in current_command.children(): # 커맨드 탭에 대하여
                self.mutex.lock()
                if self.paused: # 일시정지 체크
                    self.progress_signal.emit('주문 일시정지됨.')
                    self.condition.wait(self.mutex) 

                if not self.running: # 정지 체크
                    self.progress_signal.emit("주문 정지됨.")
                    return
                
                self.mutex.unlock()



                if isinstance(command, QWidget):
                    # command.setVisible(False)
                    # command.setStyleSheet("QWidget { background-color: white; }")
                    # command.setVisible(True)
                    self.highlightCommandWidget_signal.emit(command, "QWidget { background-color: white; }")

                    if command.layout().itemAt(0).widget().currentText() == "커서 이동":
                        try:
                            x = int(command.layout().itemAt(1).widget().text())
                            y = int(command.layout().itemAt(2).widget().text())
                            logger.debug(f"Move to: {x}, {y}")

                            pyautogui.moveTo(x, y)
                        except Exception as e:
                            logger.warning(f"Error, {e}")

                    elif command.layout().itemAt(0).widget().currentText() == "마우스 클릭":
                        logger.debug("Click")
                        pyautogui.click()

                    elif command.layout().itemAt(0).widget().currentText() == "키보드 입력":
                        if command.layout().itemAt(5).widget().currentText() == "테이블에서 가져오기":
                            logger.debug(f"Type: {command.layout().itemAt(6).widget().currentText()}")

                            idx = command.layout().itemAt(6).widget().currentText()

                            if idx == '종목':
                                pyautogui.typewrite(stock)
                            elif idx == '주문가':
                                pyautogui.typewrite(str(price))
                            elif idx == '주문 방법':
                                pyautogui.typewrite(method)
                            elif idx == '수량':
                                pyautogui.typewrite(str(quantity))
                        else:
                            logger.debug(f"""Type: '{command.layout().itemAt(7).widget().text()}'""")
                            pyautogui.typewrite(command.layout().itemAt(7).widget().text())
                    else:
                        logger.warning("Invalid Command")

                    time.sleep(0.1)

                    # command.setVisible(False)
                    # command.setStyleSheet("QWidget { background-color: none; }")
                    # command.setVisible(True)

                    self.highlightCommandWidget_signal.emit(command, "QWidget { background-color: none; }")

                    logger.info(f'command completed')

            logger.info(f'command tab completed')

            
            self.highlightExcelTable_signal.emit(row, QColor(255, 255, 255)) # 강조 해제
            
        logger.info(f'order completed')
        self.progress_signal.emit("주문 완료")

    def stop(self):
        self.running = False
        self.finished_signal.emit()  # 스레드 종료 신호

        logger.info("Order Stopped")

    def pause(self):
        self.mutex.lock()
        self.paused = True
        self.mutex.unlock()

        logger.info("Order Paused")

    def resume(self):
        self.mutex.lock()
        self.paused = False
        self.condition.wakeAll()  # 스레드를 재개
        self.mutex.unlock()
        
        logger.info("Order Resumed")


    def isPaused(self):
        return self.paused


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.excel = Excel()

        self.FileDirectory.setText("파일 경로")
        self.FileLoad.clicked.connect(self.btn_FileLoad)
        self.showExcelData = self.findChild(QTabWidget, 'showExcelData')
        self.showCommand = self.findChild(QTabWidget, 'showCommand')
        self.addCommandTab.clicked.connect(self.btn_addCommandTab)
        self.copyCommandTab.clicked.connect(self.btn_copyCommandTab)
        self.makeOrder.clicked.connect(self.btn_makeOrder)
        self.clearOrder.clicked.connect(self.btn_clearOrder)

        self.actionSave.triggered.connect(self.saveCommand)

        self.loadCommand()

        self.order_thread = None

        # pynput 키보드 리스너 설정
        self.listener = keyboard.Listener(on_press=self.on_press)
        self.listener.start()


    def setupUi(self, mainWindow):
        mainWindow.setObjectName("mainWindow")
        mainWindow.resize(1210, 561)
        self.centralwidget = QtWidgets.QWidget(mainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.FileLoad = QtWidgets.QPushButton(self.centralwidget)
        self.FileLoad.setGeometry(QtCore.QRect(10, 500, 131, 31))
        self.FileLoad.setToolTip("")
        self.FileLoad.setObjectName("FileLoad")
        self.showExcelData = QtWidgets.QTabWidget(self.centralwidget)
        self.showExcelData.setGeometry(QtCore.QRect(0, 10, 431, 481))
        self.showExcelData.setAutoFillBackground(False)
        self.showExcelData.setObjectName("showExcelData")
        self.showCommand = QtWidgets.QTabWidget(self.centralwidget)
        self.showCommand.setGeometry(QtCore.QRect(450, 10, 761, 481))
        self.showCommand.setObjectName("showCommand")
        self.makeOrder = QtWidgets.QPushButton(self.centralwidget)
        self.makeOrder.setGeometry(QtCore.QRect(1070, 500, 131, 31))
        self.makeOrder.setObjectName("makeOrder")
        self.FileDirectory = QtWidgets.QLabel(self.centralwidget)
        self.FileDirectory.setGeometry(QtCore.QRect(150, 500, 281, 20))
        self.FileDirectory.setObjectName("FileDirectory")
        self.addCommandTab = QtWidgets.QPushButton(self.centralwidget)
        self.addCommandTab.setGeometry(QtCore.QRect(460, 500, 131, 31))
        self.addCommandTab.setToolTip("")
        self.addCommandTab.setObjectName("addCommandTab")
        self.copyCommandTab = QtWidgets.QPushButton(self.centralwidget)
        self.copyCommandTab.setGeometry(QtCore.QRect(600, 500, 131, 31))
        self.copyCommandTab.setToolTip("")
        self.copyCommandTab.setObjectName("copyCommandTab")
        self.clearOrder = QtWidgets.QPushButton(self.centralwidget)
        self.clearOrder.setGeometry(QtCore.QRect(314, 510, 111, 23))
        self.clearOrder.setObjectName("clearOrder")
        mainWindow.setCentralWidget(self.centralwidget)
        self.menuBar = QtWidgets.QMenuBar(mainWindow)
        self.menuBar.setGeometry(QtCore.QRect(0, 0, 1210, 21))
        self.menuBar.setObjectName("menuBar")
        self.menuFile = QtWidgets.QMenu(self.menuBar)
        self.menuFile.setObjectName("menuFile")
        mainWindow.setMenuBar(self.menuBar)
        self.actionSave = QtWidgets.QAction(mainWindow)
        self.actionSave.setObjectName("actionSave")
        self.menuFile.addAction(self.actionSave)
        self.menuBar.addAction(self.menuFile.menuAction())

        self.retranslateUi(mainWindow)
        self.showExcelData.setCurrentIndex(-1)
        QtCore.QMetaObject.connectSlotsByName(mainWindow)

    def retranslateUi(self, mainWindow):
        _translate = QtCore.QCoreApplication.translate
        mainWindow.setWindowTitle(_translate("mainWindow", "auto"))
        self.FileLoad.setText(_translate("mainWindow", "파일 선택"))
        self.makeOrder.setText(_translate("mainWindow", "주문하기"))
        self.FileDirectory.setToolTip(_translate("mainWindow", "<html><head/><body><p>파일을 선택하면 파일의 경로가 표시됩니다.</p></body></html>"))
        self.FileDirectory.setText(_translate("mainWindow", "파일 경로"))
        self.addCommandTab.setText(_translate("mainWindow", "탭 추가"))
        self.copyCommandTab.setText(_translate("mainWindow", "현재 탭 복사"))
        self.clearOrder.setText(_translate("mainWindow", "주문 데이터 초기화"))
        self.menuFile.setTitle(_translate("mainWindow", "File"))
        self.actionSave.setText(_translate("mainWindow", "Save"))
        self.actionSave.setShortcut(_translate("mainWindow", "Ctrl+S"))





    def addExcelTab(self, stock : Stocks): # stocks 데이터로 tab 추가
        # QTableWidget 생성
        tempTable = QTableWidget()
        tempTable.setRowCount(stock.order_len())
        tempTable.setColumnCount(4)
        tempTable.setHorizontalHeaderLabels(["매수/매도", "주문가", "주문 방법", "수량"])

        # 데이터 추가
        for row, order in enumerate(stock.orders):
            tempTable.setItem(row, 0, QTableWidgetItem('매수' if order.bs else '매도'))
            tempTable.setItem(row, 1, QTableWidgetItem(str(order.price)))
            tempTable.setItem(row, 2, QTableWidgetItem(order.method))
            tempTable.setItem(row, 3, QTableWidgetItem(str(order.quantity)))
        
        # QTabWidget에 테이블 추가
        self.showExcelData.addTab(tempTable, stock.stock)

        logger.info(f"tab {stock.stock} added")
    
    def ExcelTabLoad(self):
        data = self.excel.orderData

        for stock in data:
            self.addExcelTab(stock)

    def ExcelTabClear(self):
        self.showExcelData.clear()


    def btn_FileLoad(self): # 파일 불러오기
        fname = QFileDialog.getOpenFileName(self, '파일 불러오기', '', 'Excel Files(*.xlsx *.xls);; All Files(*.*)')
        
        if fname[0]: #파일 선택
            logger.info(f"selected file : {fname[0]}")

            self.FileDirectory.setText(f"{fname[0]}")
            self.FileDirectory.setToolTip(fname[0])

            self.excel.getData(fname[0])

            logger.info("successfuly loaded data")

            # self.ExcelTabClear()
            self.ExcelTabLoad()

        else :
            QMessageBox.about(self, "Error", "파일을 선택해주세요")

    def btn_clearOrder(self):
        self.showExcelData.clear()
        self.FileDirectory.setText("파일 경로")
        self.FileDirectory.setToolTip("파일 경로가 표시됩니다.")


    def btn_addCommandTab(self, default_name=None):
        # 새로운 탭 생성
        new_tab = QWidget()
        tab_layout = QVBoxLayout(new_tab)

        # ScrollArea 생성
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        # ScrollArea 내부 콘텐츠와 레이아웃
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setSpacing(10)  # 위젯 간 간격 고정


        scroll_content.setLayout(scroll_layout)

        # ScrollArea에 위젯 추가
        scroll_area.setWidget(scroll_content)

        # 버튼 레이아웃
        button_layout = QHBoxLayout()

        # ScrollArea 내부에 위젯 추가 버튼
        add_widget_button = QPushButton("커맨드 추가")
        add_widget_button.setMaximumWidth(120)
        add_widget_button.clicked.connect(lambda: self.addCommandWidget(scroll_layout))
        button_layout.addWidget(add_widget_button)

        # 탭 이름 수정 버튼
        edit_tab_name_button = QPushButton("탭 이름 수정")
        edit_tab_name_button.setMaximumWidth(120)
        edit_tab_name_button.clicked.connect(self.btn_editCommandTabName)
        button_layout.addWidget(edit_tab_name_button)

        # 탭 삭제 버튼
        delete_tab_button = QPushButton("탭 삭제")
        delete_tab_button.setMaximumWidth(120)
        delete_tab_button.clicked.connect(self.btn_deleteCommandTab)
        button_layout.addWidget(delete_tab_button)

        # 레이아웃 구성
        tab_layout.addWidget(scroll_area)
        tab_layout.addLayout(button_layout)

        # 탭 이름 설정
        tab_name = default_name if default_name else f"탭 {self.showCommand.count() + 1}"

        # TabWidget에 새로운 탭 추가
        self.showCommand.addTab(new_tab, tab_name)

        logger.debug(f'tab {tab_name} added')

    def addCommandWidget(self, layout):
        widget = QWidget()
        widget_layout = QHBoxLayout(widget)
        widget.setFixedHeight(50)
    
        # 데이터 수정용 QComboBox
        combo_box = QComboBox()
        combo_box.addItems(["커서 이동", "마우스 클릭", "키보드 입력"])
        widget_layout.addWidget(combo_box) # 0
    
        # 커서 이동에서 사용하는 입력 필드
        int_input_1 = QLineEdit()
        int_input_1.setPlaceholderText("x")
        int_input_1.setMaximumWidth(50)
        int_input_1.setVisible(True)
        widget_layout.addWidget(int_input_1) # 1
    
        int_input_2 = QLineEdit()
        int_input_2.setPlaceholderText("y")
        int_input_2.setMaximumWidth(50)
        int_input_2.setVisible(True)
        widget_layout.addWidget(int_input_2) # 2
    
        # 커서 이동에서 사용하는 버튼
        get_mouse_btn = QPushButton("마우스 위치 가져오기")
        get_mouse_btn.setMaximumWidth(150)
        get_mouse_btn.setVisible(True)
        widget_layout.addWidget(get_mouse_btn) # 3
    
        move_mouse_btn = QPushButton("위치로 커서 이동")
        move_mouse_btn.setMaximumWidth(150)
        move_mouse_btn.setVisible(True)
        widget_layout.addWidget(move_mouse_btn) # 4
    
        # C에서 사용하는 드롭다운
        c_dropdown_1 = QComboBox()
        c_dropdown_1.addItems(["테이블에서 가져오기", "직접 입력"])
        c_dropdown_1.setVisible(False)
        widget_layout.addWidget(c_dropdown_1) # 5
    
        c_dropdown_2 = QComboBox()
        c_dropdown_2.addItems(["종목", "주문가", "주문 방법", "수량"])
        c_dropdown_2.setVisible(False)
        widget_layout.addWidget(c_dropdown_2) # 6
    
        c_text_input = QLineEdit()
        c_text_input.setPlaceholderText("텍스트 입력")
        c_text_input.setVisible(False)
        widget_layout.addWidget(c_text_input) # 7
    
        # 삭제 버튼
        delete_button = QPushButton("커맨드 삭제")
        delete_button.setMaximumWidth(120)
        widget_layout.addWidget(delete_button)
    
        # 마우스 위치 가져오기 버튼 동작
        def get_mouse_position():
            if get_mouse_btn.text() == "마우스 위치 가져오기":
                timer = QTimer()
                remaining_time = [3]  # 리스트로 선언해 mutable 상태 유지
                get_mouse_btn.setText(f"{remaining_time[0]}초 후 가져오기...")

                def countdown():
                    remaining_time[0] -= 1
                    
                    if remaining_time[0] > 0:
                        get_mouse_btn.setText(f"{remaining_time[0]}초 후 가져오기...")
                    else:
                        timer.stop()
                        timer.deleteLater()
                        x, y = pyautogui.position()
                        int_input_1.setText(str(x))
                        int_input_2.setText(str(y))
                        get_mouse_btn.setText("마우스 위치 가져오기")

                timer.timeout.connect(countdown)
                timer.start(1000)
    
        # 마우스 이동 버튼 동작
        def move_to_position():
            try:
                x = int(int_input_1.text())
                y = int(int_input_2.text())
                pyautogui.moveTo(x, y)
            except Exception as e:
                move_mouse_btn.setText("유효하지 않은 위치!")
                QTimer.singleShot(2000, lambda: move_mouse_btn.setText("위치로 커서 이동"))
                logger.error(e)
    
        # C 드롭다운 1 업데이트 함수
        def update_c_dropdown(value):
            c_dropdown_2.setVisible(False)
            c_text_input.setVisible(False)
            if value == "테이블에서 가져오기":
                c_dropdown_2.setVisible(True)
            elif value == "직접 입력":
                c_text_input.setVisible(True)
    
        # 드롭다운 변경 시 동작
        combo_box.currentTextChanged.connect(lambda: update_widget(combo_box.currentText()))
        c_dropdown_1.currentTextChanged.connect(update_c_dropdown)
    
        # 위젯 상태 업데이트
        def update_widget(value):
            # 모든 입력 필드 숨김
            int_input_1.setVisible(False)
            int_input_2.setVisible(False)
            get_mouse_btn.setVisible(False)
            move_mouse_btn.setVisible(False)
            c_dropdown_1.setVisible(False)
            c_dropdown_2.setVisible(False)
            c_text_input.setVisible(False)
    
            if value == "커서 이동":
                int_input_1.setVisible(True)
                int_input_2.setVisible(True)
                get_mouse_btn.setVisible(True)
                move_mouse_btn.setVisible(True)
            elif value == "키보드 입력":
                c_dropdown_1.setVisible(True)
                update_c_dropdown(c_dropdown_1.currentText())  # 초기 상태 처리
    
        # 삭제 버튼 동작
        delete_button.clicked.connect(lambda: self.deleteCommandWidget(widget, layout))
    
        # 버튼 동작 연결
        get_mouse_btn.clicked.connect(get_mouse_position)
        move_mouse_btn.clicked.connect(move_to_position)
    
        # 초기 상태 설정
        update_widget("커서 이동")
        layout.addWidget(widget)


    def deleteCommandWidget(self, widget, layout):
        # ScrollArea에서 해당 위젯 삭제
        layout.removeWidget(widget)
        widget.deleteLater()

    def btn_editCommandTabName(self):
        # 현재 탭의 이름을 수정
        current_index = self.showCommand.currentIndex()
        if current_index != -1:
            # 현재 탭의 이름 가져오기
            current_tab_name = self.showCommand.tabText(current_index)

            # QInputDialog를 사용해 새 이름 입력 받기
            new_name, ok = QInputDialog.getText(
                self, "탭 이름 수정", "새로운 탭 이름 입력:", text=current_tab_name
            )

            # 입력이 확인되면 탭 이름 변경
            if ok and new_name.strip():
                self.showCommand.setTabText(current_index, new_name.strip())

    def btn_deleteCommandTab(self):
        current_index = self.showCommand.currentIndex()
        if current_index != -1:
            logger.debug(f'removed tab {self.showCommand.tabText(current_index)}')
            self.showCommand.removeTab(current_index)


    def btn_copyCommandTab(self):
        # 현재 선택된 탭의 위젯을 복사하여 새로운 탭으로 추가
        current_widget = self.showCommand.currentWidget()
        
        if current_widget:
            # 탭의 데이터를 가져옴
            current_tab_index = self.showCommand.indexOf(current_widget)
            tab_title = self.showCommand.tabText(current_tab_index)
            
            current_command_tab = current_widget.layout().itemAt(0).widget()
            current_command = current_command_tab.widget()

        # 탭 위젯 생성

            # 새로운 탭 생성
            new_tab = QWidget()
            tab_layout = QVBoxLayout(new_tab)

            # ScrollArea 생성
            scroll_area = QScrollArea()
            scroll_area.setWidgetResizable(True)

            # ScrollArea 내부 콘텐츠와 레이아웃
            scroll_content = QWidget()
            scroll_layout = QVBoxLayout(scroll_content)
            scroll_layout.setSpacing(10)  # 위젯 간 간격 고정


            scroll_content.setLayout(scroll_layout)

            # ScrollArea에 위젯 추가
            scroll_area.setWidget(scroll_content)

            # 버튼 레이아웃
            button_layout = QHBoxLayout()

            # ScrollArea 내부에 위젯 추가 버튼
            add_widget_button = QPushButton("커맨드 추가")
            add_widget_button.setMaximumWidth(120)
            add_widget_button.clicked.connect(lambda: self.addCommandWidget(scroll_layout))
            button_layout.addWidget(add_widget_button)

            # 탭 이름 수정 버튼
            edit_tab_name_button = QPushButton("탭 이름 수정")
            edit_tab_name_button.setMaximumWidth(120)
            edit_tab_name_button.clicked.connect(self.btn_editCommandTabName)
            button_layout.addWidget(edit_tab_name_button)

            # 탭 삭제 버튼
            delete_tab_button = QPushButton("탭 삭제")
            delete_tab_button.setMaximumWidth(120)
            delete_tab_button.clicked.connect(self.btn_deleteCommandTab)
            button_layout.addWidget(delete_tab_button)

            # 레이아웃 구성
            tab_layout.addWidget(scroll_area)
            tab_layout.addLayout(button_layout)

        # 위젯들 생성
            new_command_tab = new_tab.layout().itemAt(0).widget()
            new_command = new_command_tab.widget()

            for command_text in current_command.children(): # 기존 커맨드
                if isinstance(command_text, QWidget):
                    self.addCommandWidget(scroll_layout)

                    command_copy = new_command.children()[-1] # 새로운 커맨드


                    for i in range(8):
                        if isinstance(command_copy.layout().itemAt(i).widget(), QComboBox): #QComboBox
                            command_copy.layout().itemAt(i).widget().setCurrentIndex(command_text.layout().itemAt(i).widget().currentIndex())
                        elif isinstance(command_copy.layout().itemAt(i).widget(), QLineEdit): #QLineEdit
                            command_copy.layout().itemAt(i).widget().setText(command_text.layout().itemAt(i).widget().text())
                        elif isinstance(command_copy.layout().itemAt(i).widget(), QPushButton): # QPushbutton
                            pass

            
        # 새로운 탭을 추가
            new_tab_index = self.showCommand.addTab(new_tab, f"{tab_title}의 복사본")
            self.showCommand.setCurrentIndex(new_tab_index)  # 새로 추가된 탭으로 이동

            logger.debug(f'copied tab {tab_title}')



    def btn_makeOrder(self): # 현재 탭에서, 모든 주문 행에 대하여 커맨드를 수행
        if self.order_thread is None or not self.order_thread.isRunning():
            logger.info('order started')

            self.order_thread = OrderThread(self.showExcelData, self.showCommand)

            self.order_thread.highlightExcelTable_signal.connect(self.ExcelTabHighlight)
            self.order_thread.highlightCommandWidget_signal.connect(self.CommandWidgetHighlight)

            self.order_thread.progress_signal.connect(self.update_status)  # 진행 상태를 UI에 업데이트
            self.order_thread.finished_signal.connect(self.on_finished) # 완료 시 삭제
            self.order_thread.start()
        elif self.order_thread and self.order_thread.isPaused():
            logger.info("Resuming order thread.")
            self.order_thread.resume()
            
            self.makeOrder.setText('재개하기')
    
    def update_status(self, status): # 주문 진행 상태 업데이트
        QMessageBox.about(self, "상태", status)

    def on_finished(self): # 주문 완료 시
        # QMessageBox.about(self, "알림", "주문 완료됨")
        self.order_thread.deleteLater()


    def ExcelTabHighlight(self, row, color):
        table = self.showExcelData.currentWidget()

        for col in range(table.columnCount()):
            table.item(row, col).setBackground(color)  # 강조

    def CommandWidgetHighlight(self, widget, color):
        widget.setStyleSheet(color)


    def on_press(self, key): # 키보드 입력 시
        try:
            if key == keyboard.Key.f4:
                if self.order_thread and self.order_thread.isRunning():
                    logger.info("Stopping the order thread.")
                    self.order_thread.stop()  # OrderThread 멈추기
            if key == keyboard.Key.f3:
                if self.order_thread:
                    if self.order_thread.isPaused():
                        logger.info("Resuming order thread.")
                        self.order_thread.resume()
                        
                        self.makeOrder.setText('주문하기')
                    else:
                        logger.info("Pausing order thread.")
                        self.order_thread.pause()

                        self.makeOrder.setText('재개하기')
        except AttributeError:
            pass


    def saveCommand(self):
        commandTabs = []

        for idx in range(self.showCommand.count()): # 모든 탭에 대하여
            Tab = Commands(self.showCommand.tabText(idx))
        
            logger.debug(f'saving tab {self.showCommand.tabText(idx)}')

            command_tab = self.showCommand.widget(idx).layout().itemAt(0).widget()
            commands = command_tab.widget()

            for command in commands.children():
                if isinstance(command, QWidget):
                    com = []

                    for i in range(8):
                        if isinstance(command.layout().itemAt(i).widget(), QComboBox): #QComboBox
                            com.append(command.layout().itemAt(i).widget().currentIndex())
                        elif isinstance(command.layout().itemAt(i).widget(), QLineEdit): #QLineEdit
                            com.append(command.layout().itemAt(i).widget().text())
                        elif isinstance(command.layout().itemAt(i).widget(), QPushButton): # QPushbutton
                            pass
                        
                    Tab.add_command(com)
            
            commandTabs.append(Tab)
            logger.debug(f'Tab {Tab.name} appended')

        with open('command.pickle', 'wb') as f:
            pickle.dump(commandTabs, f)


    def openFile(self):
        try:
            with open('command.pickle', 'rb') as f:
                return pickle.load(f)
        except FileNotFoundError:
            return []

    def loadCommand(self):
        commandTabs = self.openFile()

        if commandTabs == []:
            return
        
        for tabData in commandTabs:
        # 새로운 탭 생성
            new_tab = QWidget()
            tab_layout = QVBoxLayout(new_tab)

            # ScrollArea 생성
            scroll_area = QScrollArea()
            scroll_area.setWidgetResizable(True)

            # ScrollArea 내부 콘텐츠와 레이아웃
            scroll_content = QWidget()
            scroll_layout = QVBoxLayout(scroll_content)
            scroll_layout.setSpacing(10)  # 위젯 간 간격 고정


            scroll_content.setLayout(scroll_layout)

            # ScrollArea에 위젯 추가
            scroll_area.setWidget(scroll_content)

            # 버튼 레이아웃
            button_layout = QHBoxLayout()

            # ScrollArea 내부에 위젯 추가 버튼
            add_widget_button = QPushButton("커맨드 추가")
            add_widget_button.setMaximumWidth(120)
            add_widget_button.clicked.connect(lambda: self.addCommandWidget(scroll_layout))
            button_layout.addWidget(add_widget_button)

            # 탭 이름 수정 버튼
            edit_tab_name_button = QPushButton("탭 이름 수정")
            edit_tab_name_button.setMaximumWidth(120)
            edit_tab_name_button.clicked.connect(self.btn_editCommandTabName)
            button_layout.addWidget(edit_tab_name_button)

            # 탭 삭제 버튼
            delete_tab_button = QPushButton("탭 삭제")
            delete_tab_button.setMaximumWidth(120)
            delete_tab_button.clicked.connect(self.btn_deleteCommandTab)
            button_layout.addWidget(delete_tab_button)

            # 레이아웃 구성
            tab_layout.addWidget(scroll_area)
            tab_layout.addLayout(button_layout)

        # 위젯들 생성
            new_command_tab = new_tab.layout().itemAt(0).widget()
            new_command = new_command_tab.widget()

            for command_text in tabData.command: # 기존 커맨드
                self.addCommandWidget(scroll_layout)

                command_copy = new_command.children()[-1] # 새로운 커맨드

                i = 0

                for idx in range(8):
                    if isinstance(command_copy.layout().itemAt(idx).widget(), QComboBox): #QComboBox
                        command_copy.layout().itemAt(idx).widget().setCurrentIndex(command_text[i])
                        i += 1
                    elif isinstance(command_copy.layout().itemAt(idx).widget(), QLineEdit): #QLineEdit
                        command_copy.layout().itemAt(idx).widget().setText(command_text[i])
                        i += 1
                    elif isinstance(command_copy.layout().itemAt(idx).widget(), QPushButton): # QPushbutton
                        pass

            
        # 새로운 탭을 추가
            self.showCommand.addTab(new_tab, tabData.name)



    def closeEvent(self, event):
        reply = QMessageBox.question(self, '커맨드 저장?', '커맨드를 저장하시겠습니까?',QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel, QMessageBox.Cancel)
        
        if reply == QMessageBox.Yes:
            logger.info("saving command")  # 저장 작업을 수행하는 부분
            logging.shutdown()

            self.saveCommand()

            event.accept()  # 창을 닫는다.
        elif reply == QMessageBox.No:
            logger.info("shut down")  # 저장하지 않고 닫기
            logging.shutdown()

            event.accept()  # 창을 닫는다.
        else:
            logger.info("cancel close")
            event.ignore()  # 창 닫기를 취소한다.


if __name__ == "__main__":
    sys.excepthook = exception_handler

    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    sys.exit(app.exec_())