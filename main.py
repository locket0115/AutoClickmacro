import os
import sys
import pyautogui
import logging
import pickle

from PyQt5.QtWidgets import QWidget, QMainWindow, QTabWidget, QTableWidget, QTableWidgetItem, QFileDialog, QMessageBox, QVBoxLayout, QScrollArea, QHBoxLayout, QPushButton, QComboBox, QLineEdit, QInputDialog, QApplication
from PyQt5.QtCore import QThread, pyqtSignal, QMutex, QWaitCondition, QTimer
from PyQt5.QtGui import QColor
from PyQt5 import QtCore, QtWidgets
from pynput import keyboard


from excel import *
from data import *
from log import logger

# const
command_widget_cnt = 11  # command_widget 내부 위젯 갯수


# 디렉토리
cur_dir = os.getcwd()
cmd_dir = os.path.join(cur_dir, 'coms')

# 디렉토리 생성
if not os.path.exists(cmd_dir):
    os.makedirs(cmd_dir)


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
        self.waiting = False

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

        stock = current_table.horizontalHeaderItem(0).text()
        logger.debug(f"Stock: {stock}")



        self.executeCommand("계좌선택", stock)
    
        self.mutex.lock()
        if self.paused: # 일시정지 체크
            self.progress_signal.emit('주문 일시정지됨.')
            self.waiting = True
            self.condition.wait(self.mutex)

            self.waiting = False
        
        if not self.running: # 정지 체크
            self.progress_signal.emit("주문 정지됨.")
            return  
        self.mutex.unlock()

        self.executeCommand("종목선택", stock)



        for row in range(current_table.rowCount()): # 주문 row에 대하여
            self.highlightExcelTable_signal.emit(row, QColor(128, 128, 128)) # 강조

            self.mutex.lock()
            if self.paused: # 일시정지 체크
                self.progress_signal.emit('주문 일시정지됨.')
                self.waiting = True
                self.condition.wait(self.mutex)

                self.waiting = False

            if not self.running: # 정지 체크
                self.progress_signal.emit("주문 정지됨.")
                self.highlightExcelTable_signal.emit(row, QColor(255, 255, 255)) # 강조 해제
                return
            
            self.mutex.unlock()


            bs = current_table.item(row, 0).text()
            price = current_table.item(row, 1).text()
            method = current_table.item(row, 2).text()
            quantity = current_table.item(row, 3).text()


            logger.debug(f"Order: {bs}, {price}, {method}, {quantity}")


            self.executeCommand(f'{bs}_{method}', stock, method, price, quantity) # 수도_방법 실행

            self.mutex.lock()
            if self.paused: # 일시정지 체크
                self.progress_signal.emit('주문 일시정지됨.')
                self.waiting = True
                self.condition.wait(self.mutex)

                self.waiting = False

            if not self.running: # 정지 체크
                self.progress_signal.emit("주문 정지됨.")
                self.highlightExcelTable_signal.emit(row, QColor(255, 255, 255)) # 강조 해제
                return

            self.executeCommand("매매", stock, method, price, quantity) # 매매 실행



            self.highlightExcelTable_signal.emit(row, QColor(255, 255, 255)) # 강조 해제
            logger.info(f'command tab completed')

            
        logger.info(f'order completed')
        self.progress_signal.emit("주문 완료")
        self.finished_signal.emit()


    def executeCommand(self, tab_name, stock, method = None, price = None, quantity = None):
        command_tab = None
    
        for idx in range(self.showCommand.count()):
            if self.showCommand.tabText(idx) == tab_name:
                command_tab = self.showCommand.widget(idx).layout().itemAt(0).widget()
                self.showCommand.setCurrentIndex(idx)

                logger.debug(f"tab {idx}, {self.showCommand.tabText(idx)} used.")

                break
        

        if command_tab == None:
            logger.debug("No such Tab")
        else:
            self.executeTab(command_tab, stock, method, price, quantity)
        

    def executeTab(self, command_tab, stock, method, price, quantity):
        current_command = command_tab.widget()


        for command in current_command.children(): # 커맨드 탭에 대하여
            self.mutex.lock()

            if self.paused: # 일시정지 체크
                self.progress_signal.emit('주문 일시정지됨.')
                self.waiting = True
                self.condition.wait(self.mutex)

                self.waiting = False
            
            if not self.running: # 정지 체크
                self.progress_signal.emit("주문 정지됨.")
                return
            
            self.mutex.unlock()
            
            if isinstance(command, QWidget): # 위젯 실행
                self.highlightCommandWidget_signal.emit(command, "QWidget { background-color: white; }") # 위젯 강조

                command_tab.ensureWidgetVisible(command) # 위젯이 보이도록 스크롤

                if command.layout().itemAt(0).widget().currentText() == "커서 이동":
                    try:
                        x = int(command.layout().itemAt(1).widget().text())
                        y = int(command.layout().itemAt(2).widget().text())
                        duration = float(command.layout().itemAt(8).widget().text())
                        
                        logger.debug(f"Move to: {x}, {y}, dur: {duration}")
                        pyautogui.moveTo(x, y, duration=duration)
                    except Exception as e:
                        logger.warning(e)
                elif command.layout().itemAt(0).widget().currentText() == "마우스 클릭":
                    logger.debug("Click")
                    pyautogui.click()
                elif command.layout().itemAt(0).widget().currentText() == "키보드 입력":
                    if command.layout().itemAt(5).widget().currentText() == "테이블에서 가져오기":
                        logger.debug(f"Type: {command.layout().itemAt(6).widget().currentText()}")
                        try:
                            idx = command.layout().itemAt(6).widget().currentText()
                            interval = interval=float(command.layout().itemAt(8).widget().text())
                            if idx == '종목':
                                pyautogui.typewrite(stock, interval=interval)
                            elif idx == '주문가':
                                pyautogui.typewrite(price, interval=interval)
                            elif idx == '주문 방법':
                                pyautogui.typewrite(method, interval=interval)
                            elif idx == '수량':
                                pyautogui.typewrite(quantity, interval=interval)
                        except Exception as e:
                            logger.warning(e)
                    else:
                        logger.debug(f"""Type: '{command.layout().itemAt(7).widget().text()}'""")
                        pyautogui.typewrite(command.layout().itemAt(7).widget().text())
                else:
                    logger.warning("Invalid Command")
                self.highlightCommandWidget_signal.emit(command, "QWidget { background-color: none; }")
                logger.info(f'command completed')
                
        logger.info(f'command tab completed')

    def stop(self):
        self.running = False
        # self.finished_signal.emit()  # 스레드 종료 신호

        if self.paused == True: # 재개 후 종료
            self.mutex.lock()
            self.paused = False
            self.condition.wakeAll()  # 스레드를 재개
            self.mutex.unlock()

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

    def isWaiting(self):
        return self.waiting

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

        self.actionSave.triggered.connect(self.saveCommand_)
        self.actionSave_as.triggered.connect(self.saveCommandFile)
        self.actionLoad.triggered.connect(self.loadCommandFile)

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
        mainWindow.setCentralWidget(self.centralwidget)
        self.menuBar = QtWidgets.QMenuBar(mainWindow)
        self.menuBar.setGeometry(QtCore.QRect(0, 0, 1210, 21))
        self.menuBar.setObjectName("menuBar")
        self.menuFile = QtWidgets.QMenu(self.menuBar)
        self.menuFile.setObjectName("menuFile")
        mainWindow.setMenuBar(self.menuBar)
        self.actionSave = QtWidgets.QAction(mainWindow)
        self.actionSave.setObjectName("actionSave")
        self.actionSave_as = QtWidgets.QAction(mainWindow)
        self.actionSave_as.setObjectName("actionSave_as")
        self.actionLoad = QtWidgets.QAction(mainWindow)
        self.actionLoad.setObjectName("actionLoad")
        self.menuFile.addAction(self.actionLoad)
        self.menuFile.addSeparator()
        self.menuFile.addAction(self.actionSave)
        self.menuFile.addAction(self.actionSave_as)
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
        self.menuFile.setTitle(_translate("mainWindow", "File"))
        self.actionSave.setText(_translate("mainWindow", "Save"))
        self.actionSave.setShortcut(_translate("mainWindow", "Ctrl+S"))
        self.actionSave_as.setText(_translate("mainWindow", "Save as"))
        self.actionLoad.setText(_translate("mainWindow", "Load"))


    def addExcelTab(self, stock : Stocks, name): # stocks 데이터로 tab 추가
        # QTableWidget 생성
        tempTable = QTableWidget()
        tempTable.setRowCount(stock.order_len())
        tempTable.setColumnCount(4)
        tempTable.setHorizontalHeaderLabels([stock.stock, "주문가", "주문 방법", "수량"])

        # 데이터 추가
        for row, order in enumerate(stock.orders):
            tempTable.setItem(row, 0, QTableWidgetItem('매수' if order.bs else '매도'))
            tempTable.setItem(row, 1, QTableWidgetItem(str(order.price)))
            tempTable.setItem(row, 2, QTableWidgetItem(order.method))
            tempTable.setItem(row, 3, QTableWidgetItem(str(order.quantity)))
        
        # QTabWidget에 테이블 추가
        self.showExcelData.addTab(tempTable, name)

        logger.info(f"tab {name} added")
    
    def ExcelTabLoad(self):
        data = self.excel.orderData
        
        for stock, name in data:
            self.addExcelTab(stock, name)


    def btn_FileLoad(self): # 파일 불러오기
        fname = QFileDialog.getOpenFileName(self, '파일 불러오기', '', 'Excel Files(*.xlsx *.xls);; All Files(*.*)')
        
        if fname[0]: #파일 선택
            logger.info(f"selected file : {fname[0]}")

            self.FileDirectory.setText(f"{fname[0]}")
            self.FileDirectory.setToolTip(fname[0])

            self.excel.getData(fname[0])

            logger.info("successfuly loaded data")

            self.showExcelData.clear()
            self.ExcelTabLoad()

        else :
            QMessageBox.about(self, "Error", "파일이 선택되지 않았습니다.")


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
    
        # delay 입력 필드
        delay_input = QLineEdit()
        delay_input.setPlaceholderText("sec")
        delay_input.setText("1")
        delay_input.setToolTip("다음 커맨드까지의 시간을 지정합니다.")
        delay_input.setMaximumWidth(50)
        delay_input.setVisible(True)
        widget_layout.addWidget(delay_input) # 8

        # tooltip 입력 버튼
        widget_tip = QPushButton("...")
        widget_tip.setMaximumWidth(20)
        widget_tip.setToolTip("memo")
        widget_layout.addWidget(widget_tip) # 9


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
            delay_input.setVisible(False)
    
            if value == "커서 이동":
                int_input_1.setVisible(True)
                int_input_2.setVisible(True)
                get_mouse_btn.setVisible(True)
                move_mouse_btn.setVisible(True)
                delay_input.setVisible(True)
            elif value == "키보드 입력":
                c_dropdown_1.setVisible(True)
                delay_input.setVisible(True)
                update_c_dropdown(c_dropdown_1.currentText())  # 초기 상태 처리
    

        # 툴팁 업데이트
        def update_widget_tip():
            current_tooltip = widget_tip.toolTip()

            # QInputDialog를 사용해 새 이름 입력 받기
            new_name, ok = QInputDialog.getText(
                self, "메모 수정", "새로운 메모 입력:", text=current_tooltip
            )

            # 입력이 확인되면 탭 이름 변경
            if ok and new_name.strip():
                widget_tip.setToolTip(new_name.strip())


        # 삭제 버튼 동작
        delete_button.clicked.connect(lambda: self.deleteCommandWidget(widget, layout))
    
        # 버튼 동작 연결
        get_mouse_btn.clicked.connect(get_mouse_position)
        move_mouse_btn.clicked.connect(move_to_position)
        widget_tip.clicked.connect(update_widget_tip)
    
        # 초기 상태 설정
        update_widget("커서 이동")


        layout.addWidget(widget)
        logger.debug(f"command widget added to {layout}")


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
        reply = QMessageBox.question(self, '커맨드 탭 삭제', '현재 커맨드 탭을 삭제하시겠습니까?',QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
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


                    for i in range(command_widget_cnt):
                        if isinstance(command_copy.layout().itemAt(i).widget(), QComboBox): #QComboBox
                            command_copy.layout().itemAt(i).widget().setCurrentIndex(command_text.layout().itemAt(i).widget().currentIndex())
                        elif isinstance(command_copy.layout().itemAt(i).widget(), QLineEdit): #QLineEdit
                            command_copy.layout().itemAt(i).widget().setText(command_text.layout().itemAt(i).widget().text())
                        elif isinstance(command_copy.layout().itemAt(i).widget(), QPushButton): # QPushbutton
                            command_copy.layout().itemAt(i).widget().setToolTip(command_text.layout().itemAt(i).widget().toolTip())

            
        # 새로운 탭을 추가
            new_tab_index = self.showCommand.addTab(new_tab, f"{tab_title}의 복사본")
            self.showCommand.setCurrentIndex(new_tab_index)  # 새로 추가된 탭으로 이동

            logger.debug(f'copied tab {tab_title}')

    def initCommandTab(self):
        self.btn_addCommandTab("계좌선택")
        self.btn_addCommandTab("종목선택")


        self.btn_addCommandTab("매수_LOC")
        self.btn_addCommandTab("매수_지정가")
        
        self.btn_addCommandTab("매도_LOC")
        self.btn_addCommandTab("매도_지정가")
        
        self.btn_addCommandTab("매매")

        self.saveCommand(os.path.join(cmd_dir, 'default.pickle'))



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
        self.order_thread = None


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
                    
                    self.makeOrder.setText('주문하기')

            if key == keyboard.Key.f3:
                if self.order_thread and self.order_thread.isRunning():
                    if self.order_thread.isWaiting() == True:
                        logger.info("Resuming order thread.")
                        self.order_thread.resume()
                        
                        self.makeOrder.setText('주문하기')
                    elif self.order_thread.isWaiting() == False:
                        logger.info("Pausing order thread.")
                        self.order_thread.pause()

                        self.makeOrder.setText('재개하기')
        except AttributeError:
            pass



    def saveCommand_(self):
        self.saveCommand()

    def saveCommandFile(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getSaveFileName(self, 'Save As', '', 'pickle (*.pickle);; All Files (*)', options=options)

        if fileName:
            self.saveCommand(fileName)

            print(f"File saved as: {fileName}")

    def saveCommand(self, file_path = os.path.join(cmd_dir, 'default.pickle')):
        commandTabs = []

        for idx in range(self.showCommand.count()): # 모든 탭에 대하여
            Tab = Commands(self.showCommand.tabText(idx))
        
            logger.debug(f'saving tab {self.showCommand.tabText(idx)}')

            command_tab = self.showCommand.widget(idx).layout().itemAt(0).widget()
            commands = command_tab.widget()

            for command in commands.children():
                if isinstance(command, QWidget):
                    com = []

                    for i in range(command_widget_cnt):
                        if isinstance(command.layout().itemAt(i).widget(), QComboBox): #QComboBox
                            com.append(command.layout().itemAt(i).widget().currentIndex())
                        elif isinstance(command.layout().itemAt(i).widget(), QLineEdit): #QLineEdit
                            com.append(command.layout().itemAt(i).widget().text())
                        elif isinstance(command.layout().itemAt(i).widget(), QPushButton): # QPushbutton
                            com.append(command.layout().itemAt(i).widget().toolTip())
                            
                        
                    Tab.add_command(com)
            
            commandTabs.append(Tab)
            logger.debug(f'Tab {Tab.name} appended')


        with open(file_path, 'wb') as f:
            pickle.dump(commandTabs, f)
        
        logger.info("successfully saved command file")


    def loadCommandFile(self):
        fname = QFileDialog.getOpenFileName(self, '파일 불러오기', '', 'pickle(*.pickle *.p);; All Files(*.*)')
        
        if fname[0]: #파일 선택
            logger.info(f"selected Command file : {fname[0]}")

            self.showCommand.clear()

            self.loadCommand(fname[0])

            logger.info("successfuly loaded data")

        else :
            QMessageBox.about(self, "Error", "파일이 선택되지 않았습니다.")

    def loadCommand(self, file_path = os.path.join(cmd_dir, 'default.pickle')):
        try:
            with open(file_path, 'rb') as f:
                commandTabs = pickle.load(f)
        except FileNotFoundError:
            logger.info('command file Not found')

            self.initCommandTab()

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

            #tq
            add_widget_button.clicked.connect(self.create_lambda(scroll_layout))
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

                for idx in range(command_widget_cnt):
                    if isinstance(command_copy.layout().itemAt(idx).widget(), QComboBox): #QComboBox
                        command_copy.layout().itemAt(idx).widget().setCurrentIndex(command_text[i])
                        i += 1
                    elif isinstance(command_copy.layout().itemAt(idx).widget(), QLineEdit): #QLineEdit
                        command_copy.layout().itemAt(idx).widget().setText(command_text[i])
                        i += 1
                    elif isinstance(command_copy.layout().itemAt(idx).widget(), QPushButton): # QPushbutton
                        command_copy.layout().itemAt(idx).widget().setToolTip(command_text[i])
                        i += 1
                        

            
        # 새로운 탭을 추가
            self.showCommand.addTab(new_tab, tabData.name)
            logger.debug(f'tab {tabData.name} added')


        if self.showCommand.count() == 0:
            logger.info('initializing Command Tab')
            self.initCommandTab()

    def create_lambda(self, layout):
        return lambda: self.addCommandWidget(layout)



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