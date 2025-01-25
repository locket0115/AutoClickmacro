# 엑셀을 읽어오고 데이터를 처리하는 클래스
import logging

import pandas as pd
from data import *


# 루트 로거 생성
logger = logging.getLogger('root.excel')
logger.setLevel(logging.DEBUG)  # 모든 로그 메시지를 처리하기 위해 DEBUG 레벨로 설정

# 로그 파일 (logger.log) 핸들러 설정
file_handler = logging.FileHandler('logger.log')
file_handler.setLevel(logging.DEBUG)  # 모든 로그를 기록

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



class Excel:
    def __init__(self):
        self.file_path = None
        self.excelData = None
        self.orderData = None
    
    def readFile(self): # 데이터 읽어오기
        try:
            self.excelData = pd.read_excel(self.file_path, sheet_name=None, na_values = 0, engine='openpyxl')
            logger.info(f"엑셀 파일을 성공적으로 읽었습니다.")
        except Exception as e:
            logger.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")

    def transData(self): # 데이터 처리
        self.orderData = [] #반환값

        for sheet in self.excelData:
            name = self.excelData[sheet].columns[0]

            print(name)

            if(sheet == 'Data'): # Data 시트 제외
                continue

            stock = Stocks(name, [])

            for row in self.excelData[sheet].iterrows(): #모든 행에 대해 반복
                bs = row[1].iloc[0] # 매수/매도
                price = row[1]['주문가'] # 주문가
                method = row[1]['유형'] # 주문 방법
                quantity = row[1]['수량'] # 수량

                if(quantity == 0): # 수량이 0인 경우 제외
                    continue

                stock.add_order(Order(True if bs == '매수' else False, price, method, quantity))

            if stock.orders != []:
                self.orderData.append((stock, sheet))
                logger.debug(f'sheet {sheet} appended')
    
    def getData(self, file_path):
        self.file_path = file_path
        
        self.readFile()
        self.transData()

        return self.orderData
