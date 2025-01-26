# 엑셀을 읽어오고 데이터를 처리하는 클래스
import logging

import pandas as pd
from data import *
from log import logger



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
