from kite_trade import *
import pandas as pd    
import xlwings as xl 
import time

data_book = xl.Book("test2.xlsx")
sheets = data_book.sheets
data_sheet = sheets[0]
def start_excel():
    global kite
    while True: 
        start = time.perf_counter()
        data = []
        time.sleep(1)
        symbols = data_sheet.range("b2:b21").value
        try:
            data = pd.DataFrame(kite.ltp(symbols))
            data = data.iloc[1]
            data_sheet.range("b2:c21").value = data
        except KeyError:
            pass
        end = time.perf_counter()
        print(f"Time taken: {end - start}")


def login():
    global kite
    while True:
        try:
            if data_sheet.range("e2").value == "":
                pass
            else:
                enctoken = data_sheet.range("e2").value
                kite = KiteApp(enctoken=enctoken)
                start_excel()
        except IndexError:
            pass
login()