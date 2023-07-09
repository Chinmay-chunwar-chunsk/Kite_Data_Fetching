from kite_trade import *
import xlwings as xl
import time

def start_excel():
    data_book = xl.Book("Data.xlsx")
    sheets = data_book.sheets
    data_sheet=sheets[0]
    while True:
        time.sleep(1)
        if data_sheet.range("L2").value == None:
            pass
        else:
            enctoken = data_sheet.range("L2").value
            kite = KiteApp(enctoken=enctoken)
            counter=2
            for i in range(20):
                if data_sheet.range(f"B{counter}").value == None:
                    pass
                else:
                    try:
                        symbol = data_sheet.range(f"B{counter}").value
                        data=kite.quote(symbol)
                        ltp = kite.ltp(symbol)
                        ltp = ltp[symbol]
                        data=data[symbol]
                        ohlc=data["ohlc"]
                        open=ohlc["open"]
                        high=ohlc["high"]
                        low=ohlc["low"]
                        close=ohlc["close"]
                        last_price = ltp["last_price"]
                        # print(f"open: {open}, high: {high}, low: {low}, close: {close}")
                        data_sheet.range(f"C{counter}").value = open
                        data_sheet.range(f"D{counter}").value = high
                        data_sheet.range(f"E{counter}").value = low
                        data_sheet.range(f"F{counter}").value = close
                        data_sheet.range(f"G{counter}").value = last_price
                    except KeyError or TypeError:
                        pass

                counter+=1


start_excel()
    

# tkn=str(input("Enter enctoken: "))
# Login(tkn)
