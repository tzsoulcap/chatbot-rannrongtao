import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import re
from linebot.models import *
from linebot import *
import qrcode
import time
from datetime import datetime

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
cerds = ServiceAccountCredentials.from_json_keyfile_name("cerds.json", scope)
client = gspread.authorize(cerds)
sheet = client.open("RannRongTao").worksheet(
    'Stock')  # เป็นการเปิดไปยังหน้าชีตนั้นๆ
order_sheet = client.open("RannRongTao").worksheet(
    'Order')
data = sheet.get_all_records()  # การรับรายการของระเบียนทั้งหมด

# mes = qrcode.make("65,000")
# mes.save()

# t = time.localtime()
# t = time.strftime("%Y%m%d%H%M%S", t)
# print(t)


# js = json.loads(js)
# product = js["queryResult"]["outputContexts"]

# val = sheet.cell(1,1).value
# val = sheet.col_values(1)
# val = sheet.get('A2:A9')
# print(val)

# cells = order_sheet.find('20220512220301')
# item = order_sheet.row_values(cells.row)
# print(item)
# datestr = "2022-06-02T12:00:00+07:00"
# timestr = "2022-06-02T12:30:00+07:00"

# t = datetime.strptime(timestr, "%Y-%m-%dT%H:%M:%S%z")
# d = datetime.strptime(datestr, "%Y-%m-%dT%H:%M:%S%z")
# times = datetime.strftime(t, "%H:%M")
# date = datetime.strftime(d, "%w")
# dom = int(datetime.strftime(d, "%d"))


# def week(dd):
#     switcher = {
#         "0": "วันอาทิตย์",
#         "1": "วันจันทร์",
#         "2": "วันอังคาร",
#         "3": "วันพุธ",
#         "4": "วันพฤหัสบดี",
#         "5": "วันศุกร์",
#         "6": "วันเสาร์"
#     }
#     return switcher.get(dd)
# date = week(date)
# date_time = f"{date} ที่ {dom} เวลา {times} น."
# print(date_time)

date = datetime.now()
datestr = datetime.strftime(date, "%d-%m-%Y %H:%M")
print(datestr)