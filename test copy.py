import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import re
from linebot.models import *
from linebot import *
import qrcode
import time

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

cells = order_sheet.find('20220512220301')
item = order_sheet.row_values(cells.row)
print(item)