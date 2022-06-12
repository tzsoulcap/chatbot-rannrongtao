from cgitb import text
import json
import gspread
from numpy import product
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import re
from linebot.models import *
from linebot import *
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
sheet_service = client.open("RannRongTao").worksheet(
    'ReserveService')
data = sheet.get_all_records()  # การรับรายการของระเบียนทั้งหมด
line_bot_api = LineBotApi(
    "AM82YvNzOu37BSjeLy5LtvUbDZIdwssqEU4kTuTg7aDEUfrE9MqVoLhAqAT4H43Ggk5Bo9qC2mRRypGGhXpr694K+yxLf7IO7eIK5+CWaKLbsqKz2osEOR5QASQ7RPyjL0EOOV+MfsbDKP1fH3B9CwdB04t89/1O/w1cDnyilFU=")

def RecommendProduct(reply_token):
    image_carousel_template_message = TemplateSendMessage(
    alt_text='ImageCarousel template',
    template=ImageCarouselTemplate(
        columns=[
            ImageCarouselColumn(
                image_url='https://static.nike.com/a/images/t_prod_ss/w_960,c_limit,f_auto/i1-035eca0f-e3ff-4d69-ac4f-ec5015ef643f/-air-jordan-1-court-purple.jpg',
                action=MessageAction(
                    label='สั่งซื้อ1',
                    text='สั่งซื้อ1'
                )
            ),
            ImageCarouselColumn(
                image_url='https://static.nike.com/a/images/t_prod_ss/w_640,c_limit,f_auto/hl5u7kku1dlurbgborm6/air-jordan-4-travis-scott-cactus-jack-release-date.jpg',
                action=MessageAction(
                    label='สั่งซื้อ2',
                    text='สั่งซื้อ2'
                )
            )
        ]
    )
)

   
    line_bot_api.reply_message(reply_token, image_carousel_template_message)


def SentShoes(reply_token):
    text_message = TextMessage(text="""ลูกค้าสามารถส่งรองเท้ามาตามที่อยู่ร้านได้เลยค่ะ
        ที่อยู่ร้าน Rann Rong Tao  105/577 หมู่ 6 ตำบล สุรนารี อำเภอ เมือง จังหวัดนครราชสีมา 30000""")
    line_bot_api.reply_message(reply_token, text_message)

def ReserveService(reply_token, datestr, timestr, service, disname):
    t = datetime.strptime(timestr, "%Y-%m-%dT%H:%M:%S%z")
    d = datetime.strptime(datestr, "%Y-%m-%dT%H:%M:%S%z")
    dom = int(datetime.strftime(d, "%d"))
    times = datetime.strftime(t, "%H:%M")
    date = datetime.strftime(d, "%w")

    def week(dd):
      switcher = {
        "0": "วันอาทิตย์",
        "1": "วันจันทร์",
        "2": "วันอังคาร",
        "3": "วันพุธ",
        "4": "วันพฤหัสบดี",
        "5": "วันศุกร์",
        "6": "วันเสาร์"
      }
      return switcher.get(dd)
    date = week(date)
    date_time = f"คุณ {disname} ได้จองบริการ{service}ที่ร้าน {date} ที่ {dom} เวลา {times} น."

    text_message = TextMessage(text=date_time)
    line_bot_api.reply_message(reply_token, text_message)

    date = datetime.strftime(d, "%d/%m/%Y")
    timestamp = datetime.strftime(datetime.now(), "%d-%m-%Y %H:%M")

    detail = [service, date, times, timestamp, disname]
    last = sheet_service.find("Last Record")
    sheet_service.update(f"A{last.row}:E{last.row}", [detail])
    sheet_service.update_cell(last.row+1, last.col, "Last Record")

def Service(reply_token, id, intent):
    if intent == "Sole Shields":

      flex = """
    {
   "type":"bubble",
   "header":{
      "type":"box",
      "layout":"vertical",
      "contents":[
         {
            "type":"box",
            "layout":"horizontal",
            "contents":[
               {
                  "type":"image",
                  "url":"https://cdn.shopify.com/s/files/1/0225/9891/products/AJ1-CLOSE-SP-SQ-2A_650x.jpg",
                  "size":"full",
                  "aspectMode":"cover",
                  "aspectRatio":"150:196",
                  "gravity":"center",
                  "flex":1
               },
               {
                  "type":"box",
                  "layout":"vertical",
                  "contents":[
                     {
                        "type":"image",
                        "url":"https://cdn.shopify.com/s/files/1/0225/9891/products/BIOSHOCK-WEB-2_650x.jpg?v=1633442791",
                        "size":"full",
                        "aspectMode":"cover",
                        "aspectRatio":"150:98",
                        "gravity":"center"
                     },
                     {
                        "type":"image",
                        "url":"https://cdn.shopify.com/s/files/1/0225/9891/products/BIOSHOCK-WEB-1_650x.jpg?v=1633442794",
                        "size":"full",
                        "aspectMode":"cover",
                        "aspectRatio":"150:98",
                        "gravity":"center"
                     }
                  ],
                  "flex":1
               }
            ]
         }
      ],
      "paddingAll":"0px"
      }
    }""" 
      flex = json.loads(flex)
      flex = FlexSendMessage(alt_text='Flex Message alt text', contents=flex)
      line_bot_api.reply_message(reply_token, flex)

    text_message = TextSendMessage(
        text='เลือกช่องทางการบริการ',
        quick_reply=QuickReply(
          items=[
            QuickReplyButton(
              action=MessageAction(label="จองเวลาบริการที่ร้าน", text="จองเวลาบริการที่ร้าน")
            ),
            QuickReplyButton(
              action=MessageAction(label="ส่งรองเท้าไปที่ร้าน", text="ส่งรองเท้าไปที่ร้าน")
            )
          ]
        )
    )
    line_bot_api.push_message(id, text_message)


def showBuyDetail(intent, reply_token, product, size, id):
    rows = sheet.find(product)
    item = sheet.row_values(rows.row)
    price = item[3]
    imglink = item[5]

    flex = """
    {
        "type": "bubble",
        "header": {
            "type": "box",
            "layout": "vertical",
            "contents": [
              {
                "type": "text",
                "text": "สินค้าของคุณคือ",
                "size": "xl",
                "weight": "bold",
                "margin": "none"
              }
            ]
        },
        "hero": {
            "type": "image",
            "url": "%s",
            "size": "full",
            "aspectRatio": "20:13",
            "aspectMode": "cover"
        },
        "body": {
            "type": "box",
            "layout": "vertical",
            "contents": [
                {
                    "type": "text",
                    "text": "%s",
                    "weight": "bold",
                    "size": "xl"
                },
                {
                    "type": "box",
                    "layout": "vertical",
                    "margin": "lg",
                    "spacing": "sm",
                    "contents": [
                        {
                            "type": "box",
                            "layout": "baseline",
                            "spacing": "sm",
                            "contents": [
                                {
                                    "type": "text",
                                    "text": "Size",
                                    "color": "#aaaaaa",
                                    "size": "sm",
                                    "flex": 1
                                },
                                {
                                    "type": "text",
                                    "text": "%s",
                                    "wrap": true,
                                    "color": "#666666",
                                    "size": "sm",
                                    "flex": 5
                                }
                            ]
                        },
                        {
                            "type": "box",
                            "layout": "baseline",
                            "spacing": "sm",
                            "contents": [
                              {
                                "type": "text",
                                "text": "Price",
                                "color": "#aaaaaa",
                                "size": "sm",
                                "flex": 1
                              },
                              {
                                "type": "text",
                                "text": "%s",
                                "wrap": true,
                                "color": "#666666",
                                "size": "sm",
                                "flex": 5
                              }
                            ]
                        }
                    ]
                }
            ]
        }
    }""" % (imglink, product, size, price)
    flex = json.loads(flex)
    flex = FlexSendMessage(alt_text='Flex Message alt text', contents=flex)
    line_bot_api.reply_message(reply_token, flex)

    text_message = TextSendMessage(
        text='ขอทราบชื่อ ที่อยู่ และเบอร์โทรศัพท์ของคุณลูกค้า เพื่อใช้สำหรับการจัดส่งสินค้าด้วยค่ะ')
    line_bot_api.push_message(id, text_message)


def selectSizes(intent, reply_token, product, id):
    cells = sheet.find(product)
    size = sheet.cell(cells.row, 5).value.split(' ')

    sizeItem = [QuickReplyButton(action=MessageAction(
        # label=s, text='สินค้า {} ไซส์ {}'.format(product, s))) for s in size]
        label=s, text='ไซส์ {}'.format(s))) for s in size]

    text_message = TextSendMessage(text='ไซส์ของ {} มีดังนี้ค่ะ'.format(product),
                                   quick_reply=QuickReply(items=sizeItem))

    line_bot_api.reply_message(reply_token, text_message)


def showItem(intent, reply_token, brand, id):
    r = re.compile(r'{}'.format(brand), re.IGNORECASE)
    d = sheet.findall(r, None, 2)

    items = [sheet.row_values(cell.row) for cell in d]
    codes, names, stocks, prices, sizes, imglinks, amounts = [], [], [], [], [], [], []
    for item in items:
        codes.append(item[0])
        names.append(item[1])
        stocks.append(item[2])
        prices.append(item[3])
        sizes.append(item[4].split(' '))
        imglinks.append(item[5])
        amounts.append(item[6])

    text_message = TextSendMessage(
        text='นี่คือสินค้า {} ในร้านเราค่ะ'.format(brand))
    line_bot_api.push_message(id, text_message)

    carousel_col = []
    for i in range(len(prices)):
        if i == 3:
            break
        carousel_col.append(CarouselColumn(
            thumbnail_image_url=imglinks[i],
            title=names[i],
            text='{}\nSize: {}\nขายแล้ว {}'.format(prices[i], '/'.join(sizes[i]), amounts[i]),
            actions=[
                MessageAction(
                    label='เลือกไซส์',
                    text='เลือกไซส์ {}'.format(names[i])
                )
            ]
        ))

    carousel_template_message = TemplateSendMessage(
        alt_text='Carousel template',
        template=CarouselTemplate(
            columns=carousel_col
        )
    )
    
    line_bot_api.reply_message(reply_token, carousel_template_message)
    # line_bot_api.push_message(uid, carousel_template_message)


def ShippingAddress(reply_token, product, size, customer, address, zip_code, phone_no):
    # line_bot_api.reply_message(
    #     reply_token, TextSendMessage(text='การซื้อสำเร็จ'))

    rows = sheet.find(product)
    item = sheet.row_values(rows.row)
    price = item[3]
    imglink = item[5]

    # t = time.localtime()
    orderID = time.strftime("%Y%m%d%H%M%S", time.localtime())

    flex = """{
  "type": "bubble",
  "header": {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "text",
        "text": "ORDER COMPLETE",
        "size": "sm",
        "weight": "bold",
        "color": "#14B800"
      }
    ]
  },
  "hero": {
    "type": "image",
    "url": "%s",
    "size": "full",
    "aspectRatio": "20:13",
    "aspectMode": "cover"
  },
  "body": {
    "type": "box",
    "layout": "vertical",
    "spacing": "md",
    "contents": [
      {
        "type": "box",
        "layout": "baseline",
        "contents": [
          {
            "type": "text",
            "text": "Order ID",
            "color": "#aaaaaa",
            "flex": 2
          },
          {
            "type": "text",
            "text": "#%s",
            "flex": 4,
            "color": "#aaaaaa"
          }
        ]
      },
      {
        "type": "separator"
      },
      {
        "type": "text",
        "text": "%s",
        "wrap": true,
        "weight": "bold",
        "gravity": "center",
        "size": "xl"
      },
      {
        "type": "box",
        "layout": "vertical",
        "margin": "lg",
        "spacing": "sm",
        "contents": [
          {
            "type": "box",
            "layout": "baseline",
            "spacing": "sm",
            "contents": [
              {
                "type": "text",
                "text": "Size",
                "color": "#aaaaaa",
                "size": "sm",
                "flex": 1
              },
              {
                "type": "text",
                "text": "%s",
                "wrap": true,
                "size": "sm",
                "color": "#666666",
                "flex": 4
              }
            ]
          },
          {
            "type": "box",
            "layout": "baseline",
            "spacing": "sm",
            "contents": [
              {
                "type": "text",
                "text": "Price",
                "color": "#aaaaaa",
                "size": "sm",
                "flex": 1
              },
              {
                "type": "text",
                "text": "%s",
                "wrap": true,
                "color": "#666666",
                "size": "sm",
                "flex": 4
              }
            ]
          },
          {
            "type": "separator"
          }
        ]
      },
      {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "text",
            "text": "Shipping",
            "size": "md",
            "weight": "bold"
          },
          {
            "type": "box",
            "layout": "baseline",
            "contents": [
              {
                "type": "text",
                "text": "Address",
                "flex": 2,
                "color": "#aaaaaa",
                "size": "sm"
              },
              {
                "type": "text",
                "text": "%s %s",
                "flex": 4,
                "wrap": true,
                "color": "#666666",
                "size": "sm",
                "margin": "none"
              }
            ]
          },
          {
            "type": "box",
            "layout": "baseline",
            "contents": [
              {
                "type": "text",
                "text": "Customer",
                "flex": 2,
                "color": "#aaaaaa",
                "size": "sm"
              },
              {
                "type": "text",
                "text": "%s",
                "flex": 4,
                "wrap": true,
                "color": "#666666",
                "size": "sm",
                "margin": "none"
              }
            ]
          },
          {
            "type": "box",
            "layout": "baseline",
            "contents": [
              {
                "type": "text",
                "text": "Phone No.",
                "flex": 2,
                "color": "#aaaaaa",
                "size": "sm"
              },
              {
                "type": "text",
                "text": "%s",
                "flex": 4,
                "wrap": true,
                "color": "#666666",
                "size": "sm",
                "margin": "none"
              }
            ]
          }
        ]
      }
    ]
  }
}""" %(imglink, orderID, product, size, price, address, zip_code, customer, phone_no)
    flex = json.loads(flex)
    flex = FlexSendMessage(alt_text='Flex Message alt text', contents=flex)
    line_bot_api.reply_message(reply_token, flex)

    order_detail = [orderID, "Processing", product, size, price, customer, phone_no, address, zip_code]

    last = order_sheet.find("Last Order")
    order_sheet.update(f"A{last.row}:I{last.row}", [order_detail])
    order_sheet.update_cell(last.row+1, last.col, "Last Order")

def CheckStatus(reply_token, orderID):
    cells = order_sheet.find(orderID)
    item = order_sheet.row_values(cells.row)

    status = item[1]
    product = item[2]
    size = item[3]
    price = item[4]
    customer = item[5]
    phone_no = item[6]
    address = item[7]
    zip_code = item[8]

    flex = """{
  "type": "bubble",
  "size": "mega",
  "header": {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "box",
        "layout": "baseline",
        "contents": [
          {
            "type": "text",
            "text": "Order ID",
            "margin": "none",
            "color": "#ffffff",
            "flex": 2
          },
          {
            "type": "text",
            "text": "#%s",
            "flex": 4,
            "color": "#ffffff"
          }
        ]
      },
      {
        "type": "text",
        "text": "%s",
        "size": "sm",
        "color": "#ffffff",
        "margin": "none"
      },
      {
        "type": "text",
        "text": "Size: %s",
        "size": "sm",
        "color": "#ffffff",
        "margin": "none"
      },
      {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "text",
            "text": "FROM",
            "color": "#ffffff66",
            "size": "sm"
          },
          {
            "type": "text",
            "text": "RannRongTao",
            "color": "#ffffff",
            "size": "lg",
            "flex": 4,
            "weight": "bold"
          }
        ]
      },
      {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "text",
            "text": "TO",
            "color": "#ffffff66",
            "size": "sm"
          },
          {
            "type": "text",
            "text": "คุณ%s",
            "color": "#ffffff",
            "size": "sm"
          },
          {
            "type": "text",
            "text": "%s %s",
            "color": "#ffffff",
            "size": "sm",
            "flex": 4,
            "weight": "regular",
            "wrap": true
          },
          {
            "type": "text",
            "text": "เบอร์โทร %s",
            "size": "sm",
            "color": "#ffffff"
          }
        ]
      }
    ],
    "paddingAll": "20px",
    "backgroundColor": "#6742D8",
    "spacing": "md",
    "height": "250px",
    "paddingTop": "22px"
  },
  "body": {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "text",
        "text": "คุณจะได้รับสินค้าภายใน 3-5 วัน",
        "color": "#b7b7b7",
        "size": "xs"
      },
      {
        "type": "box",
        "layout": "horizontal",
        "contents": [
          {
            "type": "text",
            "text": "  ",
            "size": "sm"
          },
          {
            "type": "box",
            "layout": "vertical",
            "contents": [
              {
                "type": "filler"
              },
              {
                "type": "box",
                "layout": "vertical",
                "contents": [],
                "cornerRadius": "30px",
                "height": "12px",
                "width": "12px",
                "borderColor": "#EF454D",
                "borderWidth": "bold"
              },
              {
                "type": "filler"
              }
            ],
            "flex": 0
          },
          {
            "type": "text",
            "text": "กำลังจัดเตรียมสินค้า",
            "gravity": "center",
            "flex": 4,
            "size": "sm"
          }
        ],
        "spacing": "lg",
        "cornerRadius": "30px",
        "margin": "xl"
      },
      {
        "type": "box",
        "layout": "horizontal",
        "contents": [
          {
            "type": "box",
            "layout": "baseline",
            "contents": [
              {
                "type": "filler"
              }
            ],
            "flex": 1
          },
          {
            "type": "box",
            "layout": "vertical",
            "contents": [
              {
                "type": "box",
                "layout": "horizontal",
                "contents": [
                  {
                    "type": "filler"
                  },
                  {
                    "type": "box",
                    "layout": "vertical",
                    "contents": [],
                    "width": "2px",
                    "backgroundColor": "#EF454D"
                  },
                  {
                    "type": "filler"
                  }
                ],
                "flex": 1
              }
            ],
            "width": "12px"
          },
          {
            "type": "text",
            "text": "1 day",
            "gravity": "center",
            "flex": 4,
            "size": "xs",
            "color": "#8c8c8c"
          }
        ],
        "spacing": "lg",
        "height": "64px"
      },
      {
        "type": "box",
        "layout": "horizontal",
        "contents": [
          {
            "type": "box",
            "layout": "horizontal",
            "contents": [
              {
                "type": "text",
                "text": " ",
                "gravity": "center",
                "size": "sm"
              }
            ],
            "flex": 1
          },
          {
            "type": "box",
            "layout": "vertical",
            "contents": [
              {
                "type": "filler"
              },
              {
                "type": "box",
                "layout": "vertical",
                "contents": [],
                "cornerRadius": "30px",
                "width": "12px",
                "height": "12px",
                "borderWidth": "2px",
                "borderColor": "#8c8c8c"
              },
              {
                "type": "filler"
              }
            ],
            "flex": 0
          },
          {
            "type": "text",
            "text": "บริษัทขนส่งรับพัสดุเข้าระบบ",
            "gravity": "center",
            "flex": 4,
            "size": "sm"
          }
        ],
        "spacing": "lg",
        "cornerRadius": "30px"
      },
      {
        "type": "box",
        "layout": "horizontal",
        "contents": [
          {
            "type": "box",
            "layout": "baseline",
            "contents": [
              {
                "type": "filler"
              }
            ],
            "flex": 1
          },
          {
            "type": "box",
            "layout": "vertical",
            "contents": [
              {
                "type": "box",
                "layout": "horizontal",
                "contents": [
                  {
                    "type": "filler"
                  },
                  {
                    "type": "box",
                    "layout": "vertical",
                    "contents": [],
                    "width": "2px",
                    "backgroundColor": "#8c8c8c"
                  },
                  {
                    "type": "filler"
                  }
                ],
                "flex": 1
              }
            ],
            "width": "12px"
          },
          {
            "type": "text",
            "text": "2-3 days",
            "gravity": "center",
            "flex": 4,
            "size": "xs",
            "color": "#8c8c8c"
          }
        ],
        "spacing": "lg",
        "height": "64px"
      },
      {
        "type": "box",
        "layout": "horizontal",
        "contents": [
          {
            "type": "text",
            "text": " ",
            "gravity": "center",
            "size": "sm"
          },
          {
            "type": "box",
            "layout": "vertical",
            "contents": [
              {
                "type": "filler"
              },
              {
                "type": "box",
                "layout": "vertical",
                "contents": [],
                "cornerRadius": "30px",
                "width": "12px",
                "height": "12px",
                "borderColor": "#8c8c8c",
                "borderWidth": "2px"
              },
              {
                "type": "filler"
              }
            ],
            "flex": 0
          },
          {
            "type": "text",
            "text": "ลูกค้าได้รับสินค้าเรียบร้อย",
            "gravity": "center",
            "flex": 4,
            "size": "sm"
          }
        ],
        "spacing": "lg",
        "cornerRadius": "30px"
      }
    ]
  }
}""" %(orderID, product, size, customer, address, zip_code, phone_no)
    flex = json.loads(flex, strict=False)

    flex = FlexSendMessage(alt_text='Flex Message alt text', contents=flex)
    line_bot_api.reply_message(reply_token, flex)