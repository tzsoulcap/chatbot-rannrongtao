from flask import Flask, request
from linebot.models import *
from linebot import *
import test
import json
import re

app = Flask(__name__)

# channel access token and channel secret
line_bot_api = LineBotApi(
    "AM82YvNzOu37BSjeLy5LtvUbDZIdwssqEU4kTuTg7aDEUfrE9MqVoLhAqAT4H43Ggk5Bo9qC2mRRypGGhXpr694K+yxLf7IO7eIK5+CWaKLbsqKz2osEOR5QASQ7RPyjL0EOOV+MfsbDKP1fH3B9CwdB04t89/1O/w1cDnyilFU=")
handler = WebhookHandler("bffecec9ba651ad932eaf02906ea659c")

@app.route("/")
def home():
    return "Flask with heroku"

@app.route("/callback", methods=['POST'])
def callback():
    body = request.get_data(as_text=True)
    # print(body)
    req = request.get_json(silent=True, force=True)
    intent = req["queryResult"]["intent"]["displayName"]
    text = req['originalDetectIntentRequest']['payload']['data']['message']['text']
    reply_token = req['originalDetectIntentRequest']['payload']['data']['replyToken']
    id = req['originalDetectIntentRequest']['payload']['data']['source']['userId']
    disname = line_bot_api.get_profile(id).display_name
    # brand = req["queryResult"]["parameters"]["Brands"]

    # print(req)
    # print('brand = ' + brand)
    print('id = ' + id)
    print('name = ' + disname)
    print('text = ' + text)
    print('intent = ' + intent)
    print('reply_token = ' + reply_token)

    

    if intent == 'SelectBrands':
        brand = req["queryResult"]["parameters"]["Brands"]
        test.showItem(intent, reply_token, brand, id)

    elif intent == 'SelectSizes':
        product = req["queryResult"]["parameters"]["ProductName"]
        test.selectSizes(intent, reply_token, product, id)

    elif intent == 'ConfirmSize':
        product = req["queryResult"]["outputContexts"][1]["parameters"]["ProductName"]
        size = req["queryResult"]["parameters"]["Size"]
        # size = req["queryResult"]["parameters"]["Sizing"]
        test.showBuyDetail(intent, reply_token, product, size, id)

    elif intent == 'Shipping Address':
        for item in req["queryResult"]["outputContexts"]:
            if re.search("selectsizes-followup", item["name"]):
                product = item["parameters"]["ProductName"]
                size = item["parameters"]["Size"]
                # size = item["parameters"]["Sizing"]
            break

        # product = req["queryResult"]["outputContexts"][0]["parameters"]["ProductName"]
        # size = req["queryResult"]["outputContexts"][0]["parameters"]["Size"]
        customer = req["queryResult"]["parameters"]["person"]["name"]
        # lastname = req["queryResult"]["parameters"]["last-name"]
        # customer = req["queryResult"]["parameters"]["given-name"]
        address = req["queryResult"]["parameters"]["address"]
        zip_code = req["queryResult"]["parameters"]["zip-code"]
        phone_no = req["queryResult"]["parameters"]["phone-number"]
        test.ShippingAddress(reply_token, product, size, customer, address, zip_code, phone_no)

    elif intent == 'Order Number':
        orderID = req["queryResult"]["parameters"]["OrderID"]
        # print(orderID)
        test.CheckStatus(reply_token, orderID)

    elif intent == 'Sole Shields' or intent == 'Shoes Spa':
        test.Service(reply_token, id, intent)
    #     test.reserveService(reply_token)

    elif intent == 'Confirm Reserve':
        for item in req["queryResult"]["outputContexts"]:
            if re.search("soleshields-followup", item["name"]):
                service = item["parameters"]["Service"]

            break

        time = req["queryResult"]["parameters"]["time"]
        date = req["queryResult"]["parameters"]["date"]
        test.ReserveService(reply_token, date, time, service, disname)

    elif intent == 'Send Shoes to The Shop':
        test.SentShoes(reply_token)

    return 'OK'



def reply(intent, text, reply_token, id, disname):
    if intent == 'Test':
        text_message = TextSendMessage(text='ทดสอบสำเร็จ')
        line_bot_api.reply_message(reply_token, text_message)


def replyflex(intent, reply_token):
    flex_message = FlexSendMessage(
        alt_text='hello',
        contents={
            'type': 'bubble',
            'direction': 'ltr',
            'hero': {
                'type': 'image',
                'url': 'https://stackpython.co/media/django-summernote/2021-03-10/ba57b98f-535e-42f8-a6fc-a47a5355fe16.jpg',
                'size': 'full',
                'aspectRatio': '20:13',
                'aspectMode': 'cover',
                'action': {'type': 'uri', 'uri': 'https://stackpython.co/tutorial/python-chatbot-line-dialogflow-flask-heroku-ep2', 'label': 'label'}
            }
        }
    )
    line_bot_api.reply_message(reply_token, flex_message)


if __name__ == '__main__':
    app.run()
