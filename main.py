import os
import time
import json
import requests as requests
from flask import Flask

app = Flask(__name__)
app.config['SECRET_KEY'] = "请填写您的飞书SECRET_KEY"

# 测试Flask
@app.route('/')
def hello_world():
    return 'Hello World!'


def get_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/app_access_token/internal"
    headers = {
        'Content-Type': "application/json; charset=utf-8"
    }
    params = {
        "app_id": "XXXXXXXXXXX",  # 需要替换为飞书开放平台APP ID
        "app_secret": "XXXXXXXXXXX"  # 需要替换为飞书开放平台APP Secret
    }
    resp = requests.post(url, params=params, headers=headers)
    data = resp.json()
    return data['tenant_access_token']

def get_sheet_content(file_url, sheet_id, tenant_access_token):
    spread_sheet_token: object = file_url.split('/')[4].split('?')[0]
    url = 'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/' + spread_sheet_token + '/values/' + sheet_id
    params = {"valueRenderOption": "ToString",
              "dateTimeRenderOption": "FormattedString"
              }
    headers = {
        'Authorization': 'Bearer ' + tenant_access_token,
        'Content-Type': "application/json; charset=utf-8"
    }
    resp = requests.get(url, params=params, headers=headers)
    return resp.json()

# 读取飞书Excel并组装为Json
@app.route('/sync/get_sheets')
def get_sheets():
    tenant_access_token = get_token()
    to_sheet_content = get_sheet_content("替换为您的飞书表格文档连接", "表格的sheetid", tenant_access_token) # 需要替换飞书共享表格连接 + sheet id
    maas_sheet = []
    for from_row in to_sheet_content['data']['valueRange']['values']:
        maas_sheet_line = {}
        for i in range((len(from_row))):
            maas_sheet_line[to_sheet_content['data']['valueRange']['values'][0][i]] = from_row[i]
        maas_sheet.append(maas_sheet_line)
    del maas_sheet[0]  # 删除首行
    print(json.dumps(maas_sheet))
    return json.dumps(maas_sheet)  # 将list列表组装为json格式


if __name__ == '__main__':
    app.run(port=5001, host="0.0.0.0")  # host = 0.0.0.0实现公网输出
