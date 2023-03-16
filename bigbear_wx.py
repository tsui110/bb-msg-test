# 微信所有操作都在里面
from bigbear_ini import wx_api_url
import requests
import json

print(wx_api_url)
headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/537.36 (KHTML, '
                  'like Gecko) Chrome/80.0.3987.163 Safari/537.36',
    'content-type': 'application/json',
    'Connection': 'close'
}


# 发送文本消息
def sdtxt(towho, msg):
    data = json.dumps({
        "type": "sendMsg",
        "data": {
            "wxid": towho,
            "msg": msg
        }

    }, ensure_ascii=False)
    r = requests.post(wx_api_url, headers=headers, data=data.encode("utf-8"))
    r = r.json()
    print(r)


# 发送文件
def sdfile(wxid, path):
    payload = json.dumps({
        "type": "sendFile",
        "data": {
            "wxid": wxid,
            "path": path
        }
    }, ensure_ascii=False)
    response = requests.post(wx_api_url, headers=headers, data=payload.encode("utf-8"))
    print(response.text)


# 获取对象信息
def getinfo(wxid):
    payload = json.dumps({
        "type": "getFriendInfo",
        "data": {
            "wxid": wxid
        }
    })
    response = requests.post(wx_api_url, headers=headers, data=payload)
    return response.json()


# 获取群名称
def get_gp_nickname(mode, wxid):
    if mode == 'nick':
        getname = getinfo(wxid)
        return getname['result']['nick']
    else:
        getname = getinfo(wxid)
        return getname['result']['nickBrief']


# 同意转账
def accepttransfer(wxid, transid):
    payload = json.dumps({
        "type": "acceptTransfer",
        "data": {
            "wxid": wxid,
            "transferid": transid
        }
    })
    response = requests.post(wx_api_url, headers=headers, data=payload)
    print(response.text)
