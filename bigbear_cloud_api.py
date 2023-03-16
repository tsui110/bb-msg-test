import requests
headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/537.36 (KHTML, '
                      'like Gecko) Chrome/80.0.3987.163 Safari/537.36',
        'content-type': 'application/json',
        'Connection': 'close'
    }
# 添加ITEM到云服务器
def addgoodslist(item,fromwho,ow):
    dizhi = 'https://91e787c8-df7a-42ad-b4fa-12dde982b53d.bspapp.com/infos'
    data = {
            "mode": 'addGoodsList',
            "item": item,
            "ow": ow,
            "fromwho":fromwho

        }
    data=str(data)
    data=data.encode('utf-8')
    r = requests.post(dizhi, headers=headers,
                     data=data)
    return r.json()
#从云服务器删除Item
def delitemlist(item,ow,fromwho):
    dizhi = 'https://91e787c8-df7a-42ad-b4fa-12dde982b53d.bspapp.com/remove'
    data = {
        "mode": 'delGoodsList',
        "item": item,
        "ow": ow,
        "fromwho": fromwho

    }
    data = str(data)
    data = data.encode('utf-8')
    r = requests.post(dizhi, headers=headers,
                      data=data)
    return r.json()
#更新云数据运单状态
def updateexp(ono, statuscode, eno=None):
    dizhi = 'https://91e787c8-df7a-42ad-b4fa-12dde982b53d.bspapp.com/infos'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/80.0.3987.163 Safari/537.36',
        'content-type': 'application/json'
    }
    if eno == None:
        print('Eno is none')
        data = {
            "mode": 'updateExp',
            "orderNo": ono,
            'status': statuscode

        }
    else:
        data = {
            "mode": 'updateExp',
            "orderNo": ono,
            "expNo": eno,
            'status': statuscode

        }
    data=str(data)
    data=data.encode('utf-8')
    r = requests.post(dizhi, headers=headers,
                     data=data)
#获取运单号，判断运单号是否已经存在
def getexpno(expno):
    dizhi = 'https://91e787c8-df7a-42ad-b4fa-12dde982b53d.bspapp.com/infos'
    data={
                         "mode": 'getExp',
                         "expNo": expno
                     }
    data = str(data)
    data = data.encode('utf-8')
    r = requests.post(dizhi, headers=headers,
                     data=data)
    # 获取反馈信息，如果存在则对单元格进行删除操作
    # r.json()  类型为dict
    print(r.json())
    rdict = r.json()
    if rdict['affectedDocs'] > 0:
        return True
    else:
        return False
#添加文本订单
def addtextorder(yuanwen, fromwho, ow):
    dizhi = 'https://91e787c8-df7a-42ad-b4fa-12dde982b53d.bspapp.com/infos'
    data = {
            "mode": 'addOrderText',
            "status": 0,
            "yuanwen": yuanwen,
            "ow": ow,
            "fromwho":fromwho

        }
    data=str(data)
    data=data.encode('utf-8')
    r = requests.post(dizhi, headers=headers,data=data)

#获取订单号状态
def getstatus(orderno,statuscode):
    dizhi = 'https://91e787c8-df7a-42ad-b4fa-12dde982b53d.bspapp.com/infos'
    data={
            "mode":'getStatus',
            "orderNo": orderno,
        }
    data = str(data)
    data = data.encode('utf-8')
    r = requests.post(dizhi, headers=headers,
                     data=data)
    # 获取反馈信息，如果存在则对单元格进行删除操作
    # r.json()  类型为dict
    print(r.json())
    rdict = r.json()
    if rdict['affectedDocs']>0:
        #如果当前确实没发货
        if int(rdict['data'][0]['status'])==int(statuscode):
            print('确定未发货')
            return True
    else:
        return False