from bigbear_mysql import dosql
from bigbear_base import logger
from bigbear_cloud_api import addgoodslist
import os
import sys
import shutil
logger.info("sql,开始获取初始化信息")
sql = "select * from settings"
# 管理员列表
adminlist = []
# 帮卖列表，集合去重复
listenlist = set()
result = dosql(sql)
moren = ''
ghid = ''
qjltoken = ''
wx_api_url = ''
wx_api_token = ""
rid = ''
itemlist = set()
for i in result:
    if i['skeys'] == 'moren':
        moren = i['svalues']
        # print("当前默认消息接收人为：{}".format(moren,))
    if i['skeys'] == 'rid':
        rid = i['svalues']
        # print("当前默认助手ID：{}".format(rid,))
    if i['skeys'] == 'ghid':
        ghid = i['svalues']
        # print("当前群接龙ID为：{}".format(ghid,))
    if i['skeys'] == 'qjltoken':
        qjltoken = i['svalues']
        # print("当前群接龙token为：{}".format(qjltoken,))
    if i['skeys'] == 'url':
        wx_api_url = i['svalues']
        # print("当前调用网址为：{}".format(url,))
    if i['skeys'] == 'apitoken':
        wx_api_token = i['svalues']
        # print("当前调用token：{}".format(stoken,))
if rid == "" or wx_api_url == "" or ghid == "" or qjltoken == "" or moren == "" or wx_api_token == "":
    logger.info("检测到未配置默认值，程序将退出。")
    sys.exit()
# 开始处理C盘bigbear目录下的表格文件
allfiles = os.listdir(path=r'C:\bigbear')

for i in allfiles:
    if 'xls' in i or 'xlsx' in i:
        print(i)
        if os.path.isdir(r'C:\backupxls'):
            shutil.move(r'C:\bigbear' + "\\" + i, r'C:\backupxls' + "\\" + i)
            logger.info("当前文件名：" + i + "已被移动到C:\\backupxls目录下")
        else:
            os.makedirs(r'C:\backupxls')
            shutil.move(r'C:\bigbear' + "\\" + i, r'C:\backupxls' + "\\" + i)
            logger.info("当前文件名：" + i + "已被移动到C:\\backupxls目录下")

# 获取adminlist
sql = "select * from adminforwx"
result = dosql(sql)
for i in result:
    adminlist.append(i['id'])
print('当前adminlist为：{}'.format(adminlist))

# 获取帮卖列表
sql = "select * from agentlist"
result = dosql(sql)
for i in result:
    listenlist.add(i['id'])
print("当前listenlist为：{}".format(listenlist))
wx_api_url = wx_api_url + rid
# 获取商品列表
goodslist = set()
sql = "select * from goodslist"
result = dosql(sql)
for i in result:
    itemlist.add(i['id'])
    goodslist.add(i['item'])
    addgoodslist(i['item'], i['id'], rid)
# itemlist为供货商群号
print("当前goodlist为：{}".format(goodslist))
print("当前的goodslist长度为：{}".format(len(goodslist)))
print("当前itemlist为：{}".format(itemlist))
