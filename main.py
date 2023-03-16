# -*- coding: utf-8 -*-
# change name
import psutil
import requests
from flask_cors import cross_origin
import json
from bigbear_ini import listenlist, adminlist, itemlist, rid, moren, logger, goodslist
from bigbear_excel import check_upload_excel, tongbushuju_gp, file_analy_agent
import os
from bigbear_mysql import dosql
from bigbear_cloud_api import addtextorder, delitemlist
from bigbear_wx import sdfile, sdtxt, get_gp_nickname, accepttransfer
from bigbear_base import phone, checktextorder
import re
import traceback
import time
from pathlib import Path
from flask import Flask
from flask import request, jsonify
import pythoncom
from win32com.client import DispatchEx

headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/80.0.3987.163 Safari/537.36',
        'content-type': 'application/json'
    }
url = "http://127.0.0.1:7777/httpapi/?wxid=ty200947752"
filelist = []
txt_order_url = r'http://bigbear.d1shequ.com/erp/'
newfile = ''
oldfile = ''
totalitems = 0


# call return statement
def callret():
    return jsonify("{status:1}")


app = Flask(__name__)
curdir = os.path.dirname(os.path.realpath(__file__))


#@app.after_request
#def foot_log(environ):
#    for i in filelist:
#        try:
#            print(filelist)
#            os.remove(i)
#           filelist.pop(filelist.index(i))
#             print(filelist)
#         except BaseException as e:
#             print(e)
#             print('删除失败', i)
#     return environ


@app.route('/msg', methods=['POST'])
def return_msg():
    print(request)


@app.route('/callback', methods=['POST', 'GET'])
def return_json():
    if request.method == "POST":
        try:
            # 获取原始数据并进行URLDECODE
            data_json = request.json
            del data_json['data']['data']['msgBase64']
            print(data_json['data'])
            datas = data_json['data']
            eventType = data_json['event']
            # 群聊或者私聊事件
            if eventType == 10008 or eventType == 10009:
                msg_datas = datas['data']

                if 'msg' in msg_datas:
                    # 判断事件类型
                    # msgType 事件类型
                    # fromWxid 来源群
                    # finalFromWxid 来源微信个人
                    # todo 群聊消息处理，2022年10月13日06:33:23
                    # 10009表示私聊 10008表示群聊
                    if eventType == 10008:
                        #监听@命令，设置帮卖群信息
                        if msg_datas['msgType'] == 1 and msg_datas['finalFromWxid'] in adminlist:
                            if "@开启表格订单" in msg_datas['msg']:
                                id=msg_datas['fromWxid']
                                if id in itemlist:
                                    sdtxt(id,"设置失败，错误原因：当前群已经被设定为其他角色，无法开启表格订单")
                                else:
                                    #确保当前id并未在帮卖列表
                                    if id not in listenlist:
                                        nick=get_gp_nickname('nick',id)
                                        sql='INSERT INTO 帮卖群 ( id, nickname)VALUES("%s","%s")'%(id,nick)
                                        try:
                                            ret=dosql(sql)
                                            sdtxt(id,"设置成功，感谢使用大笨熊助手，祝各位老板大卖特卖")
                                            listenlist.add(id)
                                        except Exception as e:
                                            sdtxt(id,'设置失败，错误原因：未知')
                                    else:
                                        sdtxt(id,"已经开启表格订单，请勿重复设置")


                        # 监听帮卖群的文本订单
                        if msg_datas['fromWxid'] in listenlist and msg_datas['msgType'] == 1 and msg_datas[
                            'fromWxid'] != rid and msg_datas['fromWxid'] not in adminlist:
                            # 判断当前是否有需要的文本订单信息
                            msgs = msg_datas['msg']
                            patt = '.+1[0-9]{10}'
                            res = re.findall(patt, msgs)
                            if len(res) > 0 and len(msgs) < 100:
                                # 查找到电话号码
                                res = re.findall('[0-9]{11}', res[0])
                                res = res[0]
                                ret = phone(res)
                                if ret:
                                    ret = checktextorder(msgs)
                                    print(ret)
                                    if 'fail' in ret:
                                        print('没有订单信息')
                                    else:
                                        # 有订单信息
                                        print(ret)
                                        addtextorder(msg_datas['msg'], msg_datas['fromWxid'], rid)
                                        msgtxt = "识别到可能是文本订单的数据，请及时确认订单信息：" + txt_order_url + "?id=" + \
                                                 msg_datas['fromWxid'] + "&ow=" + rid
                                        sdtxt(msg_datas['fromWxid'], msgtxt)
                                        sdtxt(moren, "检测到有文本订单：{}".format(msg_datas['msg']))
                                        print('有订单信息')
                                else:
                                    # 判断电话号码无效则什么都不做
                                    print('当前号码无效')
                                    pass
                            else:
                                # 不可能是文本订单就啥都不做了
                                pass

                        # itemlist是供货商的群，监听供货商回单表格，并上传转发
                        # 不能是管理员发和机器人自己发出来的文件
                        #新增判断条件<?xml version="1.0"?>避免引用的情况下，误认为表格回单
                        xmltxt='<?xml version="1.0"?>'
                        if msg_datas['msgType'] == 49 and msg_datas['fromWxid'] in itemlist and msg_datas[
                            'finalFromWxid'] not in adminlist and msg_datas['finalFromWxid'] != rid and xmltxt not in msg_datas['msg']:
                            sdtxt(msg_datas['fromWxid'],'开始处理回单表格')
                            pathstr = msg_datas['msg']
                            ppos = pathstr.find('=')
                            flag=True
                            # 截取文件路径
                            if ppos > 0:
                                pathstr = pathstr[ppos + 1:-1]
                                newpathstr = Path(pathstr)
                                print(newpathstr)
                                fullpath = pathstr
                                logger.warning(fullpath)
                                flag = True
                                counttime = 0
                                while flag:
                                    if newpathstr.exists():
                                        flag = False
                                        break
                                    else:
                                        time.sleep(10)
                                        logger.warning('waiting for the file complete')
                                        counttime = counttime + 1
                                        if counttime > 10:
                                            sdtxt(msg_datas['fromWxid'],"当前文件下载超时，或对方已取消发送文件,文件路径：%s"%(newpathstr))
                                            break
                                # 截取文件名做比较
                                if '.xls' not in pathstr:
                                    sdtxt(msg_datas['fromWxid'],'不支持的格式，请使用.xls或xlsx的文件！')
                                    # sdatmsg(robot,msg_datas['fromWxid'],msg_datas['fromWxid'],'不支持的格式，请使用.xls或xlsx的文件！')
                                else:
                                    if '.xlsx' not in pathstr:
                                        logger.warning('修改类型')
                                        pythoncom.CoInitialize()
                                        excel = DispatchEx('Excel.Application')
                                        wb = excel.Workbooks.Open(pathstr)
                                        fullpath = pathstr + 'x'
                                        wb.SaveAs(fullpath, FileFormat=51)
                                        wb.Close()
                                        excel.Application.Quit()
                                        pythoncom.CoUninitialize()
                                    pos = fullpath.rfind("\\")
                                    # 开始检测文件
                                    filename = fullpath[pos + 1:]
                                    ret = check_upload_excel(msg_datas['fromWxid'], filename, fullpath,msg_datas['finalFromWxid'])
                                    print("当前订单处理完毕")
                                    logger.info("回单表格更新完毕")
                                    if ret:
                                        sdtxt(msg_datas['fromWxid'],"回单表格处理完毕，部分成功,文件名：%s"%(filename))
                                        #filelist.append(fullpath)
                                    else:
                                        sdtxt(msg_datas['fromWxid'],"回单表格处理完毕，部分失败")

                            callret()

                            # 检查内容是否为空
                        # 监听帮卖群提交的订单表格，并转发
                        if msg_datas['msgType'] == 49 and msg_datas['fromWxid'] in listenlist and msg_datas[
                            'finalFromWxid'] not in adminlist and msg_datas['finalFromWxid'] != rid:
                            # 开始判断

                            sdtxt(msg_datas['fromWxid'],'发现代理提交的订单表格')
                            pathstr = msg_datas['msg']
                            ppos = pathstr.find('=')
                            # 截取文件路径
                            if ppos > 0:
                                pathstr = pathstr[ppos + 1:-1]
                                newpathstr = Path(pathstr)
                                print(newpathstr)
                                fullpath = pathstr
                                logger.warning(fullpath)
                                flag = True
                                counttime = 0
                                while flag:
                                    if newpathstr.exists():
                                        flag = False
                                        break
                                    else:
                                        time.sleep(0.5)
                                        logger.warning('waiting for the file complete')
                                        counttime = counttime + 1
                                        if counttime > 200:
                                            break
                                # 截取文件名做比较
                                if '.xls' not in pathstr:
                                    sdtxt(msg_datas['fromWxid'],'不支持的格式，请使用.xls或xlsx的文件！')
                                    # sdatmsg(robot,msg_datas['fromWxid'],msg_datas['fromWxid'],'不支持的格式，请使用.xls或xlsx的文件！')
                                else:
                                    if '.xlsx' not in pathstr:
                                        logger.warning('修改类型')
                                        pythoncom.CoInitialize()
                                        excel = DispatchEx('Excel.Application')
                                        wb = excel.Workbooks.Open(pathstr)
                                        fullpath = pathstr + 'x'
                                        wb.SaveAs(fullpath, FileFormat=51)
                                        wb.Close()
                                        excel.Application.Quit()
                                        pythoncom.CoUninitialize()
                                    pos = fullpath.rfind("\\")
                                    # 开始同步订单，并转发
                                    filename = fullpath[pos + 1:]
                                    tongbushuju_gp(fullpath, msg_datas['fromWxid'])
                                    file_analy_agent(fullpath, msg_datas['fromWxid'])
                            callret()
                        # 配置开团通知列表
                        # 配置商品列表项目
                        if '@配置商品+' in msg_datas['msg'] and msg_datas['finalFromWxid'] in adminlist:
                            logger.warning('配置商品列表项')
                            print("开始配置商品列表")
                            nick = get_gp_nickname('nick', msg_datas['fromWxid'])
                            itemname = msg_datas['msg'][6:]
                            if itemname not in goodslist:
                                sql = "INSERT INTO goodslist (id,item,nickname)VALUES('%s','%s','%s')" % (
                                    msg_datas['fromWxid'], itemname, nick)
                                print(sql)
                                insert_result = dosql(sql)
                                print(insert_result)
                                # 更新当前goodslist
                                goodslist.add(itemname)
                                sdtxt(msg_datas['fromWxid'], itemname + '商品添加成功！')
                            else:
                                sdtxt(msg_datas['fromWxid'], itemname + '商品项已经存在，请勿重复添加！')
                        if '@删除商品+' in msg_datas['msg'] and msg_datas['finalFromWxid'] in adminlist:
                            print('删除商品')
                            itemname = msg_datas['msg'][6:]
                            if itemname in goodslist:
                                sql = "DELETE FROM goodslist WHERE item='{}'".format(itemname)
                                d_ret = dosql(sql)
                                print(d_ret)
                                goodslist.remove(itemname)
                                print(len(goodslist))
                                sdtxt(msg_datas['fromWxid'], itemname + '：该商品已删除！')
                            else:
                                sdtxt(msg_datas['fromWxid'], itemname + '：该商品不存在')
                            delitemlist(itemname, rid, msg_datas['fromWxid'])

                    # todo 私聊消息处理  2022年10月13日06:32:33
                    if eventType == 10009:
                        #print("return 后当前语句不会执行")
                        if msg_datas['fromWxid'] == 'gh_cd5f251d7089':
                            if '有人申请取消接龙' in msg_datas['msg']:
                                sdtxt(moren, '群接龙有客户发起售后，请及时处理。')
                        # 转发消息

                        # 如果是文本消息

                    # 当前为群邀请

            if eventType == 10014:
                if "type" in datas and datas['type'] == 1:
                    print("账号登录")
                elif datas['type'] == 0:
                    print("账号下线")
            # 当前为转账事件
            if eventType == 10006:
                msg_datas = datas['data']

                accepttransfer(msg_datas['fromWxid'], msg_datas['transferid'])

        except BaseException as e:
            print(e)
            traceback.print_exc()

    else:

        logger.warning('当前为调用')

    return jsonify({"status": "1"})
@app.route('/lovelycat',methods=["POST"])
@cross_origin()
def forlovelycat():
    data=request.json
    ret=dosql(data['sql'])
    print(ret)
    return jsonify({"data":ret})
@app.route('/sql',methods=["POST"])
@cross_origin()
def get_sqls():  # put application's code here
    print(request.json)
    rjson=request.json

    #如果是当前的选择allTables
    if 'mode' in rjson and rjson['mode']=="allTables":
        ret=dosql(rjson['sql'])
        print(ret)
        retlist=[]
        result={}
        index=1
        for i in ret:
            retobj = {}
            retobj['index']=index
            retobj['name']=i['TABLE_NAME']
            retlist.append(retobj)
            index=index+1
        result['data']=retlist
        print(result)
        return result
    #如果查询单个表格
    if 'mode' in rjson and rjson['mode']=="selectOne":
        ret=dosql(rjson['sql'])
        print(ret)
        result={}
        result['data']=ret
        print(result)
        return result
    #更新ID
    if 'mode' in rjson and rjson['mode']=="updateID":
        try:
            sql = rjson['sql']
            sql = sql.replace("nickname='null'", "nickname is null")
            ret=dosql(sql)
            print(ret)
            result = {}
            result['data'] = {"msg":"更新ID成功"}
            print(result)
            return result

        except Exception as e:
            print(e)
            result = {}
            result['data'] = {"msg": "更新ID失败"}
            print(result)
            return result
    #删除一行
    if 'mode' in rjson and rjson['mode']=="dropOneRow":
        try:
            sql=rjson['sql']
            sql=sql.replace("nickname='null'","nickname is null")
            ret=dosql(sql)
            print(ret)
            result = {}
            result['data'] = {"msg":"删除ID成功"}
            print(result)
            return result

        except Exception as e:
            print(e)
            result = {}
            result['data'] = {"msg": "删除ID失败"}
            print(result)
            return result
    # 新增一行
    if 'mode' in rjson and rjson['mode'] == "insertOneRow":
        try:
            ret = dosql(rjson['sql'])
            print(ret)
            result = {}
            result['data'] = {"msg": "新增ID成功"}
            print(result)
            return result

        except Exception as e:
            print(e)
            result = {}
            result['data'] = {"msg": "新增ID失败"}
            print(result)
            return result


    return 'Hello World!'
@app.route('/functions',methods=["POST"])
@cross_origin()
def get_funs():
    print(request.json)
    rjson=request.json
    if "func" in rjson and rjson['func'] == "checkWx":
        obj={}
        plist = psutil.pids()
        print("当前进程id:{}".format(os.getpid()))
        for i in plist:
            pi = psutil.Process(i)
            if '大笨熊微信管家.exe' in pi.name():
                obj['bigbearWx']=True
            if 'WeChat.exe' in pi.name():
                obj['wechat']=True
        print(obj)
        return obj
    if "func" in rjson and rjson['func'] == "getUserInfo":
        obj={}
        payload = json.dumps({
            "type": "getUserInfo",
            "data": {
            }})
        response = requests.post(url, headers=headers, data=payload)
        obj['data']=response.json()
        return obj
    return 'Hello World!'


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=9000)
