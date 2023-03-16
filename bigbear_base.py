import difflib
import re
import cpca
import pandas
import jionlp as jio
import time
import os
import socket
import logging

adcode = ''
sheng = ''
shi = ''
qu = ''
addressDetail = ''
shoujianren = ''
analylist = []


# 比较相似度
def string_similar(s1, s2):
    return difflib.SequenceMatcher(None, s1, s2).quick_ratio()


def phone(n):
    if re.match(r'13[0,1,2]\d{8}', n) or \
            re.match(r'14[0,4,5]\d{8}', n) or \
            re.match(r"15[5,6]\d{8}", n) or \
            re.match(r"16[6,7]\d{8}", n) or \
            re.match(r"17[5,6]\d{8}", n) or \
            re.match(r"18[5,6]\d{8}", n) or \
            re.match(r"196\d{8}", n) or \
            re.match(r"170[4,7,8,9]\d{7}", n) or \
            re.match(r"171[0-9]\d{7}", n):
        print("该号码属于：中国联通")
        return True
    # 中国移动
    #
    elif re.match(r"13[5-9]\d{8}", n) or \
            re.match(r"134[0-8]\d{7}", n) or \
            re.match(r"14[4,7,8]]\d{8}", n) or \
            re.match(r"15[0,1,2,7,8,9]\d{8}", n) or \
            re.match(r"165\d{8}", n) or \
            re.match(r"170[3,5,6]\d{7}", n) or \
            re.match(r"172\d{8}", n) or \
            re.match(r"178\d{8}", n) or \
            re.match(r"18[2,3,4,7,8]\d{8}", n) or \
            re.match(r"19[5,7,8]\d{8}", n) or \
            re.match(r"1440\d{7}", n):
        print("该号码属于：中国移动")
        return True
    elif re.match(r"133\d{8}", n) or \
            re.match(r"14[1,9]\d{8}", n) or \
            re.match(r"153\d{8}", n) or \
            re.match(r"162\d{8}", n) or \
            re.match(r"170[0-3]\d{7}", n) or \
            re.match(r"162\d{8}", n) or \
            re.match(r"1740\d{7}", n) or \
            re.match(r"177\d{8}", n) or \
            re.match(r"18[0,1,9]\d{8}", n) or \
            re.match(r"19[0,1,3,9]\d{8}", n) or \
            re.match(r"1349\d{7}", n):
        # 中国电信
        # 190、191、193、199、1349
        print("该号码属于：中国电信")
        return True
    elif re.match(r"192\d{8}", n):
        print(
            '该号码属于：中国广电'
        )
        return True
    else:
        print('未匹配正确手机号')
        return False


# 检测文本中是否有电话信息以及地址信息
def checktextorder(msgs):
    global sheng, shi, qu, addressDetail
    a = 0
    addlist = []
    analylist = []
    patt = '1[0-9]{10}'
    res = re.findall(patt, msgs)
    if len(res) > 0:
        # 查找到电话号码
        print("当前查找到的号码个数：" + str(len(res)))
        for i in res:
            print(i)
            res = re.findall('1[0-9]{10}', i)
            res = res[0]
            ret = phone(res)
            if ret:
                a = a + 1
        if a > 0:
            # 当前文本判断有电话号码，进一步判断是否有省市区信息
            msgs = msgs.replace(" ", "")
            posz = msgs.find(res)
            stra = msgs[0:posz]
            strb = msgs[posz:]
            # print(stra)
            # print(strb)
            analylist.append(stra)
            analylist.append(strb)
            df = cpca.transform(analylist)
            # print(df)
            df = pandas.DataFrame(df)
            print(df)
            dflist = df.to_dict(orient='records')
            countflag = 2
            countNone = 0
            for i in range(0, len(analylist)):
                if dflist[i]['省'] is None:
                    countNone = countNone + 1
                else:
                    sheng = dflist[i]['省']
                if dflist[i]['市'] is None:
                    countNone = countNone + 1
                else:
                    shi = dflist[i]['市']
                if dflist[i]['区'] is None:
                    countNone = countNone + 1
                else:
                    qu = dflist[i]['区']
                if dflist[i]['地址'] is None:
                    countNone = countNone + 1
                else:
                    addressDetail = dflist[i]['地址']
                if dflist[i]['adcode'] is None:
                    countNone = countNone + 1

                if countNone < countflag:
                    # 如果判断地址大于等于2则表示当前确实有地址
                    addobj = {"phone": res, "sheng": sheng, "shi": shi, "qu": qu, "addr": addressDetail}
                    print("有订单信息")
                    addlist.append(addobj)
                else:
                    print('订单信息不完整')
                    countNone = 0
            if len(addlist) > 0:
                print(addlist)
                return addlist
            else:
                resultobj = {"fail": "true"}
                print(resultobj)
                return resultobj
        else:
            # 判断电话号码无效则什么都不做
            print('当前号码无效')
            resultobj = {"fail": "true"}
            print(resultobj)
            return resultobj


    else:
        # 不可能是文本订单就啥都不做了
        resultobj = {"fail": "true"}
        print(resultobj)
        return resultobj


# 首字母大写
def chinesetoupper(text):
    tet = jio.extract_chinese(text)
    strtext = ''
    for i in tet:
        strtext = strtext + i
    from xpinyin import Pinyin
    p = Pinyin()
    result1 = p.get_pinyin(strtext)
    s = result1.split('-')
    result3 = ''.join([i[0].upper() for i in s])
    return result3


# 返回当前日期
def daymonth():
    s = time.strftime("%m.%d", time.localtime())
    return s


# 判断端口是否被占用
class Logging():
    def make_log_dir(self, dirname='runtime_logs'):  # 创建存放日志的目录，并返回目录的路径
        now_dir = os.path.dirname(__file__)
        father_path = os.path.split(now_dir)[0]
        path = os.path.join(father_path, dirname)
        path = os.path.normpath(path)
        if not os.path.exists(path):
            os.mkdir(path)
        print(path)
        return path

    def get_log_filename(self):  # 创建日志文件的文件名格式，便于区分每天的日志
        filename = "{}.log".format(time.strftime("%Y-%m-%d", time.localtime()))
        filename = os.path.join(self.make_log_dir(), filename)
        filename = os.path.normpath(filename)
        return filename

    def log(self, level='DEBUG'):  # 生成日志的主方法,传入对那些级别及以上的日志进行处理
        logger = logging.getLogger()  # 创建日志器
        levle = getattr(logging, level)  # 获取日志模块的的级别对象属性
        logger.setLevel(level)  # 设置日志级别
        if not logger.handlers:  # 作用,防止重新生成处理器
            sh = logging.StreamHandler()  # 创建控制台日志处理器
            fh = logging.FileHandler(filename=self.get_log_filename(), mode='a', encoding="utf-8")  # 创建日志文件处理器
            # 创建格式器
            fmt = logging.Formatter("%(asctime)s-%(levelname)s-%(filename)s-Line:%(lineno)d-Message:%(message)s")
            # 给处理器添加格式
            sh.setFormatter(fmt=fmt)
            fh.setFormatter(fmt=fmt)
            # 给日志器添加处理器，过滤器一般在工作中用的比较少，如果需要精确过滤，可以使用过滤器
            logger.addHandler(sh)
            logger.addHandler(fh)
        return logger  # 返回日志器


logger = Logging().log(level='INFO')


def net_is_used(port, ip='127.0.0.1'):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        # 已启用
        s.connect((ip, port))
        s.shutdown(2)
        logger.warning('%s:%d is used' % (ip, port))
        return True
    except BaseException as e:
        print(e)
        # 未启用
        logger.warning('%s:%d is unused' % (ip, port))
        return False
