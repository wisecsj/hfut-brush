"""
v 2.0
使用于各种情况，部分获取参数方法改成了用bs4
"""
# coding= utf-8
import re
import requests
import xlrd
import urllib
import time
from bs4 import BeautifulSoup
import getpass

save_url = "http://tkkc.hfut.edu.cn/student/exam/manageExam.do?1479131327464&method=saveAnswer"
# index用于提示题目序号
index = 1
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:51.0) Gecko/20100101 Firefox/41.0",
           "Host": "tkkc.hfut.edu.cn",
           "X-Requested-With": "XMLHttpRequest",
           }

ses = requests.session()
# ses.headers.update({"Cookie": "JSESSIONID=D182E6F264EBAEB5D423F11E38B15E21.tomcat2"})
ID = input("请输入学号\n")
Pwd = getpass.getpass("请输入密码\n")

login_url = "http://tkkc.hfut.edu.cn/login.do?"


def getcode():
    im = ses.get("http://tkkc.hfut.edu.cn/getRandomImage.do")
    tmp1 = urllib.parse.quote_from_bytes(im.content)
    code = ses.post('http://api.hfutoyj.cn/codeapi', data={'image': tmp1})  # return verify code
    return code


rlt = True
while rlt:
    code = getcode().text
    logInfo = {

        "logname": ID,
        "password": Pwd,
        "randomCode": code
    }
    res = ses.post(login_url, data=logInfo, headers=headers)
    if res.text.find("验证码错误") == -1:
        rlt = False
    time.sleep(0.01)

# 用于存放excel中question, answer键值对的字典
result = dict()


# retries默认为2，表示尝试次数。以防某种原因，某次连接失败
def craw(url, retries=2):
    try:
        b = ses.post(url, headers=headers)
        b.encoding = 'utf-8'
        d = b.text
        title = re.findall(r'&nbsp;(.*?)","', d, re.S)[0]
        return title
    except Exception as e:
        print(e)
        if retries > 0:
            return craw(url, retries=retries - 1)
        else:
            print("get failed", index)
            return ''


# 从字典中根据题目找到并返回答案
def answer_func(t):
    return result.get(title, "Not Found")


# 将找到的答案提交给服务器
def submit(ans, id, id2, id3, id4, index, retries=2):
    dx = ["false", "false", "false", "false", "false"]
    try:
        if ans.find('A') != -1:
            dx[0] = "true"
        if ans.find('B') != -1:
            dx[1] = "true"
        if ans.find('C') != -1:
            dx[2] = "true"
        if ans.find('D') != -1:
            dx[3] = "true"
        if ans.find('E') != -1:
            dx[4] = "true"
        if ans.find('正确') != -1:
            ans = "A"
        if ans.find('错误') != -1:
            ans = "B"
        data2 = {"examReplyId": id3,
                 "examStudentExerciseId": id2,
                 "exerciseId": id,
                 "examId": id4,
                 "DXanswer": ans,
                 "duoxAnswer": ans,
                 "PDanswer": ans,
                 "DuoXanswerA": dx[0],
                 "DuoXanswerB": dx[1],
                 "DuoXanswerC": dx[2],
                 "DuoXanswerD": dx[3],
                 "DuoXanswerE": dx[4],
                 "DuoXanswer": ans}  # 部分题库的多选是分成5个来提交，还有的是只用一个进行提交
        body = ses.post(save_url, data=data2, headers=headers)
        wb_data = body.text
        print(wb_data, index)
    except Exception as e:
        print(e)
        if retries > 0:
            return submit(ans, id, id2, id3, id4, index, retries=retries - 1)
        else:
            print("get failed", index)
            return ''


# 此变量用于判断用户是否要继续刷课
finished = 0

while finished == 0:
    start_url = input("请输入测试页面URL\n")

    myfile = xlrd.open_workbook('exercise.xls')
    lenOfXls = len(myfile.sheets())

    # 读取XLS中的题目和答案，存进字典（将这段程序放在这，是因为当用户有多门试题库时，刷完一门，切换到另一门时，不用关闭程序只需切换题库Excel即可）
    for x in range(0, lenOfXls):
        xls = myfile.sheets()[x]
        for i in range(1, xls.nrows):
            title = xls.cell(i, 0).value
            if x == 1 and lenOfXls == 2:
                answer = xls.cell(i, 2).value
            elif x == 1 and lenOfXls == 3:
                answer = xls.cell(i, 7).value
            elif x == 2 and lenOfXls == 3:
                answer = xls.cell(i, 2).value
            else:
                answer = xls.cell(i, 7).value
            result[title] = answer

    body = ses.get(start_url, headers=headers)
    body.encoding = 'utf-8'
    wb_data = body.text
    # print(wb_data)
    urlId = re.findall(r'do\?(.*?)&method', start_url, re.S)[0]

    eval = re.findall(r'eval(.*?)]\);', wb_data, re.S)[0]

    bs = BeautifulSoup(wb_data, 'lxml')
    val = bs.form.input
    examReplyId = val['value']

    examId = re.findall(r'<input type="hidden" name="examId" id="examId" value="(.*?)" />', wb_data, re.S)[0]

    exerciseId = re.findall(r'exerciseId":(.*?),', eval, re.S)

    examSEId = re.findall(r'examStudentExerciseId":(.*?),', eval, re.S)

    examStudentExerciseId = re.findall(r'"examStudentExerciseId":(.*?),"exerciseId"',
                                       wb_data, re.S)[0]

    examStudentExerciseId = int(examStudentExerciseId)

    # id对应exerciseID,id2对应examStudetExerciseId
    for id in exerciseId:
        next_url = r"http://tkkc.hfut.edu.cn/student/exam/manageExam.do?%s&method=getExerciseInfo&examReplyId=%s&exerciseId=%s&examStudentExerciseId=%d" % (
            urlId, examReplyId, id, examStudentExerciseId)
        title = craw(next_url)
        ans = answer_func(title)
        submit(ans, id, examStudentExerciseId, examReplyId, examId, index)
        # time.sleep(1)
        index += 1
        examStudentExerciseId = examStudentExerciseId + 1
    # input函数获取到的为字符串，所以进行Type conversion
    finished = int(input("继续请输入0，退出请输入1\n"))
