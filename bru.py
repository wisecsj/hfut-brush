# -*- coding: utf-8 -*-
"""
Refactor hfut-brush.
"""
import re
import requests
import xlrd
import urllib
import time
from bs4 import BeautifulSoup
import getpass


class Brush:
    _ses = requests.session()
    save_url = "http://tkkc.hfut.edu.cn/student/exam/manageExam.do?1479131327464&method=saveAnswer"
    index = 1  # record probleme brushed count
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:51.0) Gecko/20100101 Firefox/41.0",
               "Host": "tkkc.hfut.edu.cn",
               "X-Requested-With": "XMLHttpRequest",
               }
    login_url = "http://tkkc.hfut.edu.cn/login.do?"
    max_times = 6
    answers_dict = dict()
    excel_path = 'exercise.xls'  # where excel file store

    def __init__(self, **kwargs):
        self.ID = input("请输入学号\n")
        self.Pwd = getpass.getpass("请输入密码\n")
        tmp = ['save_url', 'headers', 'login_url', 'max_times', 'excel_path']
        if kwargs:
            self.__dict__.update(kwargs)

    def login(self):
        """to login tkkc.hfut.edu.cn
        """
        last_times = self.max_times
        times = 1
        while last_times:
            announce = self.get_new_data()
            code = self.get_verify_code().text
            print("Trying " + ' ' + code)
            logInfo = {
                announce: 'announce',
                'loginMethod': '{}button'.format(announce),
                "logname": self.ID,
                "password": self.Pwd,
                "randomCode": code
            }
            res = self.ses.post(self.login_url, data=logInfo, headers=self.headers)
            times += 1
            last_times -= 1
            # print(res.text)
            if res.text.find("验证码错误") != -1:
                print("Wrong verify code, Trying again ...")
                continue
            elif res.text.find("身份验证服务器未建立连接") != -1:
                print("Wrong student number, Check and reinput please ...")
                ID = input("请输入学号\n")
                continue
            elif res.text.find("密码不正确") != -1:
                print("Wrong password, Check and reinput please ...")
                Pwd = getpass.getpass("请输入密码\n")
                continue
            else:
                print('Login Success !')
                break
                # time.sleep(0.01)

        else:
            print("Login failed after trying the max_times")

    @property
    def ses(self):
        return self._ses

    def get_verify_code(self):
        """return verify code from myself api
        """
        im = self.ses.get("http://tkkc.hfut.edu.cn/getRandomImage.do")
        tmp1 = urllib.parse.quote_from_bytes(im.content)
        code = self.ses.post('http://api.hfutoyj.cn/codeapi', data={'image': tmp1})
        return code

    def get_new_data(self):
        """As website add a new parameter 'announce' in login process
        """
        r = self.ses.get(self.login_url).text
        announce = re.findall(r'name="(.*?)" value="announce"', r)[0]
        return announce

    def craw(self, url, retries=2):
        """craw the question text and return '' if failed """
        try:
            b = self.ses.post(url, headers=self.headers)
            b.encoding = 'utf-8'
            d = b.text
            title = re.findall(r'&nbsp;(.*?)","', d, re.S)[0]
            return title
        except Exception as e:
            print(e)
            if retries > 0:
                return self.craw(url, retries=retries - 1)
            else:
                print("craw the {}th question failed".format(self.index))
                return ''

    def submit(self, ans, id, id2, id3, id4, index, retries=2):
        """submit the answer found in Excel to server"""
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

    def scan_excel(self):
        f = xlrd.open_workbook('exercise.xls')
        sheets_num = len(f.sheets())

        # 读取XLS中的题目和答案，存进字典
        for x in range(0, sheets_num):
            xls = f.sheets()[x]
            for i in range(1, xls.nrows):
                title = xls.cell(i, 0).value
                if x == 1 and sheets_num == 2:
                    answer = xls.cell(i, 2).value
                elif x == 1 and sheets_num == 3:
                    answer = xls.cell(i, 7).value
                elif x == 2 and sheets_num == 3:
                    answer = xls.cell(i, 2).value
                else:
                    answer = xls.cell(i, 7).value
                self.result[title] = answer

    def get_submit_info(self):
        pass

    def start(self):
        finished = False
        self.login()
        while not finished:
            self.main()
            # input func returns str type, so we need ype conversion
            finished = True if int(input("继续请输入0，退出请输入1\n")) else False

    def main(self):
        """main logic"""
        start_url = input("请输入测试页面URL\n")

        body = self.ses.get(start_url, headers=self.headers)
        body.encoding = 'utf-8'
        wb_data = body.text
        # print(wb_data)
        urlId = re.findall(r'do\?(.*?)&method', self.start_url, re.S)[0]

        eval = re.findall(r'eval(.*?)]\);', wb_data, re.S)[0]

        bs = BeautifulSoup(wb_data, 'lxml')
        val = bs.form.input
        examReplyId = val['value']

        examId = re.findall(r'<input type="hidden" name="examId" id="examId" value="(.*?)" />',
                            wb_data, re.S)[0]

        exerciseId = re.findall(r'exerciseId":(.*?),', eval, re.S)

        examSEId = re.findall(r'examStudentExerciseId":(.*?),', eval, re.S)

        examStudentExerciseId = re.findall(r'"examStudentExerciseId":(.*?),"exerciseId"',
                                           wb_data, re.S)[0]

        examStudentExerciseId = int(examStudentExerciseId)

        # id对应exerciseID,id2对应examStudetExerciseId
        for id in exerciseId:
            next_url = r"http://tkkc.hfut.edu.cn/student/exam/manageExam.do?%s&method=getExerciseInfo&examReplyId=%s&\
            exerciseId=%s&examStudentExerciseId=%d" % (
                urlId, examReplyId, id, examStudentExerciseId)
            title = self.craw(next_url)
            ans = self.answers_dict.get[title, 'Not found']
            self.submit(ans, id, examStudentExerciseId, examReplyId, examId, index)
            # time.sleep(1)
            self.index += 1
            examStudentExerciseId = examStudentExerciseId + 1
        # input函数获取到的为字符串，所以进行Type conversion
        finished = int(input("继续请输入0，退出请输入1\n"))


if __name__ == '__main__':
    brush = Brush()
    brush.start()
