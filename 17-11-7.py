# coding= utf-8
import re
import os
import requests
import xlrd
import urllib
import time
from bs4 import BeautifulSoup
import zipfile
import getpass

save_url = "http://tkkc.hfut.edu.cn/student/exam/manageExam.do?1479131327464&method=saveAnswer"
index = 1
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:51.0) Gecko/20100101 Firefox/41.0",
           "Host": "tkkc.hfut.edu.cn",
           "X-Requested-With": "XMLHttpRequest",
           # 'Content-Type': 'application/json,text/javascript,*/*'

           }
ses = requests.session()
ID = input("请输入学号\n")
Pwd = getpass.getpass("请输入密码\n")
login_url = "http://tkkc.hfut.edu.cn/login.do?"
main_url = "http://tkkc.hfut.edu.cn"
main1_url = "http://tkkc.hfut.edu.cn/student/index.do"


def getcode():
    im = ses.get("http://tkkc.hfut.edu.cn/getRandomImage.do")
    tmp1 = urllib.parse.quote_from_bytes(im.content)
    code = ses.post('http://api.hfutoyj.cn/codeapi', data={'image': tmp1})  # return verify code
    return code


def get_new_data():
    r = ses.get(login_url).text
    announce = re.findall(r'name="(.*?)" value="announce"', r)[0]
    return announce


rlt = 10
times = 1
while rlt > 3:
    announce = get_new_data()
    code = getcode().text
    print("Trying " + ' ' + code)
    logInfo = {
        announce: 'announce',
        'loginMethod': '{}button'.format(announce),
        "logname": ID,
        "password": Pwd,
        "randomCode": code
    }
    res = ses.post(login_url, data=logInfo, headers=headers)
    # print(res.text)
    time.sleep(0.01)
    times += 1
    rlt -= 1
    if res.text.find("验证码错误") != -1:
        print("Wrong verify code, Trying again ...")
        continue
    elif res.text.find("身份验证服务器未建立连接") != -1:
        print("Wrong student number, Check and reinput please ...")
        ID = input("请重新输入学号\n")
        continue
    elif res.text.find("密码不正确") != -1:
        print("Wrong password, Check and reinput please ...")
        Pwd = getpass.getpass("请输入密码\n")
        continue
    else:
        print('Login Success !')
        break

else:
    print("Maybe you typed wrong password")

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
def answer_func(title):
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


finished = 2
while (finished == 2) or (finished == 0):
    if finished == 2:
        html = ses.get(main1_url)
        html = html.text
        s = re.findall(r'courseId=(\d+)', html, re.M)
        print("获取课程ID成功")
        ss = list(set(s))
        ss.sort(key=s.index)
        print(ss)
        courseId = input("请输入上一句打印出来的课程Id来进行自动化讨论和做题(如全部刷完请输入0退出):")
        temp = courseId
        course_url = main_url + "/student/teachingTask/coursehomepage.do?courseId=" + courseId
        course_url = ses.get(course_url).text
        # print("获取TaskId成功")
        s1 = re.findall(r'teachingTaskId=(\d+)', course_url, re.M)
        TaskId_url = main_url + "/student/resource/index.do?teachingTaskId=" + s1[0]
        taskhomepage_url = main_url + "/student/teachingTask/taskhomepage.do?&teachingTaskId=" + s1[0]
        taskhomepage_url = ses.get(taskhomepage_url).text
        TaskId_url = ses.get(TaskId_url).text
        faq_url = main_url + "/student/bbs/index.do?teachingTaskId=" + s1[0]
        faq_url = ses.get(faq_url).text
        s4 = re.findall(r'discussId=(\d+)', faq_url, re.M)
        s5 = re.findall(r'forumId=(\d+)', faq_url, re.M)
        discuss_url = main_url + "/student/bbs/manageDiscuss.do?&method=view&teachingTaskId=" + s1[0] + "&discussId=" + \
                      s4[0]
        discuss_url = ses.get(discuss_url).text
        post_url = main_url + "/student/bbs/manageDiscuss.do?method=reply"
        soup = BeautifulSoup(discuss_url, "lxml")
        content = soup.find_all("p")
        if len(content):
            content = content[-1]
            post_data = {'discussId': s4[0], 'forumId': s5[0], 'type': 1, 'teachingTaskId': s1[0], 'content': content}
        else:
            discuss_url = main_url + "/student/bbs/manageDiscuss.do?&method=view&teachingTaskId=" + s1[
                0] + "&discussId=" + s4[2]
            discuss_url = ses.get(discuss_url).text
            soup = BeautifulSoup(discuss_url, "lxml")
            content = soup.find_all("p")
            content = content[-1]
            post_data = {'discussId': s4[2], 'forumId': s5[0], 'type': 1, 'teachingTaskId': s1[0], 'content': content}
        # print (content)


        p = ses.post(post_url, data=post_data, headers=headers)
        # print ("参加讨论成功一次")
        discuss_url1 = main_url + "/student/bbs/manageDiscuss.do?&method=view&teachingTaskId=" + s1[0] + "&discussId=" + \
                       s4[1]
        discuss_url1 = ses.get(discuss_url1).text
        soup = BeautifulSoup(discuss_url1, "lxml")
        content1 = soup.find_all("p")
        if len(content1):
            content1 = content1[-1]
            post_data1 = {'discussId': s4[1], 'forumId': s5[0], 'type': 1, 'teachingTaskId': s1[0], 'content': content1}
        else:
            discuss_url1 = main_url + "/student/bbs/manageDiscuss.do?&method=view&teachingTaskId=" + s1[
                0] + "&discussId=" + s4[3]
            discuss_url1 = ses.get(discuss_url1).text
            soup = BeautifulSoup(discuss_url1, "lxml")
            content1 = soup.find_all("p")
            content1 = content1[-1]
            post_data1 = {'discussId': s4[3], 'forumId': s5[0], 'type': 1, 'teachingTaskId': s1[0], 'content': content1}
        # content1 = content1[-1]
        # print (content1)

        p1 = ses.post(post_url, data=post_data1, headers=headers)
        print("参加讨论成功两次")
        finished = 0
    if finished == 0:
        html = ses.get(main1_url).text
        s = re.findall(r'courseId=(\d+)', html, re.M)
        # print("获取课程ID成功")
        ss = list(set(s))
        ss.sort(key=s.index)
        # print (ss)
        courseId = temp
        # print("获取TaskId成功")
        s1 = re.findall(r'teachingTaskId=(\d+)', course_url, re.M)
        s2 = re.findall(r'"id":(\d+)', TaskId_url, re.M)
        down_url = main_url + "/filePreviewServlet?indirect=true&resourceId=" + s2[0]
        # print("获取下载链接成功")
        d = ses.get(down_url)
        with open("excel.zip", "wb") as code:
            code.write(d.content)
        file_list = os.listdir(r'.')
        for file_name in file_list:
            if os.path.splitext(file_name)[1] == '.zip':
                print("下载题库文件并解压完成")

                file_zip = zipfile.ZipFile(file_name, 'r')
                for file in file_zip.namelist():
                    file_zip.extract(file, r'.')
                file_zip.close()
                os.remove(file_name)
        s3 = re.findall(r'assignmentId=(\d+)', taskhomepage_url, re.M)
        # print (s3)
        s4 = re.findall(r'examId=(\d+)', taskhomepage_url, re.M)
        # exam_url = main_url+"/student/exam/manageExam.do?&method=doExam&examId="+s4[0]
        test_url = main_url + "/student/assignment/manageAssignment.do?method=doAssignment&assignmentId=" + s3[0]
        test_url2 = main_url + "/student/assignment/manageAssignment.do?method=doAssignment&assignmentId=" + s3[1]
        test_url3 = main_url + "/student/assignment/manageAssignment.do?method=doAssignment&assignmentId=" + s3[2]
        if len(s3) == 3:
            start_url_list = [test_url, test_url2, test_url3]
        elif len(s3) == 4:
            test_url4 = main_url + "/student/assignment/manageAssignment.do?method=doAssignment&assignmentId=" + s3[3]
            start_url_list = [test_url, test_url2, test_url3, test_url4]
        else:
            test_url5 = main_url + "/student/assignment/manageAssignment.do?method=doAssignment&assignmentId=" + s3[4]
            start_url_list = [test_url, test_url2, test_url3, test_url4, test_url5]
        # print("获取练习题目链接成功")
        # print (test_url,"\n",test_url2,'\n',test_url3)
        # start_url_list = [test_url,test_url2,test_url3,test_url4]
        for start_url in start_url_list:
            # print (exam_url)
            # start_url = input("请输入练习题目链接(就在上面↑)\n")
            myfile = xlrd.open_workbook('exercise.xls')
            lenOfXls = len(myfile.sheets())
            # 存储sheet名字的列表
            sheet_names = myfile.sheet_names()
            # 题库excel文件的类型
            # 3：单 双 判断
            # 2：单 双
            # 3：单 判断
            if len(sheet_names) == 3:
                excel_type = 3
            elif '多选题' in sheet_names:
                excel_type = 2
            else:
                excel_type = 1
            # 读取XLS中的题目和答案，存进字典（将这段程序放在这，是因为当用户有多门试题库时，刷完一门，切换到另一门时，不用关闭程序只需切换题库Excel即可）
            for x in range(0, lenOfXls):
                xls = myfile.sheets()[x]
                for i in range(1, xls.nrows):
                    title = xls.cell(i, 0).value.strip()
                    if x == 1 and lenOfXls == 2:
                        if excel_type == 2:
                            answer = xls.cell(i, 7).value
                        else:
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
                next_url = r"http://tkkc.hfut.edu.cn/student/exam/manageExam.do?method=getExerciseInfo&examReplyId=%s&exerciseId=%s&examStudentExerciseId=%d" % (
                    examReplyId, id, examStudentExerciseId)
                title = craw(next_url).strip()
                ans = answer_func(title)
                submit(ans, id, examStudentExerciseId, examReplyId, examId, index)
                # time.sleep(1)
                index += 1
                examStudentExerciseId = examStudentExerciseId + 1
            # input函数获取到的为字符串，所以进行Type conversion
            finished = 2
