import requests
import smtplib
import schedule
import time
from bs4 import BeautifulSoup
from email.mime.text import MIMEText
from email.header import Header
import random

'''
author: BW
vision: 1.0
'''

def get_movie():
    rank_list = []
    title_list = []
    recommendation_list = []
    comment_list = []
    link_list = []

    for n in range(3):
        i = random.randint(0,9)
        url = 'https://movie.douban.com/top250?start=' + str(i*25) + '&filter='
        res = requests.get(url)

        bs = BeautifulSoup(res.text,'html.parser')
        all_info = bs.find('ol', class_="grid_view")
        info_list = all_info.find_all('li')

        j = random.randint(0,24)
        info = info_list[j]
        rank = info.find('em').text                          
        title = info.find('span', class_="title").text        
        try:  
            recommendation = info.find('p',class_='quote').text   # 有的电影没有推荐语
        except:
            recommendation = ''
        comment = info.find('span',class_="rating_num").text   
        link = info.find('a')['href']  

        rank_list.append(rank)
        title_list.append(title)
        recommendation_list.append(recommendation)
        comment_list.append(comment)
        link_list.append(link)

    return rank_list, title_list, recommendation_list, comment_list, link_list

def send_email(rank_list, title_list, recommendation_list, comment_list, link_list):
    account = input('账号：')
    password = input('密码：')
    receiver = []     # 列表
    ct = 'y'
    while ct == 'y':
        rec = input('收件人地址：')
        receiver.append(rec)
        ct == input('是否输入其他收件人地址？继续输入"y"，否则不输入')
        

    mailhost = 'smtp.office365.com'  # 换成相应服务器
    server = smtplib.SMTP(mailhost)
    server.connect(mailhost, 25)    # 换成相应端口
    server.starttls()
    server.login(account, password)

    msg_list = []
    for i in range(3):
        content = rank_list[i] + '. ' + title_list[i] + '\n' + recommendation_list[i] + '\n分数：' + comment_list[i] + '\n\n' + link_list[i] +'\n=========\n'
        msg_list.append(content)
    msg = ''.join(msg_list).encode('gbk','ignore').decode('gbk')
    message = MIMEText(msg, 'plain', 'utf-8')
    message['Subject'] = Header('每周电影推荐', 'utf-8')

    try:
        server.sendmail(account, receiver, message.as_string())
        print ('邮件发送成功')
    except:
        print ('邮件发送失败')
    server.quit()

def job():
    rank_list, title_list, recommendation_list, comment_list, link_list = get_movie()
    send_email(rank_list, title_list, recommendation_list, comment_list, link_list)
    print('今日任务完成')

schedule.every().friday.at("18:30").do(job) 
while True:
    schedule.run_pending()
    time.sleep(1)
