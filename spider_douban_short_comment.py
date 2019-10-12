import requests
import openpyxl
import random
import time
from bs4 import BeautifulSoup

'''
author: BW
vision: 1.0
'''

wb = openpyxl.Workbook()  # 创建工作薄
sheet = wb.active  # 获取工作薄的活动表
sheet.title = "短评"  # 工作表重命名
sheet['A1'] = 'UserName'  #加表头，给单元格赋值
sheet['B1'] = 'Rating'
sheet['C1'] = 'Comment'
sheet['D1'] = 'Votes'

headers = {
    'user-agent':
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'
}

for i in range(15):    # 页数
    url = 'https://movie.douban.com/subject/26581837/comments?start=' + str(20 * i) + '&limit=20&sort=new_score&status=P'  # 以上海堡垒为例
    res = requests.get(url, headers=headers)
    res.encoding = 'utf-8'
    if res.status_code == 200:
        print('爬取第' + str(1 + i) + '页')
        bs = BeautifulSoup(res.text, 'html.parser')
        info_list = bs.find_all('div', class_='comment-item')
        for info in info_list:
            UserName = info.find('div', class_='avatar').find('a')['title']
            Votes = info.find('span', class_="votes").text
            Rating = info.find('div', class_='comment').find('h3').find('span', class_="comment-info").find(
                    'span', class_="rating")['title'] 
            Comment = info.find('span', class_="short").text

            sheet.append([UserName, Rating, Comment, Votes])
            time.sleep(random.randint(2,9))
    else:
        print('第' + str(1 + i) + '页爬取失败') 

wb.save('豆瓣上海堡垒短评.xlsx')   
