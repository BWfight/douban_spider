import requests, openpyxl, csv
from bs4 import BeautifulSoup
import os

'''
author: BW
vision: 1.0
'''

with open('豆瓣250.csv', "w", newline='') as csv_f:
    writer = csv.writer(csv_f)
    writer.writerow(['排名', '电影名', '评价', '推荐语', '链接'])

wb=openpyxl.Workbook()    # 创建工作薄
sheet=wb.active   # 获取工作薄的活动表
sheet.title="Movies"    # 工作表重命名
sheet['A1'] ='排名'     #加表头，给单元格赋值
sheet['B1'] ='电影名'   
sheet['C1'] ='评价'   
sheet['D1'] ='推荐语'    
sheet['E1'] ='链接'    

for i in range(10):
    url = 'https://movie.douban.com/top250?start=' + str(i*25) + '&filter='
    res = requests.get(url)
    # print(res.status_code) 

    bs = BeautifulSoup(res.text,'html.parser')
    all_info = bs.find('ol', class_="grid_view")
    info_list = all_info.find_all('li')
    for info in info_list:
        rank = info.find('em').text                            # 排名
        title = info.find('span', class_="title").text         # 电影名
        try:  
            recommendation = info.find('p',class_='quote').text   # 推荐语，有的电影没有推荐语
        except:
            recommendation = ''
        comment = info.find('span',class_="rating_num").text   # 评价
        link = info.find('a')['href']                          # 链接

        with open('豆瓣250.txt',"a+") as file:
            string = rank+' '+title+' '+comment+'\n'+recommendation+'\n'+link+'\n====================================\n'
            file.write(string.encode('gbk','ignore').decode('gbk')) 

        with open('豆瓣250.csv', "a+", newline='',encoding='gbk') as csv_f:
            writer = csv.writer(csv_f)
            writer.writerow([rank, \
                             title.encode('gbk','ignore').decode('gbk'), \
                             comment.encode('gbk','ignore').decode('gbk'), \
                             recommendation.encode('gbk','ignore').decode('gbk'), \
                             link])
       
        sheet.append([rank, \
                      title.encode('gbk','ignore').decode('gbk'), \
                      comment.encode('gbk','ignore').decode('gbk'), \
                      recommendation.encode('gbk','ignore').decode('gbk'), \
                      link])

wb.save('豆瓣250.xlsx')
