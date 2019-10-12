import jieba
import numpy as np
import pandas as pd
import wordcloud 
from PIL import Image

'''
author: BW
vision: 1.0
'''

excel_path = '豆瓣上海堡垒短评.xlsx'
short_comments = pd.read_excel(excel_path)  

# 切割中文
result = jieba.lcut(' '.join(short_comments.Comment))
# 打开图片
imageobj = Image.open('背景.jpg')  # 任意图片作背景
cloud_mask = np.array(imageobj)
# 绘制词云
wc = wordcloud.WordCloud(
    mask=cloud_mask,
    background_color='white',
    font_path='simfang.ttf',    # 处理中文数据时
    min_font_size=5,      # 图片中最小字体大小
    max_font_size=50,     # 图片中最大字体大小
    width=500,            # 指定生成图片的宽度
)
wc.generate(','.join(result))
# 保存图片
wc.to_file('D:/douban.png')  
