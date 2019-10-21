import jieba.analyse  
import numpy as np
import pandas as pd
from pyecharts import options as opts
from pyecharts.globals import SymbolType
from pyecharts.charts import Pie, Bar, WordCloud

'''
author: BW
vision: 1.1
'''

excel_path = '豆瓣上海堡垒短评.xlsx'
short_comments = pd.read_excel(excel_path)  
STOP_WORDS_FILE_PATH = 'stop_words.txt'    # 清洗词

# 数据清洗，去掉无效词
jieba.analyse.set_stop_words(STOP_WORDS_FILE_PATH)

# 1、词数统计
keywords_count_list_TR = jieba.analyse.textrank(' '.join(short_comments.Comment), topK=50, withWeight=True)   # jieba.analyse.textrank: 基于 TextRank 算法的关键词抽取
print(keywords_count_list_TR)
# 生成词云
word_cloud_TR = (
    WordCloud()
        .add("", keywords_count_list_TR, word_size_range=[20, 100], shape=SymbolType.DIAMOND)
        .set_global_opts(title_opts=opts.TitleOpts(title="上海堡垒短评词云TOP50"))
)
word_cloud_TR.render('comment-word-TR-cloud.html')

keywords_count_list_TI = jieba.analyse.extract_tags(' '.join(short_comments.Comment), topK=50, withWeight=True)  # 基于 TF-IDF（term frequency–inverse document frequency） 算法的关键词抽取
print(keywords_count_list_TI)
word_cloud_TI = (   
    WordCloud()
        .add("", keywords_count_list_TI, word_size_range=[20, 100], shape="cardioid")
        .set_global_opts(title_opts=opts.TitleOpts(title="上海堡垒短评词云TOP50"))
)
word_cloud_TI.render('comment-word-TI-cloud.html')    # 结果看 TI 比 TR 好一点

# 2、短评词频分析生成柱状图  # 以 keywords_count_list_TR 为例
# 2.1统计词数
# 取前20高频的关键词
keywords_count_dict = {i[0]: 0 for i in reversed(keywords_count_list_TR[:20])}  
cut_words = jieba.cut(' '.join(short_comments.Comment))
for word in cut_words:
    for keyword in keywords_count_dict.keys():
        if word == keyword:
            keywords_count_dict[keyword] = keywords_count_dict[keyword] + 1
print(keywords_count_dict)
# 2.2生成柱状图
keywords_count_bar = (
    Bar()
        .add_xaxis(list(keywords_count_dict.keys()))                
        .add_yaxis("", list(keywords_count_dict.values()), category_gap="40%")    
        .reversal_axis()                                                        
        .set_series_opts(label_opts=opts.LabelOpts(position="right"))          
        .set_global_opts(
        title_opts=opts.TitleOpts(title="上海堡垒短评热词TOP20"),
        yaxis_opts=opts.AxisOpts(name="次数"),
        xaxis_opts=opts.AxisOpts(name="热词")
        )
    )
keywords_count_bar.render('comment-word-count-bar.html')

# 3、生成饼状图
cut_words = jieba.cut(' '.join(short_comments.Comment))   
keywords_count_list = []
for i in reversed(keywords_count_list_TR[:20]):
    keywords_count_list.append(i[0])
print(keywords_count_list)
keywords_count_result = [0 for x in range(20)]
for word in cut_words:
    for j in range(20):
        if word == keywords_count_list[j]:
            keywords_count_result[j] = keywords_count_result[j] + 1
print(keywords_count_result)
keywords_count_pie = (
    Pie()
        .add("", [list(z) for z in zip(keywords_count_list, keywords_count_result)])
        .set_global_opts(title_opts=opts.TitleOpts(title="上海堡垒短评热词TOP20"))
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {c}"))
)
keywords_count_pie.render('comment-word-count-pie.html')

# 4、评星统计
rating = [0 for x in range(5)]
for result in short_comments.Rating:
    if result == '很差':
        rating[0] += 1
    elif result == '较差':
        rating[1] += 1
    elif result == '还行':
        rating[2] += 1
    elif result == '推荐':
        rating[3] += 1
    else:
        rating[4] += 1
print('很差' + str(rating[0]) + '个，较差' + str(rating[1]) + '个，还行' + str(rating[2]) + \
    '个，推荐' + str(rating[3]) + '个，力荐' + str(rating[4]) + '个。')

rating_bar = (
    Bar()
        .add_xaxis(['很差', '较差', '还行', '推荐', '力荐'])                
        .add_yaxis("", rating)    
        .set_global_opts(
        title_opts=opts.TitleOpts(title="上海堡垒短评评星频数"),
        yaxis_opts=opts.AxisOpts(name="数量"),
        xaxis_opts=opts.AxisOpts(name="评星")
        )
    )
rating_bar.render('comment-rating-bar.html')

rating_pie = (
    Pie(init_opts=opts.InitOpts(theme=ThemeType.WALDEN))
        .add("", [list(z) for z in zip(['很差', '较差', '还行', '推荐', '力荐'], rating)])
        .set_global_opts(title_opts=opts.TitleOpts(title="上海堡垒短评评星频数"))
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {c}"))
        .render('comment-rating-pie.html')
)

rating_pie = (
    Pie()
        .add("", [list(z) for z in zip(['很差', '较差', '还行', '推荐', '力荐'], rating)],
            radius=["30%", "75%"],
            center=["40%", "50%"],
            rosetype="radius")
        .set_global_opts(
            title_opts=opts.TitleOpts(title="上海堡垒短评评星频数"),
            legend_opts=opts.LegendOpts(type_="scroll", pos_left="80%", orient="vertical")
            )
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {c}"))
        .render('comment-rating-pie玫瑰图R.html'.format(word))
    )

rating_pie = (
    Pie()
        .add("", [list(z) for z in zip(['很差', '较差', '还行', '推荐', '力荐'], rating)],
            radius=["30%", "75%"],
            center=["40%", "50%"],
            rosetype="area")
        .set_global_opts(
            title_opts=opts.TitleOpts(title="上海堡垒短评评星频数"),
            legend_opts=opts.LegendOpts(type_="scroll", pos_left="70%", orient="vertical")
            )
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {c}"))
        .render('comment-rating-pie玫瑰图A.html'.format(word))
    )


