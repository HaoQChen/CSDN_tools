# -*- coding: utf-8 -*-
from xlutils.copy import copy
import xlrd
import xlwt
import os
import datetime
import requests
import re

#输入博客地址、页数以及保存路径
base_url = "https://blog.csdn.net/u013834525/"
pages = 3
record_file = "/home/night_fury/Desktop/CSDN_statistics.xls"


def main():
    titles = []
    visits = []
    comments = []
    data2save = []

    for x in range(pages):
        r = requests.get(base_url + "article/list/" + str(x + 1) + '?')
        titles = re.findall(r'<span class="article-type type-.*?">\n.*?</span>\n(.*?)</a>', r.content.decode(), re.MULTILINE)
        visits = re.findall( r'<span class="read-num">阅读数 <span class="num">(.*?)</span> </span>', r.content.decode())
        comments = re.findall( r'<span class="read-num">评论数 <span class="num">(.*?)</span> </span>', r.content.decode())
        #去掉空格存放到data2save
        for i in range(len(titles)):
            titles[i] = titles[i].strip()
            data2save.append((titles[i], visits[i], comments[i]))
    r = requests.get(base_url + "article/list/" + str(1) + '?')
    articles_num = re.findall(r'原创</a></dt>\n.*?<dd><a href=.*?<span class="count">(.*?)</span></a></dd>', r.content.decode())
    fan_num = re.findall(r'<dt>粉丝</dt>\n.*?<dd><span class="count" id="fan">(.*?)</span></dd>', r.content.decode())
    like_num = re.findall(r'<dt>喜欢</dt>\n.*?<dd><span class="count">(.*?)</span></dd>', r.content.decode())
    comment_total = re.findall(r'<dt>评论</dt>\n.*?<dd><span class="count">(.*?)</span></dd>', r.content.decode())
    visitd_total = re.findall(r'<dt>访问：</dt>\n.*?<dd title="(.*?)">', r.content.decode())
    score_total = re.findall(r'<dt>积分：</dt>\n.*?<dd title="(.*?)">', r.content.decode())
    rank = re.findall(r'<dl title="(.*?)">\n.*?<dt>排名：</dt>', r.content.decode())
    data2save.append((articles_num[0], fan_num[0], like_num[0], comment_total[0], visitd_total[0], score_total[0], rank[0]))
    WriteToExcel(data2save)
    
    

def WriteToExcel(_data2save):
    #如果excel表格不存在，则创建
    if not os.path.exists(record_file):
        workbook = xlwt.Workbook()
        sheet_visit = workbook.add_sheet("visited")
        sheet_comment = workbook.add_sheet("commented")
        sheet_total = workbook.add_sheet("total")
        #写好第一列的说明
        sheet_visit.write(0, 0, "文章名字")
        sheet_comment.write(0, 0, "文章名字")
        sheet_total.write(0, 0, "项目")
        sheet_total.write(1, 0, "原创")
        sheet_total.write(2, 0, "粉丝")
        sheet_total.write(3, 0, "喜欢")
        sheet_total.write(4, 0, "评论")
        sheet_total.write(5, 0, "访问")
        sheet_total.write(6, 0, "积分")
        sheet_total.write(7, 0, "排名")
        workbook.save(record_file)

    #写入数字格式
    style = xlwt.XFStyle()
    style.num_format_str = '0'

    # 打开excel， 为xlrd
    rexcel = xlrd.open_workbook(record_file)
    # 拷贝原来的rexcel， 变成xlwt
    wexcel = copy(rexcel)
    # 得到工作表
    sheet_visit = wexcel.get_sheet(0)
    sheet_comment = wexcel.get_sheet(1)
    sheet_total = wexcel.get_sheet(2)
    # 得到列数、行数
    read_times = rexcel.sheets()[0].ncols - 1
    article_num = rexcel.sheets()[0].nrows - 1

    #得到当前日期
    now_time = int(str(datetime.datetime.now().strftime('%Y%m%d')))

    #避免一天内重复统计
    if rexcel.sheets()[0].cell_value(0, read_times) != now_time:
        read_times += 1
    sheet_visit.write(0, read_times, now_time, style)
    sheet_comment.write(0, read_times, now_time, style)
    sheet_total.write(0, read_times, now_time, style)

    #写入文章数据
    for i in range(len(_data2save) - 1):
        #如果文章名已经存在，则放到对应行，否则在最后一行增加文件名及对应信息
        current_article_index = article_num + 1
        for j in range(rexcel.sheets()[0].nrows - 1):
            if rexcel.sheets()[0].cell_value(j + 1, 0) == _data2save[i][0]:
                current_article_index = j + 1
        if current_article_index > article_num:
            article_num += 1
            sheet_visit.write(current_article_index, 0, _data2save[i][0])
            sheet_comment.write(current_article_index, 0, _data2save[i][0])
        #保存阅读量及评论量
        sheet_visit.write(current_article_index, read_times, int(_data2save[i][1]), style)
        sheet_comment.write(current_article_index, read_times, int(_data2save[i][2]), style)

    #写入统计信息
    for i in range(rexcel.sheets()[2].nrows - 1):
        sheet_total.write(i + 1, read_times, int(_data2save[len(_data2save)-1][i]), style)

    #保存表格
    wexcel.save(record_file)


if __name__ == "__main__":
    main()
