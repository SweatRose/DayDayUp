# -*- coding = utf-8 -*-
# @Time : 2022-10-13 19:09:29
# @updateTime: 2023-12-20 09:42:52
# @Author : Anonymous
# @File : DayDayUp.py
# @Software : PyCharm

from nntplib import ArticleInfo
import requests
from bs4 import BeautifulSoup                       # 网页解析，获取数据
import time                                         # 时间处理
import re                                           # 正则匹配
import xlwt                                         # 进行excel操作
from pyquery import PyQuery as pq
#操作word的库
from docx import Document
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.shared import RGBColor

def main():
    datalist = []                                   # 保存文章标题和内容
    theme_list = {'4006': '申论热点', '4005': '申论范文', '4007': '申论技巧', '3994': '全部申论',
                  '3993': '行测技巧', '3995': '面试技巧', '4011': '面试热点', '4012': '公安基础知识', '3997': '综合指导'}
    theme = input("请选择爬取内容代码（申论热点：4006、申论范文：4005、申论技巧：4007、全部申论：3994、行测技巧：3993、面试技巧：3995、面试热点：4011、公安基础知识：4012、综合指导：3997）：")
    page = input("请输入需要爬取的页数：")
    URL_list = get_pages_url(theme, int(page))              # 获取每个文章的URL
    savepath = r'.\\' + theme_list[theme] + '.xls'
    for i in range(len(URL_list)):
        print("正在获取第{}篇范文".format(i+1))
        # datalist.extend(get_Data(URL_list[i]))
        title, content = get_Data(URL_list[i])  # 获取文章标题和内容
        datalist.append((title, content))  # 将标题和内容作为元组加入列表
        time.sleep(0.5)
        i += 1

        saveData(datalist, theme_list[theme])

def askURL(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36'
    }

    response = requests.get(url=url, headers=headers)
    html = response.content.decode("utf-8")
    bs = BeautifulSoup(html, "html.parser")  # 解析html，提取数据
    return bs

# 获取每一页的全部url
def get_pages_url(theme, page):
    data = []
    for num in range(0, page, 1):
        url = 'http://www.offcn.com/gwy/ziliao/{}/{}.html'.format(theme, num+1)
        print("正在提取第" + str(num+1) + "页")
        bs = askURL(url)
        time.sleep(0.5)
        list01 = str(bs.find_all('ul', class_="lh_newBobotm02"))
        pat_link = r'<a href="(.*?)" target="_blank" title=.*?'
        link = re.findall(pat_link, list01, re.S)
        link_list = str(link).strip("[").strip("]").replace("'", "").split(",")
        data.extend(link_list)
    return data

def get_Data(url):
    articleData = []
    bs = askURL(url)
    title_MsgStr = bs.select("h1")
    title = ""
    for title_str in title_MsgStr:
        title = title + title_str.text
    articleData.append(title)

    content = ""
    list01 = str(bs.find_all('div', class_="offcn_shocont"))
    content = re.sub(r'<.*?>', '', list01)  # 去除HTML标签
    content = re.sub(r'\[.*?\]', '', content)  # 去除[]标签
    articleData.append(str(content))
    return articleData

def saveData(datalist, theme_name):  # 修改saveData函数的参数列表
    for i, data in enumerate(datalist):
        doc = Document()
        title_paragraph = doc.add_paragraph()
        title_run = title_paragraph.add_run(data[0].replace('(进入阅读模式)', ''))
        title_run = title_paragraph.add_run(data[0])
        title_run.font.name = '微软雅黑'
        title_run.font.size = Pt(16)

        content_paragraph = doc.add_paragraph()
        content_run = content_paragraph.add_run(data[1])
        content_run.font.name = '仿宋'
        content_run.font.size = Pt(12)

        doc.save('./archive/' + data[0] +'.docx')

if __name__ == '__main__':
    main()
    print("爬取完毕！！！")