import requests
from bs4 import BeautifulSoup
import re
import xlsxwriter
import datetime
import os


# 网址
url = "https://dementiadiaries.org/?s=forgot"


# 发送 GET 请求
response = requests.get(url)


# 解析网页内容
soup = BeautifulSoup(response.content, 'html.parser')


# 找到所有链接
links = soup.find_all('a')


# 创建一个列表来存储链接
link_list = []


# 循环遍历链接
for link in links:
    # 获取链接的 href 属性
    href = link.get('href')
    
    # 检查链接是否包含关键词
    if href and 'forgot' in href:
        # 将链接添加到列表中
        link_list.append(href)


# 获取当前时间
now = datetime.datetime.now()


# 创建文件名
filename = f"links of keyword_{now.strftime('%Y%m%d_%H%M%S')}.xlsx"


# 指定文件保存路径
save_path = os.path.join("E:\\IME-S1\\SD5913\\asm4-taro\\data_collect")


# 如果保存路径不存在，则创建
if not os.path.exists(save_path):
    os.makedirs(save_path)


# 创建一个 xlsx 文件
workbook = xlsxwriter.Workbook(os.path.join(save_path, filename))
worksheet = workbook.add_worksheet()


# 写入链接到 xlsx 文件
for i, link in enumerate(link_list):
    worksheet.write(i, 0, link)


# 关闭 xlsx 文件
workbook.close()