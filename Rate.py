# -*- coding: utf-8 -*-
import requests
import json
import regex as re
import xlwt, xlrd
from xlutils.copy import copy
from multiprocessing import Pool


def get_text(text):
    f = open(text, 'r', encoding=('unicode_escape'))
    text = f.read().lower()
    for i in '!@#$%^&*()_¯+-;:`~\'"<>=./?,[]{}0123456789':
        text = text.replace(i, ' ')
    return text.split()


def clean_space(text):  ##去除空格
    match_regex = re.compile(u'[\u4e00-\u9fa5。\.,，:：《》、\(\)（）]{1} +(?<![a-zA-Z])|\d+ +| +\d+|[a-z A-Z]+')
    should_replace_list = match_regex.findall(text)
    order_replace_list = sorted(should_replace_list, key=lambda i: len(i), reverse=True)
    for i in order_replace_list:
        if i == u' ':
            continue
        new_i = i.strip()
        text = text.replace(i, new_i)
    return text


def sort():  ##排序
    clean_space('part1.txt')
    ls = get_text('part1.txt')
    counts = {}
    # print(len(ls))
    for i in ls:
        counts[i] = counts.get(i, 0) + 1
    items = list(counts.items())  # items()将字典中的键值对打包返回一个元组
    items.sort(key=lambda x: x[1], reverse=True)  # 对字典中的值进行从大到小的排序其中的lambda表示的是将列表中的元素的第二位传递给x,key = x ,对key进行从大到小的排序
    ##print(type(items))
    return items


def translate(word):
    url = "http://dict-co.iciba.com/api/dictionary.php?w=" + word + "&type=json" + "&key=19EE6AA32AC373F7938024DDC2C8EB23"
    response = requests.get(url=url)
    obj = json.loads(response.text)
    return obj["symbols"][0]["parts"]


def set_style(name, height, bold=False):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name  # 'Times New Roman'
    font.bold = bold
    font.color_index = 4
    font.height = height

    # borders= xlwt.Borders()
    # borders.left= 6
    # borders.right= 6
    # borders.top= 6
    # borders.bottom= 6

    style.font = font
    # style.borders = borders

    return style


def write_excel(word, rate, translate, i):
    # index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook("词汇.xls")  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    new_worksheet.write(i, 0, word, set_style('Times New Roman', 220, True))  # 追加写入数据，注意是从i+rows_old行开始写入
    new_worksheet.write(i, 1, rate, set_style('Times New Roman', 220, True))
    new_worksheet.write(i, 2, translate, set_style('Times New Roman', 220, True))
    new_workbook.save("词汇.xls")

    # sheet = workbook.add_sheet('四级', cell_overwrite_ok=False)  ##创建工作表
    # row0 = ["单词", "频率", "释义"]
    #
    # sheet.write(i, 0, word, set_style('Times New Roman', 220, True))  ##写入单词，write(行，列，数据，格式)
    # sheet.write(i, 1, rate, set_style('', 220, True))
    # sheet.write(i, 2, translate, set_style('', 220, True))
    # workbook = xlwt.Workbook(encoding='utf-8', style_compression=0)  ## 创建Excel对象
    # workbook.save("词汇.xls")


def main():
    for i in range(1, len(word_list)):
        try:
            word = word_list[i]
            meaning = translate(word[0])
            print("正在处理："+word[0])
            write_excel(word[0], word[1], meaning[0]['means'], i)
        except:
            pass


if __name__ == '__main__':
    # workbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
    # sheet = workbook.add_sheet('四级', cell_overwrite_ok=False)
    # workbook.save("词汇.xls")
    word_list = sort()
    print("----start----")
    po = Pool(3)
    po.apply_async(main(),())
    po.close()  # 关闭进程池，关闭后po不再接收新的请求
    po.join()  # 等待po中所有子进程执行完成，必须放在close语句之后
    print("-----end-----")
