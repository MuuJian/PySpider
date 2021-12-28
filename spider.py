from os import name
import requests
import time
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
import xlrd
import xlwt
from xlutils.copy import copy


def write_excel_xls_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i+rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    print("xls格式表格【追加】写入数据成功！")

def citation(text, value):
    str = ""
    soup = BeautifulSoup(text, "lxml")
    citations = soup.find_all(class_ = 'card')
    if len(citations) == 0:
        value.append("")
        return
    for list in citations:
        for a in list.find_all(class_ = 'issue-item__title'):
            href = a['href']
            href = href.replace("/doi/abs/10.1111/psj.","")
            str += (a.text + "(https://doi.org/10.1111/psj." + href + ")")
    value.append(str)

def abstract(text, value):
    str = ""
    soup = BeautifulSoup(text, "lxml")
    Abstracts = soup.find(class_ ='article-section__content en main')
    if Abstracts == None:
        value.append("")
        return
    str = Abstracts.text
    value.append(str)
    
def keyword(text, value):
    str = ""
    soup = BeautifulSoup(text, "lxml")
    Keywords = soup.find_all(class_ = 'rlist rlist--inline')
    if len(Keywords) == 0:
        value.append("")
        return
   
    for li in Keywords:
        str += li.text
    value.append(str)

def bs(weblist, values):
    #https://onlinelibrary.wiley.com/action/showCitFormats?doi=10.1111%2Ftwec.13199
    #https://onlinelibrary.wiley.com/doi/10.1111/twec.13199
    #https://onlinelibrary.wiley.com/action/ajaxShowPubInfo?widgetId=5cf4c79f-0ae9-4dc5-96ce-77f62de7ada9&ajax=true&doi=10.1111/twec.13199

    value = []
    print(weblist)
    ua = UserAgent(verify_ssl = False)
    user_agent = ua.random

    response = requests.get('https://onlinelibrary.wiley.com/action/showCitFormats?doi=10.1111%2Fpsj.' + weblist, headers= {'user-agent': user_agent})
    citation(response.text, value)

    response = requests.get('https://onlinelibrary.wiley.com/doi/10.1111/psj.' + weblist, headers= {'user-agent': user_agent})
    abstract(response.text,value)

    response = requests.get('https://onlinelibrary.wiley.com/action/ajaxShowPubInfo?widgetId=5cf4c79f-0ae9-4dc5-96ce-77f62de7ada9&ajax=true&doi=10.1111/psj.' + weblist, headers= {'user-agent': user_agent})
    keyword(response.text, value)

    values.append(value)
    time.sleep(5)

    

def start(page):
    #page = 'https://onlinelibrary.wiley.com/loi/14679701/year/2021'
    list_ = []
    ua = UserAgent(verify_ssl = False)
    user_agent = ua.random
    response = requests.get(page, headers= {'user-agent': user_agent})
    time.sleep(5)

    soup = BeautifulSoup(response.text, 'lxml')
    for li in soup.find_all(class_ = "parent-item"):
        for a  in li.find_all(name = "a"):
            list_.append(a["href"])
    
    for page_ in list_:
        print(page_)
        spider(page_)

def spider1():
    links = ['/toc/15410072/2021/49/3']
    for link in links:
        print(link)
        values = []
        ua = UserAgent(verify_ssl = False)
        user_agent = ua.random
        response = requests.get('https://onlinelibrary.wiley.com' + link, headers= {'user-agent': user_agent})
        time.sleep(5)
   
        weblist = []
        soup  = BeautifulSoup(response.text, 'lxml')
        for list in soup.find_all(class_ = "issue-item__title visitable"):
            if(list.text != "\nIssue Information" and list.text != "\nCover Image"):
                str = list['href']
                str = str.replace('/doi/10.1111/twec.','')
                weblist.append(str)

        for list in weblist:
            bs(list, values)
        
        write_excel_xls_append("journal.xls",values)


def spider(link):
    values = []
    ua = UserAgent(verify_ssl = False)
    user_agent = ua.random
    response = requests.get('https://onlinelibrary.wiley.com' + link, headers= {'user-agent': user_agent})
    time.sleep(5)
    #https://onlinelibrary.wiley.com/action/showCitFormats?doi=10.1111%2Ftwec.13199
    #https://onlinelibrary.wiley.com/doi/10.1111/twec.13199
    #/doi/10.1111/twec.13168
    weblist = []
    soup  = BeautifulSoup(response.text, 'lxml')
    for list in soup.find_all(class_ = "issue-item__title visitable"):
        if(list.text != "\nIssue Information" and list.text != "\nCover Image" and list.text[0: 10] != "\nEditorial"):
            str = list['href']
            str = str.replace('/doi/10.1111/psj.','')
            weblist.append(str)

    for list in weblist:
        bs(list, values)
        
    write_excel_xls_append("journal.xls",values)


start('https://onlinelibrary.wiley.com/loi/15410072/year/2021')
