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

def citation(link, soup, value):
    try:
        author = ""
        title = soup.find(class_ = "publicationContentTitle").text
        authors = soup.find(class_ = "authors")
        for i in authors.find_all(class_ = "contribDegrees"):
            author += i.find("a").text + ","
        cite = author + title + ".Politics. " + "https://journals.sagepub.com" + link
        cite = cite.replace("\n", "")
    except:
        cite = ""
    finally:
        value.append(cite)

def abstract(soup, value):
    try:
        str = soup.find(class_ = "abstractSection abstractInFull").text
    except:
        str = ""
    finally:
        value.append(str)
    
def keyword(soup, value):
    try:
        str = soup.find(class_ = "abstractKeywords").text
        str = str.replace("Keywords ", "")
    except:
        str = ""
    finally:
        value.append(str)

def bs(weblist, values):
    #https://journals.sagepub.com/doi/full/10.1177/02633957211035096
    
    value = []
    print(weblist)

    ua = UserAgent(verify_ssl = False)
    user_agent = ua.random
    response = requests.get('https://journals.sagepub.com' + weblist, headers= {'user-agent': user_agent})
    soup = BeautifulSoup(response.text, "lxml")

    citation(weblist, soup, value)
    abstract(soup, value)
    keyword(soup, value)

    values.append(value)
    time.sleep(5)


def spider():
    links = ['/toc/pola/41/4', '/toc/pola/41/3', '/toc/pola/41/2', '/toc/pola/41/1']
    for link in links:
        print(link)
        values = []

        ua = UserAgent(verify_ssl = False)
        user_agent = ua.random
        response = requests.get('https://journals.sagepub.com' + link, headers= {'user-agent': user_agent})
        time.sleep(5)

        soup  = BeautifulSoup(response.text, 'lxml')
        for list in soup.find_all("a", attrs={"data-item-name": "click-article-title"}):
            ahref = list['href']
            bs(ahref, values)
        
        write_excel_xls_append("journal.xls",values)

spider()


