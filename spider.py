from os import name
import base64
import requests
import time
from fake_useragent import UserAgent
from bs4 import BeautifulSoup, element
from selenium.webdriver.chrome import options
import xlrd
import xlwt
from xlutils.copy import copy
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

ch_options = Options()


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
    #https://academic.oup.com/ser/article-abstract/19/1/7/5299221?redirectedFrom=fulltext
    #https://academic.oup.com/ser/article/19/1/7/5299221
    #/ser/article/19/1/1/6307095
    value = []
    print(weblist)

    ua = UserAgent(verify_ssl = False)
    user_agent = ua.random
    ch_options.add_argument(user_agent)
    driver = webdriver.Chrome('/Users/mumu/Documents/GitHub/PySpider/chromedriver', options = ch_options)
    driver.get('https://academic.oup.com' + weblist)
    img64 = driver.find_element_by_xpath('/html/body/div/div/img').get_attribute('src')
    img64 = img64.replace('data:image/jpg;base64,', '')
    img64 = base64.b64decode(img64)
    with open("img.png", "wb") as img:
        img.write(img64)


    capinput = driver.find_element_by_id('txtCaptchaInputId')
    capinput.send_keys(str)
    btnSubmit = driver.find_element_by_id('btnSubmit')
    btnSubmit.click()
    soup = BeautifulSoup(driver.page_source, "lxml")

    print(soup.text)
    citation(weblist, soup, value)
    abstract(soup, value)
    keyword(soup, value)

    values.append(value)
    time.sleep(5)


def spider():
    #https://academic.oup.com/ser/issue/19/1
    links = ['/issue/19/1', '/issue/19/2', '/issue/19/3', '/issue/19/4']
    

    for link in links:
        print(link)
        values = []

        ua = UserAgent(verify_ssl = False)
        user_agent = ua.random
        ch_options.add_argument(user_agent)
        driver = webdriver.Chrome('/Users/mumu/Documents/GitHub/PySpider/chromedriver', options = ch_options)
        driver.get('https://academic.oup.com/ser' + link)
        time.sleep(5)
        soup  = BeautifulSoup(driver.page_source, 'lxml')
        driver.close()
        for list in soup.find_all("a", attrs={"class": "at-articleLink"}):
            ahref = list['href']
            bs(ahref, values)
        write_excel_xls_append("journal.xls",values)

ocr()
bs("/ser/article/19/1/1/6307095", [])