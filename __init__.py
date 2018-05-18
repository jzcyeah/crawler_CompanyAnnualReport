import xlrd
import requests
from bs4 import BeautifulSoup


# 主要爬取股吧东方财富、东方财富、金融界

def getTitle(num, name, url):
    # proxy = {'http': '182.88.129.4:8123'}
    # star_html = requests.get(url, headers=headers, proxies=proxy)
    star_html = requests.get(url, headers=headers)
    soup = BeautifulSoup(star_html.text, 'lxml')
    res = soup.find('div', class_='container_s')
    res1 = res.find('div', class_='result c-container ')
    link = res1.find('a').get('href')
    totalBaiduLinke = soup.findAll('h3', class_="t")
    for i in totalBaiduLinke:
        title = i.get_text()
        if (title.find("年度报告") != -1 and title.find("2017") != -1
            and title.find("摘要") == -1 and title.find("审计") == -1 and title.find("业绩") == -1 and title.find(
            "决算") == -1 and title.find("股东") == -1 and title.find("2016") == -1 and title.find(
            "2015") == -1 and title.find("半年度") == -1 and title.find("公告") == -1 and title.find("董事会") == -1):
            url = i.find('a').get('href')
            realLink = requests.get(url, headers=headers).url  # 获取真实url
            if (realLink.find("http://data.eastmoney.com/notice") != -1):
                getEastPdf(num, name, url)  # 获取新网址的pdf地址
                break
            if (realLink.find("http://guba.eastmoney.com") != -1):
                getGubaEastPdf(num, name, url)
                break
            if (realLink.find("stock.jrj") != -1):
                getJrj(num, name, url)
                break


def getGubaEastPdf(num, name, link):
    star_html = requests.get(link, headers=headers)
    soup = BeautifulSoup(star_html.text, 'lxml')
    finalLink = soup.find('span', class_='zwtitlepdf').find('a').get('href')
    download(num, name, finalLink)


def getEastPdf(num, name, link):
    star_html = requests.get(link, headers=headers)
    soup = BeautifulSoup(star_html.text, 'lxml')
    res1 = str(soup)
    a = res1[res1.find('http://pdf'):res1.find('http://pdf') + 300]
    finalLink = a.split(" ")[0]
    download(num, name, finalLink)


def getJrj(num, name, link):
    star_html = requests.get(link, headers=headers)
    soup = BeautifulSoup(star_html.text, 'lxml')
    finalLink = soup.find('div', class_='warp').find('div', class_="main").find('span', class_='down').find('a').get(
        'href')
    download(num, name, finalLink)


def download(num, name, finalLink):
    if num < 10:
        path = "D:\\download\\gift3\\" + "0" + str(num) + name + ".pdf"
    else:
        path = "D:\\download\\gift3\\" + str(num) + name + ".pdf"
    r = requests.get(finalLink, headers=headers)
    with open(path, "wb") as f:
        f.write(r.content)
    f.close()



def read_excel():
    list = []
    workbook = xlrd.open_workbook('needToDo.xlsx')
    sheets = workbook.sheet_names()
    worksheet = workbook.sheet_by_name(sheets[0])
    for i in range(0, worksheet.nrows)[1:]:
        list.append(worksheet.cell_value(i, 6))
    return list


if __name__ == '__main__':
    headers = {
        'User-Agent': "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1"}
    totalCompany = read_excel()
    url1 = "https://www.baidu.com/s?ie=UTF-8&wd="
    # url1="https: // www.baidu.com / baidu?wd ="
    num = 0
    for i in totalCompany:
        # if num == 10:
        #     break
        num = num + 1
        url = url1 + i + "2017年报"
        # url = url1 + i + "2015年报"+"&tn=monline_dg&ie=utf-8"
        try:
            getTitle(num, i, url)
        except:
            continue
