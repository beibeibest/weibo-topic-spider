import time
import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import os
import requests
import json
from selenium.webdriver.common.by import By
import excelSave as save

# 用来控制页面滚动
def Transfer_Clicks(browser):
    time.sleep(5)
    try:
        browser.execute_script("window.scrollBy(0,document.body.scrollHeight)", "")
    except:
        pass
    return "Transfer successfully \n"

#插入数据
def insert_data(elems,path,yuedu,taolun,num):
    for elem in elems:
        workbook = xlrd.open_workbook(path)  # 打开工作簿
        sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
        worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
        rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数       
        rid = rows_old
        #用户名
        weibo_username = elem.find_elements("css selector","h3.m-text-cut")[0].text

        #微博内容
        #点击“全文”，获取完整的微博文字内容
        weibo_content = get_all_text(elem)

        #获取分享数，评论数和点赞数               
        shares = elem.find_elements("css selector","i.m-font.m-font-forward + h4")[0].text
        if shares == '转发':
            shares = '0'
        comments = elem.find_elements("css selector",'i.m-font.m-font-comment + h4')[0].text
        if comments == '评论':
            comments = '0'
        likes = elem.find_elements("css selector",'i.m-icon.m-icon-like + h4')[0].text
        if likes == '赞':
            likes = '0'

        #发布时间
        weibo_time = elem.find_elements("css selector",'span.time')[0].text
        value1 = [[rid, weibo_username,weibo_content, shares,comments,likes,weibo_time,yuedu,taolun],]
        print("当前插入第%d条数据" % rid)
        save.write_excel_xls_append_norepeat(book_name_xls, value1)

#获取“全文”内容   
def get_all_text(elem):
    try:
        #判断是否有“全文内容”，若有则将内容存储在weibo_content中
        href = elem.find_element_("link text",'全文').get_attribute('href')
        driver.execute_script('window.open("{}")'.format(href))
        driver.switch_to.window(driver.window_handles[1])
        weibo_content = driver.find_element("class name",'weibo-text').text
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
    except:
        weibo_content = elem.find_elements("css selector","div.weibo-text")\
                        [0].text
    return weibo_content

#获取当前页面的数据
def get_current_weibo_data(book_name_xls,yuedu,taolun,maxWeibo,num):
    #开始爬取数据
        before = 0 
        after = 0
        n = 0 
        timeToSleep = 100
        while True:
            before = after
            Transfer_Clicks(driver)
            time.sleep(2)
            elems = driver.find_elements("css selector","div.card.m-panel.card9")
            print("当前包含微博最大数量：%d,n当前的值为：%d, n值到5说明已无法解析出新的微博" % (len(elems),n))
            after = len(elems)
            if after > before:
                n = 0
            if after == before:        
                n = n + 1
            if n == 5:
                print("当前关键词最大微博数为：%d" % after)
                insert_data(elems,book_name_xls,yuedu,taolun,num)
                break
            if len(elems)>maxWeibo:
                print("当前微博数以达到%d条"%maxWeibo)
                insert_data(elems,book_name_xls,yuedu,taolun,num)
                break

#爬虫运行 
def spider(book_name_xls,sheet_name_xls,maxWeibo,num):
    
    #创建文件
    if os.path.exists(book_name_xls):
        print("文件已存在")
    else:
        print("文件不存在，重新创建")
        value_title = [["rid", "用户名称", "微博内容", "微博转发量","微博评论量","微博点赞","发布时间","话题阅读数","话题讨论数"],]
        save.write_excel_xls(book_name_xls, sheet_name_xls, value_title)

    
    #输入你需要爬取的超话网页
    driver.get("https://m.weibo.cn/p/index?containerid=100808327ab6bd2f01c79bf2a70afb70a5b246&luicode=10000011&lfid=100103type%3D1%26q%3D%E6%95%B0%E5%AD%97%E8%97%8F%E5%93%81")

    time.sleep(2)
    yuedu_taolun = driver.find_element("xpath","//*[@id='app']/div[1]/div[1]/div[1]/div[4]/div/div/div/a/div[2]/h4[1]").text
    yuedu = yuedu_taolun.split("　")[0]
    taolun = yuedu_taolun.split("　")[1]
    time.sleep(2)
    shishi_element = driver.find_element("xpath","//*[@id='app']/div[1]/div[1]/div[2]/div[2]/div[1]/div/div/div/ul/li[1]/span")

    get_current_weibo_data(book_name_xls,yuedu,taolun,maxWeibo,num) #爬取实时
    time.sleep(2)


    
if __name__ == '__main__':
    driver = webdriver.Chrome("/usr/local/bin/chromedriver")#你的chromedriver的地址
    driver.implicitly_wait(2)#隐式等待2秒
    book_name_xls = "/Users/niubei/Desktop/计算机/test.xls" #填写你想存放excel的路径，没有文件会自动创建
    sheet_name_xls = '微博数据' #sheet表名
    maxWeibo = 10 #设置最多多少条微博
    num = 1
    spider(book_name_xls,sheet_name_xls,maxWeibo,num)
