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

#判断页面是否加载出来
def isPresent():
    temp =1
    try: 
        driver.find_elements_(By.CSS.SELECTOR,'div.line-around.layout-box.mod-pagination > a:nth-child(2) > div > select > option')
    except:
        temp =0
    return temp

#插入数据
def insert_data(elems,path,name,yuedu,taolun,num,save_pic):
    for elem in elems:
        workbook = xlrd.open_workbook(path)  # 打开工作簿
        sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
        worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
        rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数       
        rid = rows_old
        #用户名
        weibo_username = elem.find_elements("xpath","//*[@id='app']/div[1]/div[1]/div[19]/div/div/div[3]/div/div/div/header/div/div/a/h3")[0].text
        
        weibo_userlevel = "普通用户"
        '''
        #微博等级
        try: 
            weibo_userlevel_color_class = elem.find_elements(By.cssSelector("img.m-icon"))[0].get_attribute("class").replace("m-icon ","")
            if weibo_userlevel_color_class == "m-icon-yellowv":
                weibo_userlevel = "黄v"
            if weibo_userlevel_color_class == "m-icon-bluev":
                weibo_userlevel = "蓝v"
            if weibo_userlevel_color_class == "m-icon-goldv-static":
                weibo_userlevel = "金v"
            if weibo_userlevel_color_class == "m-icon-club":
                weibo_userlevel = "微博达人"     
        except:
            weibo_userlevel = "普通用户"
        '''
        #微博内容
        #点击“全文”，获取完整的微博文字内容
        weibo_content = get_all_text(elem)
        #获取微博图片
        if save_pic:
            num = get_pic(elem,num)
        #获取分享数，评论数和点赞数               
        shares = elem.find_elements("xpath","//*[@id='app']/div[1]/div[1]/div[21]/div/div/div[10]/div/div/div/footer/div[1]/h4")[0].text
        if shares == '转发':
            shares = '0'
        comments = elem.find_elements("xpath","//*[@id='app']/div[1]/div[1]/div[19]/div/div/div[3]/div/div/div/footer/div[2]/h4")[0].text
        if comments == '评论':
            comments = '0'
        likes = elem.find_elements("xpath","//*[@id='app']/div[1]/div[1]/div[21]/div/div/div[4]/div/div/div/footer/div[3]/h4")[0].text
        if likes == '赞':
            likes = '0'

        #发布时间
        weibo_time = elem.find_elements("xpath","//*[@id='app']/div[1]/div[1]/div[22]/div/div/div[2]/div/div/div/header/div/div/h4/span[1]")[0].text
        '''
        print("用户名："+ weibo_username + "|"
              "微博等级："+ weibo_userlevel + "|"
              "微博内容："+ weibo_content + "|"
              "转发："+ shares + "|"
              "评论数："+ comments + "|"
              "点赞数："+ likes + "|"
              "发布时间："+ weibo_time + "|"
              "话题名称" + name + "|" 
              "话题讨论数" + yuedu + "|"
              "话题阅读数" + taolun)
        '''
        value1 = [[rid, weibo_username, weibo_userlevel,weibo_content, shares,comments,likes,weibo_time,keyword,name,yuedu,taolun],]
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
        weibo_content = elem.find_elements("xpath","//*[@id='app']/div[1]/div[1]/div[12]/div/div/div/div/div/div/article/div[2]/div[1]")\
                        [0].text
    return weibo_content
'''
def get_pic(elem,num):
    try:
        #获取该条微博中的图片元素,之后遍历每个图片元素，获取图片链接并下载图片
        #如果是多张图片
        if elem.find_elements\
           (By.CSS_SELECTOR,'div > div > article > div > div:nth-child(2) > div > ul > li') != [] :
            pic_links = elem.find_elements\
           (By.CSS_SELECTOR,'div > div > article > div > div:nth-child(2) > div > ul > li')
            for pic_link in pic_links:
                pic_link = pic_link.find_element\
                           (By.CSS_SELECTOR,'div > img').get_attribute('src')
                response = requests.get(pic_link)
                pic = response.content
                with open(pic_addr + str(num) + '.jpg', 'wb') as file:
                    file.write(pic)
                    num += 1
        #如果图片只有一张
        else:
            pic_link = elem.find_element\
                       (By.CSS_SELECTOR,'div > div > article > div > div:nth-child(2) > div > div > img').\
                       get_attribute('src')
            response = requests.get(pic_link)
            pic = response.content
            with open(pic_addr + str(num) + '.jpg', 'wb') as file:
                file.write(pic)
                num += 1
    except Exception as e:
        print(e)
    return num
    '''
#获取当前页面的数据
def get_current_weibo_data(elems,book_name_xls,name,yuedu,taolun,maxWeibo,num):
    #开始爬取数据
        before = 0 
        after = 0
        n = 0 
        timeToSleep = 100
        while True:
            before = after
            Transfer_Clicks(driver)
            time.sleep(3)
            elems = driver.find_elements(By.CSS_SELECTOR,'div.card.m-panel.card9')
            print("当前包含微博最大数量：%d,n当前的值为：%d, n值到5说明已无法解析出新的微博" % (len(elems),n))
            after = len(elems)
            if after > before:
                n = 0
            if after == before:        
                n = n + 1
            if n == 5:
                print("当前关键词最大微博数为：%d" % after)
                insert_data(elems,book_name_xls,name,yuedu,taolun,num,save_pic)
                break
            if len(elems)>maxWeibo:
                print("当前微博数以达到%d条"%maxWeibo)
                insert_data(elems,book_name_xls,name,yuedu,taolun,num,save_pic)
                break
            '''
            if after > timeToSleep:
                print("抓取到%d多条，插入当前新抓取数据并休眠5秒" % timeToSleep)
                timeToSleep = timeToSleep + 100
                insert_data(elems,book_name_xls,name,yuedu,taolun,num,save_pic) 
                time.sleep(5) 
            '''


#爬虫运行 
def spider(username,password,book_name_xls,sheet_name_xls,keyword,maxWeibo,num,save_pic):
    
    #创建文件
    if os.path.exists(book_name_xls):
        print("文件已存在")
    else:
        print("文件不存在，重新创建")
        value_title = [["rid", "用户名称", "微博等级", "微博内容", "微博转发量","微博评论量","微博点赞","发布时间","搜索关键词","话题名称","话题讨论数","话题阅读数"],]
        save.write_excel_xls(book_name_xls, sheet_name_xls, value_title)
    
    #加载驱动，使用浏览器打开指定网址  
    driver.set_window_size(496, 790)
    driver.get('https://m.weibo.cn')
    
    driver.get("https://m.weibo.cn/login?backURL=https%3A%2F%2Fm.weibo.cn%2F")
    print("开始自动登陆，若出现验证码手动验证")
    time.sleep(3)

    elem=driver.find_element("xpath","//*[@id='app']/div[1]/div/div/form/div/input");
    elem.send_keys(username)
    elem=driver.find_element("xpath","//*[@id='app']/div[1]/div/div/a");
    elem.click() 
    print("暂停60秒，用于验证码验证")
    time.sleep(60)
    
    
    # 添加cookie
    #将提前从chrome控制台中复制来的cookie保存在txt中，转化成name, value形式传给selenium的driver
    #实现自动登录
    #如果txt中的cookie是用selenium保存的，则可以直接使用, 无需转化
    '''
    driver.delete_all_cookies()
    with open(r'./weibocookie.txt') as file:
        cookies = json.loads(file.read())
    for name, value in cookies.items():
        print(name, value)
        driver.add_cookie({'name': name, 'value': value})
    driver.refresh()
    '''
    
        
        
    while 1:  # 循环条件为1必定成立
        result = 1
        # 解决输入验证码无法跳转的问题
        driver.get('https://m.weibo.cn/') 
        print ('判断页面1成功 0失败  结果是=%d' % result )
        if result == 1:
            elems = driver.find_elements(By.CSS_SELECTOR,'div.line-around.layout-box.mod-pagination > a:nth-child(2) > div > select > option')
            #return elems #如果封装函数，返回页面
            break
        else:
            print ('页面还没加载出来呢')
            time.sleep(20)

    time.sleep(2)

    #搜索关键词
    driver.get("https://m.weibo.cn/search?containerid=231583")
    elem = driver.find_element("xpath","//*[@id='app']/div[1]/div[1]/div[1]/div/div/div[2]/form/input");
    time.sleep(2)
    elem.send_keys(keyword)
    elem.send_keys(Keys.ENTER)
    time.sleep(5)
    
    # elem = driver.find_element_by_xpath("//*[@class='box-left m-box-col m-box-center-a']")
    # 修改为点击超话图标进入超话，减少错误
    time.sleep(10)
    elem = driver.find_element("xpath","//*[@id='app']/div[1]/div[1]/div[3]/div/div/div[3]/div/div/div[2]/div/img")
    elem.click()
    print("超话链接获取完毕，休眠2秒")
    time.sleep(2)
    yuedu_taolun = driver.find_element("xpath","//*[@id='app']/div[1]/div[1]/div[1]/div[4]/div/div/div/a/div[2]/h4[1]").text
    yuedu = yuedu_taolun.split("　")[0]
    taolun = yuedu_taolun.split("　")[1]
    time.sleep(2)
    name = keyword
    shishi_element = driver.find_element("xpath","//*[@id='app']/div[1]/div[1]/div[2]/div[2]/div[1]/div/div/div/ul/li[1]/span")

    get_current_weibo_data(elems,book_name_xls,name,yuedu,taolun,maxWeibo,num) #爬取实时
    time.sleep(2)

    
if __name__ == '__main__':
    username = "18726085281" #你的微博登录名
    password = "" #密码
    driver = webdriver.Chrome("/usr/local/bin/chromedriver")#你的chromedriver的地址
    driver.implicitly_wait(2)#隐式等待2秒
    book_name_xls = "/Users/niubei/Desktop/计算机/test.xls" #填写你想存放excel的路径，没有文件会自动创建
    sheet_name_xls = '微博数据' #sheet表名
    maxWeibo = 10 #设置最多多少条微博
    keywords = ["数字藏品",] # 此处可以设置多个超话关键词
    num = 1
    save_pic = False  #设置是否同时爬取微博图片，默认不爬取
    pic_addr = 'img/' #设置自己想要放置图片的路径

    for keyword in keywords:
        spider(username,password,book_name_xls,sheet_name_xls,keyword,maxWeibo,num,save_pic)
