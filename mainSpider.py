# This program depends on:
# (1)  python3.x
# (2)  Chrome80.0.3987.87(Latest Version)
# (3)  ChromeDriver 80.0.3987.16
# (4)  selenium
# (5)  re
# (6)  csv

import time, datetime
import ast
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from PIL import Image, ImageEnhance, ImageTk
import PIL
import matplotlib.pyplot as plt
import os
import re
import csv
import random
import tkinter as tk
from tkinter import *
from tkinter import ttk
import tkinter.messagebox

class Spider():
    def __init__(self):
        '''
        Search information from sina weibo according to keywords.
        Default browser: Chrome
        :param keywords: A list including the keywords you want to search.
        '''
        self.browser = None
        self.username = ''
        self.keywords = None
        self.login_flag = False
        self.codePath = None

    def _get_cdoe_img(self):
        codeImg = self.browser.find_elements_by_css_selector("#pl_login_form > div > div:nth-child(3) > div.info_list.verify.clearfix > a > img")
        time.sleep(1)
        # img location
        location = codeImg[0].location
        size = codeImg[0].size
        return location, size

    def _showCode(self, location, size):
        print("验证码的位置为：", location)
        print("验证码的尺寸为：", size)
        left = location['x']
        top = location['y']
        bottom = top + size['height']
        right = left + size['width']

        try:
            os.mkdir("./img")
        except:
            pass

        # save full shot
        suffix = str(int(time.time()))
        imgPath = './img/full_snap'+ suffix +'.png'
        self.browser.save_screenshot(imgPath)

        # obtain safe code img
        page_snap_obj = PIL.Image.open(imgPath)
        image_obj = page_snap_obj.crop((1.5*left, 1.5*top, 1.5*right, 1.5*bottom))
        self.codePath = './img/verification_code'+ suffix +'.png'


        # enhance
        image = ImageEnhance.Contrast(image_obj)
        image = image.enhance(4)
        image.save(self.codePath)
        # plt.imshow(image)
        # plt.show()


    def cookie_login(self):
        flag = False        # identify whether need to login by mormal login operation(without cookie)

        try:
            os.mkdir("./cookies_log")
        except:
            pass

        try:
            with open('./cookies_log/log.txt', "r") as f:
                cookie = f.read(-1)
        except:
            flag = True
            return flag

        cookies = ast.literal_eval(cookie)  # str list -> list

        print("通过cookie登录中......")
        for cookie in cookies:
            if 'expiry' in cookie:
                del cookie['expiry']
            self.browser.add_cookie(cookie)
        self.browser.get("https://www.weibo.com/")
        time.sleep(2)

        try:
            wait = WebDriverWait(self.browser, 15)
            id_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#v6_pl_rightmod_myinfo > div > div > div.WB_innerwrap > div > a.name.S_txt1")))
            self.username = id_element.text
            print("恭喜您，"+ id_element.text +"，通过cookie登录成功")
            return flag
        except:
            print("登录异常")
            return True



    def login(self, username, password):
        self.browser = webdriver.Chrome()
        self.browser.get("https://www.weibo.com/")
        self.browser.delete_all_cookies()
        self.browser.maximize_window()

        flag = self.cookie_login()
        if  flag == False:
            # login successfully
            self.login_flag = True
            return

        # explict wait
        wait = WebDriverWait(self.browser, 15)
        try:
            name_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#loginname")))
            passwd_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pl_login_form > div > div:nth-child(3) > div.info_list.password > div > input")))
        except:
            raise Exception("Time out, please check Internet connection")

        # login
        # first, clear the cache recorded by browser
        name_input.clear()
        name_input.send_keys(username)
        time.sleep(1)
        passwd_input.clear()
        passwd_input.send_keys(password)

        # discern verification code
        location, size = self._get_cdoe_img()
        self._showCode(location, size)


    def loginWithCode(self, code):
        # code = input("请输入验证码：")
        code_input = self.browser.find_element_by_css_selector("#pl_login_form > div > div:nth-child(3) > div.info_list.verify.clearfix > div > input")
        for ch in code:
            code_input.send_keys(ch)
        print("提交请求中....")

        # click login button
        button = self.browser.find_elements_by_css_selector("#pl_login_form > div > div:nth-child(3) > div.info_list.login_btn > a")
        button[0].click()
        time.sleep(8)

        if self.browser.find_element_by_css_selector("#v6_pl_rightmod_myinfo > div > div > div.WB_innerwrap > div > a.name.S_txt1"):
            print("登录成功！")

        # get cookies
        cookies = self.browser.get_cookies()
        with open('./cookies_log/log.txt', "w+") as f:
            f.write(str(cookies))
        self.login_flag = True


    def search(self, keywords):
        self.keywords = keywords
        # search information according to keywords
        assert self.login_flag, "登录未成功，请先执行登录操作"
        wait = WebDriverWait(self.browser, 15)
        input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#plc_top > div > div > div.gn_search_v2 > input")))

        for word in self.keywords:
            input.send_keys(word)
            input.send_keys(' ')
        input.send_keys(Keys.ENTER)
        wait = WebDriverWait(self.browser, 15)
        _ = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pl_feedlist_index > div:nth-child(1) > div:nth-child(2) > div.card-top > h4")))


    def grabSinglePage(self):
        contentElements = []

        baseFileName = 'sina_data'
        for i in self.keywords:
            baseFileName += '_'
            baseFileName += i
        baseFileName = baseFileName + '.csv'
        with open(baseFileName, 'a+', encoding='utf-8-sig', newline='') as csvfile:
            fieldsname = ['account', 'content', 'time']
            writer = csv.DictWriter(csvfile, fieldnames=fieldsname)
            # weibo contents id are in range(2,24)
            for i in range(1, 25):
                try:
                    contentElements.append('#pl_feedlist_index > div:nth-child(1) > div:nth-child('+ str(i) +') > div.card > div.card-feed > div.content')
                except:
                    pass
            for element in contentElements:
                try:
                    account = self.browser.find_element_by_css_selector(element + ' > div.info > div:nth-child(2) > a.name').text
                    content = self.browser.find_element_by_css_selector(element + ' > p.txt').text
                    timeStr = self.browser.find_element_by_css_selector(element + ' > p.from > a:nth-child(1)').text
                    if re.match("^\S*分钟前$", timeStr):
                        time = int(timeStr.replace("分钟前", ""))
                        delta = datetime.timedelta(minutes=time)
                        time = (datetime.datetime.now() - delta).strftime('%m{m}%d{d} %H:%M').format(m='月', d='日')
                    elif re.match("^\S*秒前$", timeStr):
                        time = int(timeStr.replace("秒前", ""))
                        delta = datetime.timedelta(minutes=time)
                        time = (datetime.datetime.now() - delta).strftime('%m{m}%d{d} %H:%M').format(m='月', d='日')
                    elif re.match("^今天.+$", timeStr):
                        time = datetime.datetime.now().strftime('%m{m}%d{d} ').format(m='月', d='日') + timeStr.replace("今天", "")
                    elif re.match("^\S*分钟前 转赞人数超过\d*$", timeStr):
                        time = int(timeStr.split(" ")[0].replace("分钟前", ""))
                        delta = datetime.timedelta(minutes=time)
                        time = (datetime.datetime.now() - delta).strftime('%m{m}%d{d} %H:%M').format(m='月', d='日')
                    else:
                        time = timeStr

                    # write to csv file
                    data = {"account": account, "content": content, "time": time}
                    writer.writerow(data)
                except:
                    pass

    def grabPages(self, pages):
        '''
        :param pages: the number of pages you want to grab(limit 0, 50)
        '''
        for page in range(0, int(pages)):
            # first, grad single page's information
            print("正在爬取第" + str(page+1) + "页......")
            self.grabSinglePage()
            # second, click the button to next page
            if page == 0:
                nextButton = self.browser.find_element_by_css_selector("#pl_feedlist_index > div.m-page > div > a")
            else:
                nextButton = self.browser.find_element_by_css_selector("#pl_feedlist_index > div.m-page > div > a.next")
            try:
                nextButton.click()
            except:
                print("爬取结束！")
            # sleep few seconds
            duration = (random.random() + 3) * 1.5
            time.sleep(duration)


class Frame:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("新浪微博——爬虫")
        self.window.geometry('500x300')

        self.worker = Spider()
        self.keywordsList = None

        self.e1 = None
        self.e2 = None
        self.input = None
        self.com = None
        self.codeInput = None
        self.img = None

        self.login()

    def login(self):

        label1 = tk.Label(self.window, text="请输入爬取关键字，中间用空格分开")
        label1.pack()
        self.input = tk.Entry(self.window, show=None, font=('Arial', 9))
        self.input.pack()

        label2 = tk.Label(self.window, text="请选择要爬取的微博页数")
        label2.pack()
        self.com = tk.ttk.Combobox(self.window)
        self.com.pack()
        self.com["value"] = (5, 10, 15, 20, 25, 30, 35, 40, 45, 50)
        self.com.current(0)

        label3 = tk.Label(self.window, text="请输入微博账号和密码：")
        label3.pack()

        self.e1 = tk.Entry(self.window, show=None, font=('Arial', 14))
        self.e1.pack()
        self.e2 = tk.Entry(self.window, show='*', font=('Arial', 14))
        self.e2.pack()

        loginButton = tk.Button(self.window, text="开始爬取", command=self.post)
        loginButton.pack()

        self.window.mainloop()

    def newTop(self):
        top = Toplevel()
        top.title("输入验证码")
        top.geometry('280x150')
        codeArea = tk.Label(top, image=self.img)
        codeArea.pack(pady=3)

        self.codeInput = tk.Entry(top, show=None, font=('Arial', 9))
        self.codeInput.pack(pady=10)

        button = tk.Button(top, text="提交", command=self.postCode)
        button.pack(side = "bottom", pady = 13)


    def post(self):

        username = self.e1.get()
        password = self.e2.get()
        self.keywordsList = self.input.get().split(' ')

        self.worker.login(username=username, password=password)

        if self.worker.login_flag:
            tk.messagebox.showinfo("登录提示", self.worker.username + ":登录成功")
        else:
            codeImage = PIL.Image.open(self.worker.codePath)
            self.img = ImageTk.PhotoImage(codeImage)
            self.newTop()
            return
        self.grab()


    def postCode(self):
        code = self.codeInput.get()
        self.worker.loginWithCode(code)
        self.grab()


    def grab(self):
        if self.worker.login_flag:
            pages = self.com.get()
            self.worker.search(self.keywordsList)
            self.worker.grabPages(pages)


if __name__ == "__main__":
    ui = Frame()