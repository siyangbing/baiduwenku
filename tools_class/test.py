import os
import time

from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from scrapy import Selector
import requests
from my_fake_useragent import UserAgent
import docx
from docx.shared import Inches
import cv2
from pptx import Presentation
from pptx.util import Inches

url = "https://wenku.baidu.com/view/ea815d9c1fd9ad51f01dc281e53a580217fc50e0.html"


# 浏览器操纵类
class ControlChrome():
    # 初始化并打开浏览器
    def __init__(self, chromedriver_path="./chromedriver"):
        capabilities = DesiredCapabilities.CHROME
        capabilities['loggingPrefs'] = {'browser': 'ALL'}
        options = webdriver.ChromeOptions()
        self.brower = webdriver.Chrome(executable_path=chromedriver_path, desired_capabilities=capabilities,
                                       options=options)

    def go_url(self, url):
        # 启动浏览器，打开需要下载的文档网页
        self.brower.get(url)

    def scrolled_to_end(self):
        self.brower.execute_script("window.scrollTo(0,document.body.scrollHeight)")

    def click_ele(self, click_xpath):
        # 单击指定控件
        click_ele = self.brower.find_elements_by_xpath(click_xpath)
        if click_ele:
            click_ele[0].location_once_scrolled_into_view  # 滚动到控件位置
            self.brower.execute_script('arguments[0].click()', click_ele[0])  # 单击控件，即使控件被遮挡，同样可以单击





if __name__ == "__main__":
    cc = ControlChrome()
    cc.go_url(url)
    cc.scrolled_to_end()
    # time.sleep(5)
    xpath_close_button = "//a[@id='TANGRAM__PSP_4__closeBtn']"
    cc.click_ele(xpath_close_button)
