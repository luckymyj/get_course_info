from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from util import get_element_text
from util import get_dict_from_keyvalue
from util import ExcelWriter
import os

#保存腾讯课堂课程信息至文件
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(BASE_PATH, 'data')
excel_name = os.path.join(DATA_PATH, 'course_info.xls')
driver = webdriver.Chrome()
driver.get('https://ke.qq.com/course/list?mt=1001&st=2006&tt=3034&page=1&tuin=c9595489')
driver.maximize_window()

course_title_list = []
course_num_list = []
course_company_list = []
course_price_list = []

course_li_title_css = '[data-report-module="middle-course"] > ul.course-card-list > li > h4'
course_li_num_css = '[data-report-module="middle-course"] > ul.course-card-list > li > div[class="item-line item-line--bottom"] > span:nth-child(2)'
course_li_company_css = '[data-report-module="middle-course"] > ul.course-card-list > li > div[class="item-line item-line--middle"] > a.line-cell.item-source-link'
course_li_price_css = '[data-report-module="middle-course"] > ul.course-card-list > li > div[class="item-line item-line--bottom"] > span:nth-child(1)'
next_page_btn_css = 'a[class="page-next-btn icon-font i-v-right"]'

while True:
    course_title_list.extend(get_element_text(driver, course_li_title_css))
    course_num_list.extend(get_element_text(driver, course_li_num_css))
    course_company_list.extend(get_element_text(driver, course_li_company_css))
    course_price_list.extend(get_element_text(driver, course_li_price_css))
    try:
        next_page_btn = driver.find_element_by_css_selector(next_page_btn_css)
        next_page_btn.click()
    except NoSuchElementException:
        break
    
driver.close()
title_list = ["课程名称", "课程报名人数", "课程所属学院", "课程价格"]
course_dict = get_dict_from_keyvalue(title_list, course_title_list, course_num_list, course_company_list, course_price_list)
excelwrite= ExcelWriter(excel_name)
excelwrite.write_data_by_dict(course_dict)