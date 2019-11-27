from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from util import get_element_text
from util import get_dict_from_keyvalue
from util import ExcelWriter

from sqlserver import SQLSVR
import os

#保存腾讯课堂课程信息至文件
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(BASE_PATH, 'data')
excel_name = os.path.join(DATA_PATH, 'course_info.xls')

#如果保存类型为1，保存在excel，如果为2，保存到数据库表
save_type = 2

#链接数据库套接字
sql_conn_string = ['LUCKYMYJ-PC\SQLSERVER2012', 'sa', 'myj810714', 'courseinfo']
delete_course_string = "delete from courseinfos"
insert_course_string = "insert into courseinfos values(%d, \'%s\', \'%s\', \'%s\', \'%s\')"

driver = webdriver.Chrome()
driver.get('https://ke.qq.com/course/list?mt=1001&st=2006&tt=3034&page=1&tuin=c9595489')
driver.maximize_window()

course_title_list = []
course_num_list = []
course_company_list = []
course_price_list = []

#获取腾讯课堂中课堂链接等相关CSS_SELECTOR，
# 关于这步保存在脚本中有些累赘，可以将它抽离出来放在yaml文件或其他文件中读取出来
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
# for i in range(5):
#     course_title_list.extend(get_element_text(driver, course_li_title_css))
#     course_num_list.extend(get_element_text(driver, course_li_num_css))
#     course_company_list.extend(get_element_text(driver, course_li_company_css))
#     course_price_list.extend(get_element_text(driver, course_li_price_css))
#     try:
#         next_page_btn = driver.find_element_by_css_selector(next_page_btn_css)
#         next_page_btn.click()
#     except NoSuchElementException:
#         break   

driver.close()
if save_type == 1:
    title_list = ["课程名称", "课程报名人数", "课程所属学院", "课程价格"]
    course_dict = get_dict_from_keyvalue(title_list, course_title_list, course_num_list, course_company_list, course_price_list)
    excelwrite= ExcelWriter(excel_name)
    excelwrite.write_data_by_dict(course_dict)
elif save_type == 2:
    sqlserver_conn = SQLSVR(*sql_conn_string)   
    sqlserver_conn.execSql(delete_course_string)
    for j in range(len(course_title_list)):
        sqlserver_conn.execSql(insert_course_string%(j+1, course_title_list[j], course_num_list[j],  
                                    course_company_list[j], course_price_list[j]))
    # print(sqlserver_conn.querySql('select * from courseinfos'))