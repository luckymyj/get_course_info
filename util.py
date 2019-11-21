from xlwt import Workbook
import os



def get_element_text(temp_driver, temp_elements_css):

    all_elements = temp_driver.find_elements_by_css_selector(temp_elements_css)
    rtn_text_list = []
    try:

        for each_element in all_elements:
            rtn_text_list.append(each_element.text)
    
    except NoSuchElementException:
        print('=====元素定位错误=======')

    return rtn_text_list

def get_dict_from_keyvalue(keys, *values):
    value_list = []
    for value in values:
          value_list.append(value)
    temp_dict = dict(zip(keys, value_list))
    return temp_dict


class ExcelWriter:
    def __init__(self, excel_name, sheet_name = "Sheet1"):
        self.sheet_name = sheet_name
        self.excel_name = excel_name
        self.workbook = Workbook(encoding = "utf-8")
        self.sheet = self.workbook.add_sheet(self.sheet_name)
    
    def write_data_by_dict(self, data_dict):
           if os.path.exists(self.excel_name):
                 os.remove(self.excel_name)
           if data_dict != None:
                title_list = list(data_dict.keys())
                for j in range(len(title_list)):
                    self.sheet.write(0, j, title_list[j])
                    for i in range(1, len(data_dict.get(title_list[j]))+1):
                        self.sheet.write (i,j,data_dict.get(title_list[j])[i-1])
                
                self.workbook.save (self.excel_name)
     
