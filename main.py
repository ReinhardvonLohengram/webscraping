import requests, re, time
import xlwt, xlrd
from number import department
from xlutils.copy import copy
import pandas as pd
import random
def first(value, headers):
    response = requests.get(f"https://zyk.bjhd.gov.cn/jbdt/{value}/index_bm.shtml", headers=headers,timeout=(3,7))
    response.encoding='UTF-8' #编码转换
    html = response.text
    urls = re.findall(r"document.write\('<a href\=\"\.\/(\S+)\"", html)
    for url in urls:
        file = []
        date = []
        type = []
        time.sleep(0.2)
        response_detail = requests.get(f"https://zyk.bjhd.gov.cn/jbdt/{value}/{url}", headers=headers,timeout=(3,7))
        response_detail.encoding = 'UTF-8'
        if response_detail.status_code == 404:
            continue
        else:
            detail = response_detail.text
            san = re.findall(r"三山五园", detail)
            if san == []:
                continue
            else:
                date.append(re.findall(r"content=\"(\d{4}-\d{2}-\d{2})", detail))
                file.append(re.findall(r"<meta name\=\"ArticleTitle\" content\=\"(\S+)\">", detail))
                type.append(re.findall(r"<meta name\=\"ColumnName\" content\=\"(\S+)\">", detail))
                workbook = xlrd.open_workbook(f"./{key}.xls", formatting_info=True)
                newbook = copy(workbook)
                worksheet = newbook.get_sheet(0)
                data = pd.read_excel(f"./{key}.xls")
                i = len(data)
                worksheet.write(i+1, 0, type[0])
                worksheet.write(i+1, 1, date[0])
                worksheet.write(i+1, 2, file[0])
                newbook.save(f"./{key}.xls")

def after(value, headers):
    for start_num in range(1, 150):# index_bm    index_bm_1   index_bm_2...
        print(start_num)
        response = requests.get(f"https://zyk.bjhd.gov.cn/jbdt/{value}/index_bm_{start_num}.shtml", headers=headers,timeout=(3,7))
        if response.status_code == 404:
            break
        response.encoding = 'UTF-8'  # 编码转换
        html = response.text
        urls = re.findall(r"document.write\('<a href\=\"\.\/(\S+)\"", html)
        for url in urls:
            file = []
            date = []
            type = []
            time.sleep(random.uniform(0.2, 0.4))
            response_detail = requests.get(f"https://zyk.bjhd.gov.cn/jbdt/{value}/{url}", headers=headers,
                                           timeout=(3, 7))
            response_detail.encoding = 'UTF-8'
            if response_detail.status_code == 404:
                continue
            else:
                detail = response_detail.text
                san = re.findall(r"三山五园", detail)
                if san == []:
                    continue
                else:
                    date.append(re.findall(r"content=\"(\d{4}-\d{2}-\d{2})", detail))
                    file.append(re.findall(r"<meta name\=\"ArticleTitle\" content\=\"(.+)\">", detail))
                    type.append(re.findall(r"<meta name\=\"ColumnName\" content\=\"(\S+)\">", detail))
                    workbook = xlrd.open_workbook(f"./{key}.xls", formatting_info=True)
                    newbook = copy(workbook)
                    worksheet = newbook.get_sheet(0)
                    data = pd.read_excel(f"./{key}.xls")
                    i = len(data)
                    worksheet.write(i+1, 0, type[0])
                    worksheet.write(i+1, 1, date[0])
                    worksheet.write(i+1, 2, file[0])
                    newbook.save(f"./{key}.xls")

def excel(path):
    workbook = xlwt.Workbook()  # 新建一个工作簿
    workbook.add_sheet('sheet1')  # 在工作簿中新建一个表格
    worksheet = workbook.get_sheet(0)
    for j in range(0, len(titles)):
        worksheet.write(0, j, str(titles[j]))  # 表格中写入数据（对应的行）
    workbook.save(path)  # 保存工作簿

if __name__=='__main__':
    headers = {"User-Agent": "",
               'Connection':'close'}
    titles = ['类别', '日期', '名称']
    for key in department.keys():
        excel(f"./{key}.xls")
        first(department.get(key), headers)
        after(department.get(key), headers)


