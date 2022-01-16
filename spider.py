# coding ='utf-8 '
# @Time : 2022/1/10 21:15
# @File : spider.py
# @Software: PyCharm

import pandas as pd
import requests
import json
import os
import csv


class One_log:
    def __init__(self, id, item_name, example_GUID, form_GUID, example_GUID_filename, form_GUID_filename):
        self.event_id = id
        self.Item_name = item_name
        self.example_GUID = example_GUID
        self.form_GUID = form_GUID
        self.example_GUID_filename = example_GUID_filename
        self.form_GUID_filename = form_GUID_filename



def make_message(dict):
    list = dict["AUDIT_MATERIAL"]
    log = []
    for item in list:

        if len(item["EXAMPLE_GUID"]) > 0 or len(item["FORM_GUID"]) > 0:
            material_nm = item["MATERIAL_NAME"].replace("\t", "")
            download_path = id + "\\" + material_nm + "\\"

            path_build(download_path)

            try:
                sample = item["EXAMPLE_GUID"][0]
                sample_url = "https:" + sample["FILEPATH"]
                sample_path = download_path + "示例文件_" + sample["ATTACHNAME"]
                example_GUID_filename = sample["ATTACHNAME"]
            except:
                sample = {}
                sample_url = ""
                sample_path = ""
                example_GUID_filename = ""

            try:
                blank_form = item["FORM_GUID"][0]
                blank_form_url = "https:" + blank_form["FILEPATH"]
                blank_form_path = download_path + "空白表格_" + blank_form["ATTACHNAME"]
                form_GUID_filename = blank_form["ATTACHNAME"]

            except:
                blank_form = {}
                blank_form_url = ""
                blank_form_path = ""
                form_GUID_filename = ""

            print(sample_url)
            print(blank_form_url)
            One_log = {
                "event_id": id,
                "Item_name": material_nm,
                "example_GUID": sample_url,
                "form_GUID": blank_form_url,
                "example_GUID_filename": example_GUID_filename,
                "form_GUID_filename": form_GUID_filename,
            }

            # 下载文件
            if sample_url:
                download(sample_url, sample_path)
            if blank_form_url:
                download(blank_form_url, blank_form_path)

            Save_logtxt(One_log)

            log.append(One_log)
            print("-----------")
        else:
            print("样本和表单为空,跳过!")

    return log

def respose_by_id(id):
    url = "https://www.gdzwfw.gov.cn/portal/api/v2/item-event/getAuditItemDetailCur?TASK_CODE=" + str(id)
    headers = {  # 模拟浏览器头部信息
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82 Safari/537.36"
    }
    response = requests.get(url, headers=headers).text
    json1 = json.loads(response)
    return json1

#写入csv
def Save_logcsv(Log):
    with open('Log.csv', 'a', encoding="utf-8-sig", newline='')as f:
        writer = csv.DictWriter(f, fieldnames=['event_id', 'Item_name', 'example_GUID', 'form_GUID','example_GUID_filename','form_GUID_filename'])
        writer.writeheader()
        for one_log in Log:
            writer.writerow(one_log)
    # keys = ['event_id', 'Item_name', 'example_GUID', 'form_GUID','example_GUID_filename','form_GUID_filename']
    # dataframe = pd.DataFrame(columns=keys,data=log)
    # dataframe.to_csv('D:\spider\Log.csv', index=False,sep=',')
    # for one_log in log:
    #     dataframe = pd.DataFrame(data=one_log,index=[0])
    #     dataframe.to_csv('D:\spider\Log.csv', index=True,sep=',')

#写入Log.txt
def Save_logtxt(Log):
    with open("D:\spider\Log.txt", 'a', encoding='utf-8')as f:
        r = json.dumps(Log, ensure_ascii=False)
        f.write(r)
        f.write("\n")
        f.close()

def download(url, path):
    headers = {  # 模拟浏览器头部信息
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82 Safari/537.36"
    }
    req = requests.get(url, headers=headers)
    with open(path, "wb") as f:
        f.write(req.content)
        f.close()
    print("下载完成")
#创建目录
def path_build(path):
    isExists = os.path.exists(path)
    if not isExists:
        os.makedirs(path)
        print(f"创建目录{path}")
if __name__ == '__main__':
    df = pd.read_excel("D:\spider\佛山市实施清单明细表（8类）-0109.xlsx",nrows=11157, usecols=[13], names=None,dtype=str)
    #以str形式读取excel第14列，范围11157行
    df_li=df.values.tolist()#转为列表
    del (df_li[0])#去掉“实施编码”
    for id in df_li:
        id=''.join(id)
        try:
            # check_path(id)
            dict = respose_by_id(id)
            log = make_message(dict)
        except:
            with open("D:\spider\error.txt", 'a')as f:
                f.write(id)
                f.write("\n")
                f.close()
            continue
        Save_logcsv(log)