import pandas as pd
import requests
import xml.etree.ElementTree as ET
import time
import re

#################### Config ################################

## 檔案位置
file_path = '/test.csv'

## 公司登記
company_register = 'https://data.gcis.nat.gov.tw/od/data/api/F05D1060-7D57-4763-BDCE-0DAF5975AFE0?$format=xml&$filter=Business_Accounting_NO%20eq '

## 商業登記
commercial_register = 'https://data.gcis.nat.gov.tw/od/data/api/426D5542-5F05-43EB-83F9-F1300F14E1F1?$format=xml&$filter=President_No eq '

def data_conf():
    # limit_data_flag = False
    # limit = [2,101]
    
    limit_req = input('是否要限制查詢的數據範圍:\n請填入: y(yes) or (no) \n(提醒: 不填入、亂填或填no，就會執行全部數據)\n')
    
    req = limit_req.lower()
    
    if req == 'yes' or req == 'y':
        print('限制數據範圍')
        limit_data_flag = True
        dataset_length = len(pd.read_csv(file_path))
        print('資料數量: ', dataset_length)
        a = input('輸入起始位置(提醒:可以查看excel上的編號): ')
        b = input('輸入結束位置: ')
        if a.isdigit() == False or b.isdigit() == False:
            assert False, '位置請輸入數字'
        limit_start = int(a) - 2
        limit_end = int(b) - 1

        if limit_start >= limit_end:
            assert False, '起始位置需要比結束位置小，感謝'
            
        if limit_start > dataset_length or limit_end > dataset_length:
            assert False, '起始與結束位置不可以大於資料大小'
                
        limit = [limit_start, limit_end]
        
    else:
        print('執行預設狀況(全部資料)')
    
        limit_data_flag = False
        limit = []
    
    return limit_data_flag, limit
def load_data(path, flag, limit):
    dataset = pd.read_csv(path)
    ## test 1000 data
    ## dataset = dataset[:10]pip
    if flag == True:
        dataset = dataset[limit[0]:limit[1]]
    
    dataset['CUST_ID_NO'] = dataset['CUST_ID_NO'].astype('str')
    data = list(dataset['CUST_ID_NO'])
    
    ## 檢查是否有八位數，沒有就前面補0
    for i in range(len(data)):
        data[i] = data[i].zfill(8)
    
    return dataset, data
def exc(data, company_register):
    ## load
    status = []
    c_name = []
    c_type = []
## 檢查營業狀態: 公司登記
    progress = 1
    for i in data:
        time.sleep(3)
        print('------------------------------------------------')
        print('進度: ' + str(progress) + ' / '+ str(len(data)))
        print('查詢公司: ', i)
        progress += 1
        ## Filter out Person ID
        if re.match('[A-Z]', i):
            status.append('查無資料')
            c_name.append('查無資料')
            c_type.append('查無資料')
            print('為個人ID')
        else:
            html = requests.get(company_register+ i +'&$skip=0&$top=1')
            root = ET.fromstring(html.text)
#             print(html.text)
        #     print(root[0][3].text)
            if len(root) == 1:
                status.append(root[0][3].text)
                c_name.append(root[0][1].text)
                c_type.append('公司登記')
                print(root[0][3].text)
            elif len(root) == 0:
                print('公司登記查不到...')
                print('開始查詢商業登記狀態...')
                try:
                    ## 檢查營業狀態: 商業登記(經公司登記篩選為查無此資料的公司，再透過商業登記篩選)
                    html = requests.get(commercial_register+i+'&$skip=0&$top=1')
                    root = ET.fromstring(html.text)
                except:
                    print('系統發生異常')
#                     assert False, '商工系統發生異常: 建議手動查詢這間公司'
                    status.append('系統異常，請手動查詢')
                    c_name.append('系統異常，請手動查詢')
                    c_type.append('商業登記')
                else:
#                 ## 檢查營業狀態: 商業登記(經公司登記篩選為查無此資料的公司，再透過商業登記篩選)
#                 html = requests.get(commercial_register+i+'&$skip=0&$top=1')
#                 root = ET.fromstring(html.text)
                    if len(root) == 1:
                        status.append(root[0][3].text)
                        c_name.append(root[0][1].text)
                        c_type.append('商業登記')
                        print(root[0][3].text)
                    elif len(root) == 0:
                        status.append('查無資料')
                        c_name.append('查無資料')
                        c_type.append('查無資料')
                        print('公司登記與商業登記都查無資料')
    return status, c_name, c_type
def save(dataset, status, c_name, c_type):
    dataset['status'] = status
    dataset['company'] = c_name
    dataset['company_type'] = c_type
    ## 為了讓index列可以對應原始資料的位置
    origin_index = dataset.index + 2
    dataset.insert(0, '原始位置', origin_index)
    dataset.to_csv('C:/Users/michaely/Downloads/company_status.csv', encoding = 'utf-8-sig', index = False)
if __name__ == '__main__':
    limit_data_flag, limit = data_conf()
    dataset, data = load_data(file_path, limit_data_flag, limit)
    status, c_name, c_type = exc(data, company_register)
    save(dataset, status, c_name, c_type)
    print('執行完成')
