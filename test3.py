# coding=utf-8
import requests, json, os, time
import xlwt

newlist = []
newdict = {}
newdict1 = {}
cookies = {'xb-gw-tag': 'v130', ' Hm_lvt_c0f6801399d3530bed7564e42dd32b4a': '1602323105',
           ' sensorsdata2015jssdkcross': '%7B%22distinct_id%22%3A%220%22%2C%22%24device_id%22%3A%2217511e644f6bb9-01b05013f89b27-193a6153-2073600-17511e644f7d45%22%2C%22props%22%3A%7B%22%24latest_traffic_source_type%22%3A%22%E7%9B%B4%E6%8E%A5%E6%B5%81%E9%87%8F%22%2C%22%24latest_referrer%22%3A%22%22%2C%22%24latest_referrer_host%22%3A%22%22%2C%22%24latest_search_keyword%22%3A%22%E6%9C%AA%E5%8F%96%E5%88%B0%E5%80%BC_%E7%9B%B4%E6%8E%A5%E6%89%93%E5%BC%80%22%7D%2C%22first_id%22%3A%2217511e644f6bb9-01b05013f89b27-193a6153-2073600-17511e644f7d45%22%7D',
           ' Hm_lvt_2a034bf7b5bd7f81e0722393ab96cce9': '1602323106', ' computerLoginfail': 'time%3D1%26maxtime%3D3',
           ' XSRF-TOKEN': '2ee10b4fecbe4634af31d830d2ee7e23', ' __root_domain_v': '.schoolpal.cn',
           ' _qddaz': 'QD.k8i02i.nyc5mq.kg3huckr', ' _qdda': '3-1.1', ' _qddab': '3-jeqw65.kg3wrg7j',
           ' _qddamta_2355128213': '3-0', ' SessionId': 'XIAOBAO-130-_f3d2c2c6-d062-4459-bf27-211e7be182a3',
           ' SessionId.c': 'XIAOBAO-130-_f3d2c2c6-d062-4459-bf27-211e7be182a3',
           ' Hm_lpvt_2a034bf7b5bd7f81e0722393ab96cce9': '1602351860',
           ' idtoken.c': 'eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1bmlxdWVfbmFtZSI6IjM2NjAwNjAiLCJ1c2VySWQiOjM2NjAwNjAsIm9yZ0lkIjo1NjAxMDcsImhyZG9jSWQiOjIzMDM3NjYsInZlciI6InYxMy4wIiwianRpIjoiOWRlOTgyZTQwOTZiMjE4MiIsInRoZW1lIjoiOCIsInVzZXIiOiJ3YW5neWFucWluZyIsIm9yZyI6ImJ3dCIsIm5iZiI6MTYwMjM1MTg3MSwiZXhwIjoxNjAyMzgwNjcxLCJpYXQiOjE2MDIzNTE4NzEsImlzcyI6Imlzc3Vlci5zY2hvb2xwYWwuY29tIiwiYXVkIjoic2Nob29scGFsLnBjIn0.GJeqlcoH1QA33t6lHOJ_afJKKcnxty3ogVbyecXgRBNEqPkOBQUoTjvcwKEUXplSJcRSZYTDVDAktUC39rsxKS-hAujhEcQzBiLmFABzrI_xXr_3FLrgn_-m9GZVsqlcjTev5wJx9a6raEuODT5McOshut_tCaaQ_WtTcX5OzZLuA3Ss6g6_RgKHZqi2urytCgSbIBMddxCmBqkI0w2Iw0EMGVkTcydofR5lZVAMT_5jG7SdHQkxe94y5d1JFn1H7JrG4ygK0-4ej311ueX54nAYCIKBvzUg7vt21zJLwDAKECzDVYRnN5UB54z64De2wQge5l148vheWiW1D3XLxA',
           ' acw_tc': '781bad2016023518849027872e25e496af97ebb57b8e954103dfbfaf5b93d7',
           ' Hm_lpvt_c0f6801399d3530bed7564e42dd32b4a': '1602352436'}
# cookies='xb-gw-tag=v130; Hm_lvt_c0f6801399d3530bed7564e42dd32b4a=1602323105; sensorsdata2015jssdkcross=%7B%22distinct_id%22%3A%220%22%2C%22%24device_id%22%3A%2217511e644f6bb9-01b05013f89b27-193a6153-2073600-17511e644f7d45%22%2C%22props%22%3A%7B%22%24latest_traffic_source_type%22%3A%22%E7%9B%B4%E6%8E%A5%E6%B5%81%E9%87%8F%22%2C%22%24latest_referrer%22%3A%22%22%2C%22%24latest_referrer_host%22%3A%22%22%2C%22%24latest_search_keyword%22%3A%22%E6%9C%AA%E5%8F%96%E5%88%B0%E5%80%BC_%E7%9B%B4%E6%8E%A5%E6%89%93%E5%BC%80%22%7D%2C%22first_id%22%3A%2217511e644f6bb9-01b05013f89b27-193a6153-2073600-17511e644f7d45%22%7D; Hm_lvt_2a034bf7b5bd7f81e0722393ab96cce9=1602323106; computerLoginfail=time%3D1%26maxtime%3D3; XSRF-TOKEN=2ee10b4fecbe4634af31d830d2ee7e23; __root_domain_v=.schoolpal.cn; _qddaz=QD.k8i02i.nyc5mq.kg3huckr; _qdda=3-1.1; _qddab=3-jeqw65.kg3wrg7j; _qddamta_2355128213=3-0; SessionId=XIAOBAO-130-_f3d2c2c6-d062-4459-bf27-211e7be182a3; SessionId.c=XIAOBAO-130-_f3d2c2c6-d062-4459-bf27-211e7be182a3; Hm_lpvt_2a034bf7b5bd7f81e0722393ab96cce9=1602351860; idtoken.c=eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1bmlxdWVfbmFtZSI6IjM2NjAwNjAiLCJ1c2VySWQiOjM2NjAwNjAsIm9yZ0lkIjo1NjAxMDcsImhyZG9jSWQiOjIzMDM3NjYsInZlciI6InYxMy4wIiwianRpIjoiOWRlOTgyZTQwOTZiMjE4MiIsInRoZW1lIjoiOCIsInVzZXIiOiJ3YW5neWFucWluZyIsIm9yZyI6ImJ3dCIsIm5iZiI6MTYwMjM1MTg3MSwiZXhwIjoxNjAyMzgwNjcxLCJpYXQiOjE2MDIzNTE4NzEsImlzcyI6Imlzc3Vlci5zY2hvb2xwYWwuY29tIiwiYXVkIjoic2Nob29scGFsLnBjIn0.GJeqlcoH1QA33t6lHOJ_afJKKcnxty3ogVbyecXgRBNEqPkOBQUoTjvcwKEUXplSJcRSZYTDVDAktUC39rsxKS-hAujhEcQzBiLmFABzrI_xXr_3FLrgn_-m9GZVsqlcjTev5wJx9a6raEuODT5McOshut_tCaaQ_WtTcX5OzZLuA3Ss6g6_RgKHZqi2urytCgSbIBMddxCmBqkI0w2Iw0EMGVkTcydofR5lZVAMT_5jG7SdHQkxe94y5d1JFn1H7JrG4ygK0-4ej311ueX54nAYCIKBvzUg7vt21zJLwDAKECzDVYRnN5UB54z64De2wQge5l148vheWiW1D3XLxA; acw_tc=781bad2016023518849027872e25e496af97ebb57b8e954103dfbfaf5b93d7; Hm_lpvt_c0f6801399d3530bed7564e42dd32b4a=1602352436'
url = 'https://pro.schoolpal.cn/api2/Stuinfo/GetStuInfoListData'
headers = {'content-type': 'application/json', 'X-XSRF-TOKEN': '2ee10b4fecbe4634af31d830d2ee7e23'}
data = {"query": "", "queryType": "", "pageIndex": 1, "pageSize": 500, "totalCount": 0, "asc": 'false',
        "accurateQuery": 'false', "orderKey": "", "extendSearchList": [], "enrollInfoStatus": ["1"], "schoolIds": [],
        "sex": [], "schoolPalHome": [], "faceSyncState": [], "isArrearage": [], "isNewStuInfo": [], "isCollection": [],
        "balance": [], "channelId": [], "channelCategoryId": []}

ret = requests.post(url=url, data=json.dumps(data), headers=headers, cookies=cookies).json()

print(ret)

# l=json.loads(studentlist)
data = ret['data']['list']
print(data)

for i in data:
    # newlist.append((i['data']['list']['id'],i['data']['list']['id']))
    id = i['id']
    telPhone = i['telPhone']
    stuName = i['stuName']

    newdict[stuName] = [id, telPhone]
print(newdict)
#
# url = 'https://pro.schoolpal.cn/api2/StuFeeDoc/GetStuInfoLessonList?stuInfoId=%s' % '140503470'
# headers = {'content-type': 'application/json'}
# ret = requests.get(url=url, headers=headers, cookies=cookies).json()
# print(ret)
# tt = ret['data']['stuEnrollList'][0]
# classTimes = tt['classTimes']
# print(classTimes)
#
# enrollat = tt['enrollat']
# print(enrollat)
#
# totalClassTimes = tt['totalClassTimes']
# print(totalClassTimes)
#
# usedClasstimes = tt['usedClasstimes']
# totalTuition = tt['totalTuition']
# shouru=int(totalTuition)/int(totalClassTimes) * int(usedClasstimes)
# print(shouru)
#
# newdict1['王彦青']=[totalTuition,totalClassTimes,usedClasstimes,classTimes,shouru,enrollat]
# print(newdict1)

ws = xlwt.Workbook(encoding='utf-8')
w = ws.add_sheet(u"数据报表第一页")
w.write(0, 0, "学员姓名")
w.write(0, 1, u"报名费")
w.write(0, 2, u"总课时数")
w.write(0, 3, u"使用课时数")
w.write(0, 4, u"剩余课时数")
w.write(0, 5, u"创收入")
w.write(0, 6, u"报名时间")
w.write(0, 7, u"报名预留电话")
excel_row = 1
for k, v in newdict.items():
    url = 'https://pro.schoolpal.cn/api2/StuFeeDoc/GetStuInfoLessonList?stuInfoId=%s' % v[0]
    headers = {'content-type': 'application/json'}
    ret = requests.get(url=url, headers=headers, cookies=cookies).json()
    for dv in ret['data']['stuEnrollList']:
        tt = ret['data']['stuEnrollList'][0]
        classTimes = tt['classTimes']
        enrollat = tt['enrollat']
        totalClassTimes = tt['totalClassTimes']
        usedClasstimes = tt['usedClasstimes']
        totalTuition = tt['totalTuition']
        shouru = int(totalTuition) / int(totalClassTimes) * int(usedClasstimes)
        newdict1[k] = [totalTuition, totalClassTimes, usedClasstimes, classTimes, shouru, enrollat, v[1], ]
        w.write(excel_row, 0, k)
        w.write(excel_row, 1, totalTuition)
        w.write(excel_row, 2, totalClassTimes)
        w.write(excel_row, 3, usedClasstimes)
        w.write(excel_row, 4, classTimes)
        w.write(excel_row, 5, shouru)
        w.write(excel_row, 6, enrollat)
        w.write(excel_row, 7, v[1])
        excel_row += 1
name = str(int(time.time()))

ws.save("test_%s.xls" % name)
