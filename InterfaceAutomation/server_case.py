#!/usr/bin/python  
#encoding:utf-8 

#****************************************************************    
# Description: case处理区 
#****************************************************************  
  
from TestFrame import *  
import json
import requests
# import sys
# reload(sys)
# sys.setdefaultencoding( "utf-8" )


def run(suiteid):

    print("开始清除实际结果，请稍等。。。")

    # 清除实际结果数据
    excelobj.del_testrecord(suiteid)

    # print("清除数据完成。。。")

    # print ('【'+suiteid+'】' + ' Test Begin,please waiting...').decode('utf-8')
    print('【' + suiteid + '】' + ' Test Begin,please waiting...')

    global checkitem, requesturi, mailbody

    mailbody=[]

    head = {"Content-Type":"application/Json","X-AUTH-TOKEN":excelobj.getToken("15810420717","111111")} # 输入手机号和密码获取token

    # 获取接口地址（表中第一行接口地址）
    requesturi = excelobj.read_data(suiteid, 0, 1)
    # print("API url: ",requesturi)

    #获取case个数
    casecount = excelobj.get_ncase(suiteid)
    print('【' + suiteid + '】' +'测试用例总数： ',casecount)

    #遍历执行case
    for caseid in range(0, casecount):
        #获取预期结果的key字段
        checkitem = excelobj.read_data(suiteid, caseid+4, 13)
        # print("预期结果： ",checkitem)

        # 拼接接口url
        # sArge得到的是除了接口以为的参数拼接串，最多十个参数
        sArge = excelobj.GetPra(suiteid,caseid)
        # print("参数： ",sArge)

        # 拼接为完整的url，并对"&="的字符做处理
        fullURL=(requesturi+sArge).replace("&=","")
        # print("API地址： ",fullURL)

        # excelobj.read_data(suiteid, 2, 1) 用来获取请求方式（post、get）
        if excelobj.read_data(suiteid, 2, 1) == "post":
            # print("开始进入post请求")
            sArge = excelobj.PostPra(suiteid, caseid)
            # print("请求的参数： ",sArge)
            # print(type(sArge))
            # sArge = json.dumps(sArge)    #讲python对象解码为json数据
            ## r = requests.post(requesturi,sArge,headers=head)
            data = excelobj.read_data(suiteid, 4, 1)
            if excelobj.read_data(suiteid, 3, 1) == "json":
                r = requests.post(requesturi, data=data, headers=head)
            else:
                r = requests.post(requesturi, sArge, headers=head)
        else:
            r = requests.get(fullURL, headers=head)

        print("请求地址: ", r.url)
        # print("r.content: ", r.content)
        # print(type(r.content))
        print("响应数据: ", r.text)
        # print(type(r.text))
        try:
            excelobj.write_data(suiteid, caseid+4, 11, r.url)
            excelobj.write_data(suiteid, caseid+4, 12, r.text)
        except Exception as e:
            print(Exception,":",e)


        dict = json.loads(r.content)

        try:
            if checkitem == "0":
                excelobj.write_data(suiteid,caseid+4, 14, dict["code"])
                excelobj.assert_result(dict["code"], 0)

            else:
                excelobj.write_data(suiteid,caseid+4, 14, dict["message"])
                excelobj.assert_result(dict["message"], checkitem)

        except Exception as e:
            excelobj.compare_result(suiteid,caseid+4, 14)
            print(Exception,":",e)
            continue

        expect_result, actual_result, result = excelobj.compare_result(suiteid,caseid+4, 14)

        mailbody.append(r.url)
        mailbody.append(expect_result)
        mailbody.append(actual_result)
        mailbody.append(result)

        # print ("mailbody: ", mailbody)

        print("第" + str(caseid+1) + "条测试用例执行完成")
        print()
        # print ResString.decode('utf-8')
    # print ('【'+suiteid+'】' + ' Test End!' + '\n' +'**********************************************************************************************************').decode('utf-8')
    print('【' + suiteid + '】' + ' Test End!' + '\n' + '**********************************************************************************************************')

    count = excelobj.statisticresult(excelobj)
    print()
    return mailbody,count