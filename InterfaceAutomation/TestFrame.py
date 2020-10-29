#!/usr/bin/python  
#encoding:utf-8 

#****************************************************************    
# Description: 主要函数文件  
#**************************************************************** 
import datetime,time,requests
import xlrd,xlwt
from xlutils.copy import copy
# from xml2dict import XML2Dict
import win32com.client  
# import xml.etree.ElementTree as et
import sendEmail

class create_excel:

    global allresult
    allresult=[]

    def __init__(self, sFile,suiteid):
        #定义参数个数、请求发送方法、预期结果中检查的项
        global sFilepath
        sFilepath = sFile
        self.workbook = xlrd.open_workbook(sFile, formatting_info=True)
        # #global argsconut,reqmethod,reqHeaders
        # #self.xlApp = win32com.client.Dispatch('Excel.Application')   #MS:Excel  WPS:et
        # self.xlApp = win32com.client.Dispatch('et.Application')  # MS:Excel  WPS:et
        # self.workbook = xlrd.open_workbook(r'D:\work\InterfaceAutomation\TestCaseDir\Testcase.xls',formatting_info=True)
        ##try:
            ##self.book = self.xlApp.Workbooks.Open(sFile)
            # self.workbook = xlrd.open_workbook(sFile, formatting_info=True)
        ##except:
            # print ("打开文件失败").decode('utf-8')
            ##print("打开文件失败")
            ##exit()
        ##if suiteid == 'ALL':
            ##suiteid =  self.book.Worksheets(1).Name
        ##self.file=sFile
        # self.allresult=[]

    # def setUp(self):
        # self.workbook = xlrd.open_workbook(r'/Users/apple/Desktop/YXF/interface/TestCaseDir/Testcase.xls', formatting_info=True)
        # self.workbook = xlrd.open_workbook(r'D:\work\InterfaceAutomation\TestCaseDir\Testcase.xls',formatting_info=True)

    def GetUtf8Str(content):
        try:
            #如果是unicode字符，则进行utf-8编码
            value = content.encode("utf-8")
            return value
        except:
            #否则就是str类型
            #先进行gbk解码
            try:
                value = content.decode("gbk").encode("utf-8")
                return value
            except:
                #否则进行utf-8解码
                try:
                    value = content.decode("utf-8").encode("utf-8")
                    return value
                except:
                    #如果都不是，返回空，暂时写到这
                    return str(value)

    def GetGBKStr(content):
        try:
            #如果是unicode字符，则进行gbk编码
            value = content.encode("gbk")
            return value
        except:
            #否则就是str类型
            #先进行gbk解码
            try:
                value = content.decode("gbk").encode("gbk")
                return value
            except:
                #否则进行utf-8解码
                try:
                    value = content.decode("utf-8").encode("gbk")
                    return value
                except:
                    #如果都不是，返回空，暂时写到这
                    return str(value)

    def GetStr(text):
        if type(text) == str:
            return text
        else:
            try:
                return GetUtf8Str(text)

            except:
                return GetGBKStr(text)

#统一变为为unicode

    def setColor(self, color=1):
        # 更改背景色
        # PASS 通过为绿色, FAIL 不通过为红色
        # Create the Pattern
        pattern = xlwt.Pattern()
        # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
        # 设置底纹的前景色，对应为excel文件单元格格式中填充中的背景色
        pattern.pattern_fore_colour = color
        # Create the Pattern
        style = xlwt.XFStyle()
        # Add Pattern to Style
        style.pattern = pattern
        return style

    def get_all_sheetname(self):
        #获取所有sheet页名称
        # self.workbook = xlrd.open_workbook(r'D:\study\newInterface\TestCaseDir\dyhoa_Testcase.xls', formatting_info=True)
        # self.workbook = xlrd.open_workbook(r'D:\work\InterfaceAutomation\TestCaseDir\Testcase.xls', formatting_info=True)
        return self.workbook.sheet_names()

    def sheetnameToIndex(self, iSheet):
        # 获取所有sheet表名
        l = self.workbook.sheet_names()
        # 获取列表中某个值的索引
        return l.index(iSheet)

    def del_testrecord(self,iSheet):

        # 从第5行（ 第15和16列）开始是实际结果和测试结果（PASS/FAIL）
        iRowbegin = 4

        try:
            self.workbook = xlrd.open_workbook(r'D:\work\InterfaceAutomation\TestCaseDir\Testcase.xls', formatting_info=True)
            sht = self.workbook.sheet_by_name(iSheet)
            iRowend = sht.nrows
            for i in range(iRowbegin,iRowend):
                for j in range(14,16):
                    self.write_data(iSheet, i, j, 'NT')
            print("清除数据完成。。。")
        except Exception as e:
            print(Exception,":",e)
            # print('清除数据失败').decode('utf-8')
            print('清除数据失败')

    def getToken(self, username, pwd):  # 根据账号和密码登录获得token

        url = "http://projtest.jingcaiwang.cn:18080/rbac/sys/login"
        try:
            headers = {
                "Content-Type": "application/json"  # 设置登录接口的请求头
            }
            data = {  # 登录接口传参
                "loginName": username,  # 账号
                "pwd": pwd,  # 密码
                "os": 1,
                "idEntity": 1
            }
            res = requests.post(url, headers=headers, json=data)  # 发起请求
            print("token: ",res.json()['body']['token'])
            return res.json()['body']['token']  # 获取token
        except Exception as err:
            print(err)

    def ToUnicode(self,text):
            """
            | ##@函数目的: 将字符串转化成unicode字符串
            | ##@参数说明：
            | ##@返回值：  text的字符串形式
            | ##@函数逻辑：先后以utf8、gbk、utf16的形式转化text。通常不考虑utf16be的情形
            """
            result = text
            if type(text) == str:
                try:
                    result = text.decode("utf8")
                    if result.encode("utf8") == text:
                        pass
                    else:
                        raise Exception("not right conversion")
                except:
                    try:
                        result = text.decode("gbk")
                        if result.encode("gbk") == text:
                            pass
                        else:
                            raise Exception("not right conversion")
                    except:
                        try:
                            result = text.decode("utf16")
                            if result.encode("utf16") == text:
                                pass
                            else:
                                raise Exception("not right conversion")
                        except:
                            pass

            return result

    # （get方式）sArge得到的是除了接口以为的参数拼接串，最多十个参数
    def GetPra(self, iSheet, iRow):
        sArge=""
        sht = self.workbook.sheet_by_name(iSheet)
        for i in range(10):
            # if sht.cell(iRow+2,i+1).value:
            #     sArge = sArge + self.read_data(iSheet, 2, i+1) + "=" + self.read_data(iSheet, iRow+3, i+1) + "&"
            sArge = sArge + self.read_data(iSheet, 3, i+1) + "=" + self.read_data(iSheet, iRow+4, i+1) + "&"
        return sArge

    # （post方式）sArge得到的是除了接口以为的参数拼接串，最多十个参数
    def PostPra(self, iSheet, iRow):
        sArge = ""
        sht = self.workbook.sheet_by_name(iSheet)
        for i in range(10):
            sArge = sArge + '"'+ self.read_data(iSheet, 3, i + 1) + '":'+ '"'+self.read_data(iSheet, iRow + 4, i + 1) + '"'+ ","
        sArge = "{" + sArge + "}"
        sArge = sArge.replace('"":"",','')
        sArge = sArge.replace(',}', '}')
        return sArge

    # 获取指定行和列的值，字符为str类型，汉字为unicode
    def read_data(self, iSheet, iRow, iCol):
        global sValue
        try:
            sht = self.workbook.sheet_by_name(iSheet)
            sValue = sht.cell(iRow, iCol).value
            # 为了兼容中文，这里做一下try处理。
            try:
                sValue=self.GetStr(sValue)
            except:
                sValue=str(sValue)
        except Exception as e:
            # print(str(iRow)+'行'+str(iCol)+'列读取数据失败').decode('utf-8')
            print(Exception,":",e)
        #去除'.0'
        if sValue[-2:]=='.0':
            sValue = sValue[0:-2]

        return sValue

    def write_data(self, iSheet, iRow, iCol, sData, color=setColor(1)):
        try:
            # self.workbook = xlrd.open_workbook(r'D:\yxf\interface\TestCaseDir\dyhoa_Testcase.xls', formatting_info=True)
            self.workbook = xlrd.open_workbook(r'D:\work\InterfaceAutomation\TestCaseDir\Testcase.xls', formatting_info=True)
            # sht = self.workbook.sheet_by_name(iSheet)
            sData=self.ToUnicode(sData)
            wb = copy(self.workbook)
            #通过get_sheet()获取的sheet有write()方法
            ws = wb.get_sheet(self.sheetnameToIndex(iSheet))
            ws.write(iRow, iCol, sData)
            wb.save(r'D:\work\InterfaceAutomation\TestCaseDir\Testcase.xls')


        except Exception as e:
            # print(str(iRow)+'行'+str(iCol)+'列写入数据失败').decode('utf-8')
            print(str(iRow) + '行' + str(iCol) + '列写入数据失败')
            print(Exception,":",e)

    def compare_result(self, iSheet, iRow, iCol):

        try:
            # self.workbook = xlrd.open_workbook(r'D:\yxf\interface\TestCaseDir\dyhoa_Testcase.xls', formatting_info=True)
            self.workbook = xlrd.open_workbook(r'D:\work\InterfaceAutomation\TestCaseDir\Testcase.xls', formatting_info=True)
            wb = copy(self.workbook)
            #通过get_sheet()获取的sheet有write()方法
            ws = wb.get_sheet(self.sheetnameToIndex(iSheet))
            # 获取期望结果和实际结果
            # 第14列是预期结果，第15列是实际结果，第16列是测试结果（PASS/FAIL）
            expect_result = self.read_data(iSheet, iRow,iCol-1)
            actual_result = self.read_data(iSheet, iRow,iCol)
            result = "PASS"
            print("预期结果: ", expect_result)
            print("实际结果: ", actual_result)
            if expect_result == actual_result:
                ws.write(iRow,iCol+1, "PASS", self.setColor(3))
                allresult.append(result)
                print("结果: " + result)
            else:
                result = "FAIL"
                ws.write(iRow,iCol+1,"FAIL", self.setColor(2))
                allresult.append(result)
                print("Result: " + result)

            time.sleep(3)
            # wb.save(r'D:\yxf\interface\TestCaseDir\dyhoa_Testcase.xls')
            wb.save(r'D:\work\InterfaceAutomation\TestCaseDir\Testcase.xls')
        except Exception as e:
            print(Exception,"excelobj:",e)

        return expect_result, actual_result, result

    #测试结果计数器，类似Switch语句实现
    def countflag(self, flag, ok, fail):
        calculation  = {'PASS':lambda:[ok+1, fail],
                             'FAIL':lambda:[ok, fail+1]}
        # print "calculation: ", calculation
        # print "calculation[flag](): ", calculation[flag]()
        return calculation[flag]()

    #打印测试结果
    def statisticresult(self, excelobj):
        # allresultlist=excelobj.allresult
        allresultlist = allresult
        print("allResultList: ", allresultlist)
        count=[0, 0]
        for i in range(0, len(allresultlist)):
            #print 'case'+str(i+1)+':', allresultlist[i]
            # count=self.countflag(allresultlist[i],count[0], count[1])
            count=self.countflag(allresultlist[i],count[0], count[1])
        print('Statistic result as follow:')
        print('OK:', count[0])
        print('FAIL:', count[1])
        return count

    def get_ncase(self, iSheet):
        try:
            sht = self.workbook.sheet_by_name(iSheet)
            cases = sht.nrows-4
            return cases
        except:
            # print('获取Case个数失败').decode('utf-8')
            print('获取Case个数失败')

    #结果判断
    def assert_result(self, sReal, sExpect):
        #预期结果和实际结果都是空的情况需要特别处理下
        if  sReal is None and sExpect == 'None':
            return 'OK'
        sReal=self.ToUnicode(sReal)
        sExpect=self.ToUnicode(sExpect)
        if sReal==sExpect:
            return 'OK'
        else:
            return 'FAIL'

    def close(self):
        self.book.Save()
        self.book.Close()

def Sendmail(maillist,bodycontent,count,interfacename ):
    strHtml = ''
    interfacename = str(interfacename)
    strHtml += '<B><p style="font-size:16px">' + interfacename + '测试结果:</p></B>'
    strHtml += '<B><p style="font-size:14px">' + 'Statistic result as follow:</p></B>'
    strHtml += '<B><p style="font-size:14px">' + 'PASS:\t' + str(count[0]) + '</p></B>'
    strHtml += '<B><p style="font-size:14px">' + 'FAIL:\t' + str(count[1]) + '</p></B>'
    strHtml += '<table width="1000" border="1">'
    strHtml += '<tr>'
    strHtml += '<td>'+'接口url'+'</td>'
    strHtml += '<td>'+'预期结果'+'</td>'
    strHtml += '<td>'+'实际结果'+'</td>'
    strHtml += '<td>'+'结果'+'</td>'
    strHtml += '</tr>'

    for i in range(len(bodycontent)):
        content=''.join(bodycontent[i])
        # print("content: ", content)
        # content = GetStr(content)
        # content = str(content)
        #print type(content)
        if i==0 :
            strHtml += '<tr>'
        strHtml += '<td>' + content + '</td>'
        if (i+1) % 4 == 0:
            strHtml += '</tr><tr>'
    strHtml = strHtml[:-4]
    strHtml += '</table>'
    mail_from = 'interfacetest11@163.com'
    mail_to = maillist
    timenow = datetime.datetime.utcnow() + datetime.timedelta(hours=8)#东8区增加8小时
    title = '【'+interfacename + '_接口测试结果】    '+ timenow.strftime( '%Y-%m-%d %H:%M:%S' )
    body = strHtml
    if count[1] == 0:
        print('All Case is OK!')
    else:
        sendEmail.SendMail(mail_from, mail_to, title, body)
    # sendEmail.SendMail(mail_from, mail_to, title, body)
