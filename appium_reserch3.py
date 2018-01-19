#coding=utf-8
from appium import webdriver
import xlrd
import time
from xlutils.copy import copy
import os

def Read_Col(Path,Index,Col):    
    Excel = xlrd.open_workbook(Path)
    Sheet=Excel.sheet_by_index(Index)
    Value= Sheet.col_values(Col)
    return  Value

def Exchange_Str(Value):
    Type1=type(Value)
    #float转str
    if Type1 is float:
        #先转换为长整型
        Long_Value=long(Value)
        #再转换为字符型
        Str_Value=str(Long_Value)
        return Str_Value        
    #str转unicode
    elif  Type1 is str:
        return Value        
    #unicode不转换
    elif Type1 is unicode:
        EValue=Value.encode('utf-8')
        return EValue    

desired_caps = {}
desired_caps['platformName'] = 'Android' #设置操作平台
desired_caps['platformVersion'] = '4.4.2' #操作系统版本
desired_caps['deviceName'] = 'emulator-5558' #设备名称
desired_caps['appPackage'] = 'cn.com.hkgt.gasapp' 
desired_caps['appActivity'] = '.AppStartActivity'
desired_caps['udid'] = 'emulator-5558'
desired_caps['connection'] = 'keep-alive'
desired_caps['agent'] = 'false'
driver = webdriver.Remote('http://127.0.0.1:4732/wd/hub', desired_caps)
time.sleep(20)

#try验证是否已经登录
driver.find_element_by_id("cn.com.hkgt.gasapp:id/login_button").click()
time.sleep(5)
driver.find_element_by_id("cn.com.hkgt.gasapp:id/login_account").clear()
time.sleep(5)
driver.find_element_by_id("cn.com.hkgt.gasapp:id/login_account").send_keys('liu199223')
time.sleep(5)
driver.find_element_by_id("cn.com.hkgt.gasapp:id/login_password").send_keys('liu19920723')
time.sleep(5)
driver.find_element_by_id("cn.com.hkgt.gasapp:id/btn_login_login").click()
time.sleep(5)

'''
except:
    print (u'用户未登录，进行登录操作')
    driver.find_element_by_id("id/login_account").send_keys('liuchinho') #登录操作
    driver.find_element_by_id("id/login_password").send_keys('liu19920723')
    driver.find_element_by_id("id/btn_login_login").click()
    time.sleep(10)
'''
driver.find_element_by_id('cn.com.hkgt.gasapp:id/traiff_item_chaxun_line_tv').click()
time.sleep(5)
driver.find_element_by_id('cn.com.hkgt.gasapp:id/select_card_state').click()
time.sleep(10)

def readcol(Path,Index,Col):    
    Excel = xlrd.open_workbook(Path)
    Sheet=Excel.sheet_by_index(Index)
    value= Sheet.col_values(Col)
    rows=Sheet.nrows
    return  value, rows
'''
def Exchange_Str(value):
    Type1=type(value)
    print Type1
    #float转str
    if Type1 is float:
        #先转换为长整型
        int_Value=int(value)
        #再转换为字符型
        Str_Value=str(int_Value)
        return Str_Value      
        print 1  
    #str转unicode
    elif  Type1 is str:
        return value
        print 2
    #unicode不转换
    elif Type1 is unicode:
        evalue=value.encode('utf-8')
        return evalue
        print 3    
'''
def Write(Path,Index,Row,Col,Value):
    Excel=xlrd.open_workbook(Path)
    Copy_Excel=copy(Excel)
    Sheet=Copy_Excel.get_sheet(Index)
    Type1=type(Value) 
    #将value转换为unicode(value为str格式) 
    if Type1 is str:
        Dvalue=Value.decode('utf-8')  
        Sheet.write(Row,Col,Dvalue) 
    elif Type1 is unicode:
        Sheet.write(Row,Col,Value) 
    os.remove(Path)   
    Copy_Excel.save(Path)
    
Path = 'C:\Users\Liuhaha\Desktop\GASserch3.xls'
Index = 0
Col = 0
v, r = readcol(Path, Index, Col)
j=3516
for i in v:
    type1=type(i)
    int_i=int(i)
    str_i=str(int_i)
    print j
    print str_i
#    driver.find_element_by_id('cn.com.hkgt.gasapp:id/input_card_number').clear()
    driver.find_element_by_id('cn.com.hkgt.gasapp:id/input_card_number').send_keys(str_i)
    driver.find_element_by_id('cn.com.hkgt.gasapp:id/select_btn').click()
    text = driver.find_element_by_id('cn.com.hkgt.gasapp:id/card_type').text
    if text == u'该充值卡状态为:未使用':
        res = u'未使用'
    elif text == u'该充值卡状态为:已使用':
        res = u'已使用'
    else:
        res = u'查询失败'
    Write(Path, Index, j, 1, res)
    j = j+3

print u'已完成！'