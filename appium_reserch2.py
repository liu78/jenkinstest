#coding=utf-8
from appium import webdriver
import xlrd
import time
from xlutils.copy import copy
import os
import sys
import subprocess

"""def Read_Col(Path,Index,Col):    
    Excel = xlrd.open_workbook(Path)
    Sheet=Excel.sheet_by_index(Index)
    Value= Sheet.col_values(Col)
    return  Value"""

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

def enter_no(str_no):
    driver.find_element_by_id('cn.com.hkgt.gasapp:id/input_card_number').send_keys(str_i)
    driver.find_element_by_id('cn.com.hkgt.gasapp:id/select_btn').click()
    text = driver.find_element_by_id('cn.com.hkgt.gasapp:id/card_type').text
    return text

def readj(txtP):
    rfile = open(txtP, 'r')
    j = rfile.read()
    rfile.close()
    j = int(j)
    print j
    return j

def writej(txtP, row):
    str_j = str(row)
    wfile = open(txtP, 'w')
    wfile.write(str_j)
    wfile.close()
    
desired_caps = {}
desired_caps['platformName'] = 'Android' #设置操作平台
desired_caps['platformVersion'] = '4.4.2' #操作系统版本
desired_caps['deviceName'] = 'emulator-5556' #设备名称
desired_caps['appPackage'] = 'cn.com.hkgt.gasapp' 
desired_caps['udid'] = 'emulator-5556'
desired_caps['appActivity'] = '.AppStartActivity'
desired_caps['connection'] = 'keep-alive'
desired_caps['agent'] = 'false'
try:
	driver = webdriver.Remote('http://127.0.0.1:4730/wd/hub', desired_caps)
	'''
	except:
		print 'server will restart..'
		subprocess.Popen('D:\\2.bat', shell=True) 
		time.sleep(30)
		subprocess.Popen('appium -a 127.0.0.1 -p 4730', shell=True)
		subprocess.Popen('python D:\\workspace\\test\\SINOPEC\\appium_reserch2.py', shell=True) 
	'''	
	time.sleep(20)

	account = 'liuchinho'
	password = 'liu19920723'
	#try验证是否已经登录
	driver.find_element_by_id("cn.com.hkgt.gasapp:id/login_button").click()
	time.sleep(5)
	driver.find_element_by_id("cn.com.hkgt.gasapp:id/login_account").clear()
	time.sleep(5)
	driver.find_element_by_id("cn.com.hkgt.gasapp:id/login_account").send_keys(account)
	time.sleep(5)
	driver.find_element_by_id("cn.com.hkgt.gasapp:id/login_password").send_keys(password)
	time.sleep(5)
	driver.find_element_by_id("cn.com.hkgt.gasapp:id/btn_login_login").click()
	time.sleep(5)
	driver.find_element_by_id('cn.com.hkgt.gasapp:id/traiff_item_chaxun_line_tv').click()
	time.sleep(5)
	driver.find_element_by_id('cn.com.hkgt.gasapp:id/select_card_state').click()
	time.sleep(10)

	txtP = 'D:\g.txt'
	Path = r'D:\4.xls'
	Index = 0
	Col = 0
	v, r = readcol(Path, Index, Col)
	j = readj(txtP)

	for row in range(j, r, 1):
		i = v[row]
		int_i=int(i)
		str_i=str(int_i)
		writej(txtP, row)
		now = time.strftime('%Y-%m-%d  %H:%M:%S', time.localtime(time.time()))
		print now, row
		print str_i
		text = enter_no(str_i)
		'''
		driver.find_element_by_id('cn.com.hkgt.gasapp:id/input_card_number').clear()
		driver.find_element_by_id('cn.com.hkgt.gasapp:id/input_card_number').send_keys(str_i)
		driver.find_element_by_id('cn.com.hkgt.gasapp:id/select_btn').click()
		text = driver.find_element_by_id('cn.com.hkgt.gasapp:id/card_type').text
		'''
		if text == u'该充值卡状态为:未使用':
			res = u'未使用'
		elif text == u'该充值卡状态为:已使用':
			res = u'已使用'
		elif text == u'用户没有登录':
			res = u'正在重新登录'
			driver.find_element_by_id('cn.com.hkgt.gasapp:id/back').click()
			time.sleep(5)
			driver.find_element_by_id("cn.com.hkgt.gasapp:id/login_button").click()
			time.sleep(5)
			driver.find_element_by_id('cn.com.hkgt.gasapp:id/safe').click()
			time.sleep(15)
			driver.find_element_by_id("cn.com.hkgt.gasapp:id/login_account").clear()
			time.sleep(5)
			driver.find_element_by_id("cn.com.hkgt.gasapp:id/login_account").send_keys(account)
			time.sleep(5)
			driver.find_element_by_id("cn.com.hkgt.gasapp:id/login_password").send_keys(password)
			time.sleep(5)
			driver.find_element_by_id("cn.com.hkgt.gasapp:id/btn_login_login").click()
			time.sleep(5)
			driver.find_element_by_id('cn.com.hkgt.gasapp:id/traiff_item_chaxun_line_tv').click()
			time.sleep(5)
			driver.find_element_by_id('cn.com.hkgt.gasapp:id/select_card_state').click()
			time.sleep(10)
			text = enter_no(str_i)
			res = text
		else:
			print u'查询失败'
		Write(Path, Index, row, 1, res)
			
except:
	print 'server will restart in 600s..'
	time.sleep(180)
	'''
	subprocess.Popen('D:\\2.bat', shell=True) 
	time.sleep(30)
	subprocess.Popen('appium -a 127.0.0.1 -p 4730', shell=True)
	time.sleep(30)
	'''
	subprocess.Popen('python D:\\workspace\\test\\SINOPEC\\appium_reserch2.py', shell=True) 

else:
    print u'已完成！'