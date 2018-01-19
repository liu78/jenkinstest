#coding=utf-8
from appium import webdriver
import xlrd
import time
from xlutils.copy import copy
import os
import sys
import subprocess
import configparser

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
    return j

def writej(txtP, row):
    str_j = str(row)
    wfile = open(txtP, 'w')
    wfile.write(str_j)
    wfile.close()
	
def getSize():
    x=driver.get_window_size()['width']
    y=driver.get_window_size()['height']
    return(x,y)

def swipeDown(t):
    l=getSize()
    x1=int(l[0]*0.5)
    y1=int(l[1]*0.90)
    y2=int(l[1]*0.30)
    driver.swipe(x1,y1,x1,y2,t)

prjDir = os.path.split(os.path.realpath(__file__))[0]
configfile_path = os.path.join(prjDir, "config.ini")

cf = configparser.ConfigParser()
cf.read(configfile_path)

platformName = cf.get('devices', 'platformName')
platformVersion = cf.get('devices', 'platformVersion')
deviceName = cf.get('devices', 'deviceName')
appPackage = cf.get('devices', 'appPackage')
udid = cf.get('devices', 'udid')
appActivity = cf.get('devices', 'appActivity')
connection = cf.get('devices', 'connection')
agent = cf.get('devices', 'agent')

desired_caps = {}
desired_caps['platformName'] = platformName #设置操作平台
desired_caps['platformVersion'] = platformVersion #操作系统版本
desired_caps['deviceName'] = deviceName #设备名称
desired_caps['appPackage'] = appPackage
desired_caps['udid'] = udid
desired_caps['appActivity'] = appActivity
desired_caps['connection'] = connection
desired_caps['agent'] = agent
try:
	port = cf.get('appium', 'port')
	ac = cf.get('user', 'account')
	pw = cf.get('user', 'password')
	excel = cf.get('file', 'excel')
	txt = cf.get('file', 'txt')
	step = cf.get('file', 'step')

	aurl = 'http://127.0.0.1:' + port + '/wd/hub'
	driver = webdriver.Remote(aurl, desired_caps)
	time.sleep(20)

	account = ac
	password = pw
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
	time.sleep(30)
	driver.find_element_by_id('cn.com.hkgt.gasapp:id/traiff_item_chaxun_line_tv').click()
	time.sleep(5)
	swipeDown(200)
	time.sleep(10)
	driver.find_element_by_id('cn.com.hkgt.gasapp:id/MyInvoice').click()
	time.sleep(10)

	txtP = txt
	Path = excel
	foot = int(step)
	Index = 0
	Col = 0
	v, r = readcol(Path, Index, Col)
	j = readj(txtP)

	for row in range(j, r, foot):
		i = v[row]
		int_i=int(i)
		str_i=str(int_i)
		writej(txtP, row)
		now = time.strftime('%Y-%m-%d  %H:%M:%S', time.localtime(time.time()))
		print now, row
		print str_i
		text = enter_no(str_i)
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
			swipeDown(200)
			time.sleep(10)
			driver.find_element_by_id('cn.com.hkgt.gasapp:id/MyInvoice').click()
			time.sleep(10)
			text = enter_no(str_i)
			res = text
		else:
			print u'查询失败'
		Write(Path, Index, row, 1, res)

except :
	print 'please wait'
	time.sleep(60)
	subprocess.Popen('D:\\1.bat', shell=True)
	time.sleep(30)
	subprocess.Popen('node D:\\appium\\node_modules\\appium\\bin\\appium.js -a 127.0.0.1 -p 4728', shell=True)
	time.sleep(30)
	subprocess.Popen('python D:\\workspace\\test\\SINOPEC\\appium_reserch.py', creationflags = subprocess.CREATE_NEW_CONSOLE) 
	'''
	print 'server will restart..'
	time.sleep(1000)
	subprocess.Popen('D:\\1.bat', shell=True) 
	time.sleep(30)
	subprocess.Popen('node D:\Appium\node_modules\appium\bin\appium.js -a 127.0.0.1 -p 4728', shell=True)
	time.sleep(30)
	subprocess.Popen('python D:\\workspace\\test\\SINOPEC\\appium_reserch.py', shell=True) 
	'''
else:
	print u'已完成！'