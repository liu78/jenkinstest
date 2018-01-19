#coding=utf-8
import xlrd, xlwt
from xlutils.copy import copy
import os

def Read_Col(Path,Index,Col):    
    Excel = xlrd.open_workbook(Path)
    Sheet=Excel.sheet_by_index(Index)
    value= Sheet.col_values(Col)
    rows=Sheet.nrows
    return  value, rows

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
    
Path = 'C:\Users\Liuhaha\Desktop\gas.xlsx'
Index = 0
Col = 0
v, r = Read_Col(Path, Index, Col)
print v
e = Exchange_Str(v)
print e
print r

for i in v:
    type1=type(i)
    int_i=int(i)
    str_i=str(int_i)
    print str_i
    print range(2)
    j = 0
    Write('C:\Users\Liuhaha\Desktop\gas.xlsx', 0, j, 1, str_i)
    j+=1