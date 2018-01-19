#coding=utf-8
def readj(txtP):
    rfile = open(txtP, 'r')
    j = rfile.read()
    rfile.close()
    j = int(j)
    print j
    return j

def writej(txtP, j):
    str_j = str(j)
    wfile = open(txtP, 'w')
    wfile.write(str_j)
    wfile.close()
    print j
        
txtP = 'C:\Users\Liuhaha\Desktop\j.txt'

j = readj(txtP)
while j < 1000:
    j = j + 10
    print j
    writej(txtP, j)