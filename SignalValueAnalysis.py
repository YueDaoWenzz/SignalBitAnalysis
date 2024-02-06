import os#用于获取相对路径的OS模块
import xlwings as xw
#信号解析函数
def Bytes2Bits(data,startPos,endPos):
    mask = 0
    len = endPos - startPos + 1
    for i in range(len):
        mask |= 1 << i
    BitsVal = (data & (mask << startPos)) >> startPos
    return BitsVal
path = os.getcwd()#获取相对路径
print("The file path is " + path)
input("Please press any key to continue!")
app = xw.App(visible=True,add_book=False)
wb = xw.Book(os.path.join(path,"SignalAnalysis.xlsx"))
sht = wb.sheets['SignalAyalysis']
PackDataRnge = sht.range('B2').expand('down').value#按列放回列表
StartBitPosRnge = sht.range('D2').expand('down').value
EndBitPosRnge = sht.range('E2').expand('down').value
PackDataNamesRnge = sht.range('A2').expand('down').value
SignalNamesRnge = sht.range('C2').expand('down').value
LstLength = len(PackDataRnge)
signalcnt = LstLength
for i in range(0,LstLength):#去除列表中的空字符
    if -1 == int(PackDataRnge[i]):
        signalcnt = signalcnt - 1#统计有效信号个数
PackDatalst = list(map(int,PackDataRnge))#列表中的数据强制转换成整数型数据
StartBitPoslst = list(map(int,StartBitPosRnge))
EndBitPoslst = list(map(int,EndBitPosRnge))
print("There are " + str(signalcnt) + " signals need be analysised!")
print("Analysis the signals .....please wait!")
SignalValLst = PackDatalst.copy()
for j in range(0,LstLength):
       if -1 == PackDatalst[j]:
            SignalValLst[j] = -1
       else:
            SignalValLst[j] = Bytes2Bits(PackDatalst[j],StartBitPoslst[j],EndBitPoslst[j])#analysis the bits value
print("Write the signal into the Excel.......")
sht.range('F2').options(transpose=True).value = SignalValLst#转置后写入数据
cnt = len(sht.range('F2').expand('down').value)
#填充单元格代码
for k in range(0,cnt):
    if -1 == int(SignalValLst[k]):
        a = 'F' + str(k + 2)
        sht.range(a).color = [0, 0, 255]#蓝色
    elif 1 == int(SignalValLst[k]):
        b = 'F' + str(k + 2)
        sht.range(b).color = [255, 0, 0]#红色     
        print(str(PackDataNamesRnge[k]) + ' ' + '|' + ' ' + str(SignalNamesRnge[k])
              + ' ' + 'maybe occur fault!' +  '\n')
    # elif 0 == int(SignalValLst[k]):
    #     c = 'F' + str(k)
    #     sht.range(c).color = [0, 0, 255]    
for l in range(0,cnt):
    if 0 == int(SignalValLst[l]):
        c = 'F' + str(l + 2)
        sht.range(c).color = [0, 255, 0]#绿色        
wb.save()#保存更改
wb.close()#关闭excel对象
app.quit()#退出程序
print("Signal value write finshed!")
input("Please press any key to exit!")