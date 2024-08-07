import openpyxl
from mailmerge import MailMerge  # 引用邮件处理模块
import os
import shutil
import operator
import testimage
import electricity
import cn2an
import pyodbc
import decimal
import datetime

#逆變器型錄、VPC
def invCert(diff_inv_list):
    for inv in range(len(diff_inv_list)):
        srcList=[]
        if(diff_inv_list[inv]["逆變器廠牌"]=="新望"):
            if(diff_inv_list[inv]["逆變器額定輸出功率千瓦"]=="5"):
                srcs = pdf_file_path+"變流器\\"+diff_inv_list[inv]["逆變器廠牌"]+"\\5k\\"
                srcList = os.listdir(srcs)
            elif(diff_inv_list[inv]["逆變器額定輸出功率千瓦"]=="10" or diff_inv_list[inv]["逆變器額定輸出功率千瓦"]=="15"):
                srcs = pdf_file_path+"變流器\\"+diff_inv_list[inv]["逆變器廠牌"]+"\\10k&15k\\"
                srcList = os.listdir(srcs)
            elif(diff_inv_list[inv]["逆變器額定輸出功率千瓦"]=="22"):
                srcs = pdf_file_path+"變流器\\"+diff_inv_list[inv]["逆變器廠牌"]+"\\22k\\"
                srcList = os.listdir(srcs)
            elif(diff_inv_list[inv]["逆變器型號"]=="PV-30000H-U" or diff_inv_list[inv]["逆變器型號"]=="PV-30000S-U"):
                srcs = pdf_file_path+"變流器\\"+diff_inv_list[inv]["逆變器廠牌"]+"\\30k\\"
                srcList = os.listdir(srcs)
                if(diff_inv_list[inv]["逆變器型號"]=="PV-30000S-U"):
                    srcList.remove("型錄-新望-30k-PV-30000H-U-V11-1.pdf")
                elif(diff_inv_list[inv]["逆變器型號"]=="PV-30000H-U"):
                    srcList.remove("型錄-新望-22k-30k-PV-22000S-U，PV-30000S-U-V22-1.pdf")
            elif(diff_inv_list[inv]["逆變器額定輸出功率千瓦"]=="60" or diff_inv_list[inv]["逆變器型號"]=="PV-75000H-U"):
                srcs = pdf_file_path+"變流器\\"+diff_inv_list[inv]["逆變器廠牌"]+"\\60k\\"
                srcList = os.listdir(srcs)
        elif(diff_inv_list[inv]["逆變器廠牌"]=="solaredge"):
            if(diff_inv_list[inv]["逆變器型號"]=="SE33.3K-L" or diff_inv_list[inv]["逆變器型號"]=="SE33.3K-L-TW"):
                srcs = pdf_file_path+"變流器\\"+diff_inv_list[inv]["逆變器廠牌"]+"\\220v\\18.4k\\"
                srcList = os.listdir(srcs)
        elif(diff_inv_list[inv]["逆變器廠牌"]=="Sungrow"):
            if(diff_inv_list[inv]["逆變器型號"]=="30CX-P2" or diff_inv_list[inv]["逆變器型號"]=="50X-P2"):
                srcs = pdf_file_path+"變流器\\"+diff_inv_list[inv]["逆變器廠牌"]+"\\30&50\\"
                srcList = os.listdir(srcs)
                if(diff_inv_list[inv]["逆變器型號"]=="SG30CX-P2"):
                    srcList.remove("VPC_證書_SG50X-P2.pdf")
                elif(diff_inv_list[inv]["逆變器型號"]=="SG50X-P2"):
                    srcList.remove("VPC證書_SG30CX-P2.pdf")
        #elif(diff_inv_list[inv]["逆變器廠牌"]=="HUAWEI"):
        for k in range(len(srcList)) :
            dst = destination+merge_fields["識別碼"]+merge_fields["設置者名稱"]+"\\"
            src=srcs+srcList[k]
            dst=dst+srcList[k]
            shutil.copyfile(src, dst)
#模組型錄、VPC
def panelCert(merge_fields):
    if(merge_fields["模組廠牌"]=="URE"):
        if(int(merge_fields["模組容量瓦"])>=440 and int(merge_fields["模組容量瓦"])<=470):
            srcs = pdf_file_path+"模組\\聯合再生\\440~470w\\"
            srcList = os.listdir(srcs)
            for k in range(len(srcList)) :
                dst = destination+merge_fields["識別碼"]+merge_fields["設置者名稱"]+"\\"
                src=srcs+srcList[k]
                dst=dst+srcList[k]
                shutil.copyfile(src, dst)
    elif(merge_fields["模組廠牌"]=="AUO"):
        srcs = pdf_file_path+"模組\\友達\\"
        srcList = os.listdir(srcs)
        for k in range(len(srcList)) :
            dst = destination+merge_fields["識別碼"]+merge_fields["設置者名稱"]+"\\"
            src=srcs+srcList[k]
            dst=dst+srcList[k]
            shutil.copyfile(src, dst)
    elif(merge_fields["模組廠牌"]=="Anji"):
        srcs = pdf_file_path+"模組\\安集\\"
        srcList = os.listdir(srcs)
        for k in range(len(srcList)) :
            dst = destination+merge_fields["識別碼"]+merge_fields["設置者名稱"]+"\\"
            src=srcs+srcList[k]
            dst=dst+srcList[k]
            shutil.copyfile(src, dst)
#維護者證書
def maintainCert(merge_fields):
    if(merge_fields["維護者"]=="寬福系統有限公司"):
        src=pdf_file_path+r"承裝業會員卡\寬福\電氣承裝業會員卡2024.pdf"
        dst = destination+merge_fields["識別碼"]+merge_fields["設置者名稱"]+"\\電氣承裝業會員卡2024.pdf"
        shutil.copyfile(src, dst)
    elif(merge_fields["維護者"]=="凱強水電有限公司"):
        src=pdf_file_path+r"承裝業會員卡\凱強\凱強水電-甲級承裝業證書.印鑑卡.登記執照(113年度).pdf"
        dst = destination+merge_fields["識別碼"]+merge_fields["設置者名稱"]+"\\凱強水電-甲級承裝業證書.印鑑卡.登記執照(113年度).pdf"
        shutil.copyfile(src, dst)
#技師執照
def techCert(merge_fields):
    if(float(merge_fields["躉售容量"])>=100):
        src=pdf_file_path+r"技師執照\技師執照.pdf"
        dst = destination+merge_fields["識別碼"]+merge_fields["設置者名稱"]+"\\技師執照.pdf"
        shutil.copyfile(src, dst)
    else: 
        return
#公司登記表
def comRegist(merge_fields,dst):
    
    coms = pdf_file_path+"公司登記表"
    comList = os.listdir(coms)
    for i in range(len(comList)):
        if(operator.contains(merge_fields["設置者名稱"],comList[i])):
            srcs = coms+"\\"+comList[i]
            srcList = os.listdir(srcs)
            for j in range(len(srcList)) :
                src=srcs+"\\"+srcList[j]
                shutil.copyfile(src, dst+srcList[j])
#生成檔案
def generFile(merge_fields,diff_inv_list,diff_period_list,period):
    os.chdir(destination)
    os.mkdir(merge_fields["識別碼"]+merge_fields["設置者名稱"])
    os.chdir(destination+merge_fields["識別碼"]+merge_fields["設置者名稱"])
    #invCert(diff_inv_list)
    #panelCert(merge_fields)
    #maintainCert(merge_fields)
    #techCert(merge_fields)
    #comRegist(merge_fields,destination+merge_fields["識別碼"]+merge_fields["設置者名稱"]+"\\")
    for file in range(len(files)):
        wordname=origin_file_path+files[file]
        mergeFile = MailMerge(wordname)
        if(files[file]=="03台電並聯圖說.docx"):
            os.chdir(destination+merge_fields["識別碼"]+merge_fields["設置者名稱"])
            mergeFile.merge_rows('逆變器型號',diff_inv_list)
            mergeFile.choice(merge_fields)
            #只有一期
            if(period==1 and len(diff_period_list)==len(diff_inv_list)):
                mergeFile.remove_period()
            if merge_fields["期別"]!="一期":
                mergeFile.for_short_multi(diff_period_list,period)
            else:
                mergeFile.for_short_multi(diff_period_list,period)
                mergeFile.remove_short_first()
            
            mergeFile.merge(**merge_fields)
            mergeFile.write(files[file])#存在cmd當前路徑下
            #testimage.insertImg(destination+merge_fields["識別碼"]+merge_fields["設置者名稱"]+"\\03台電並聯圖說.docx")
        elif(files[file]=="申請表1122.docx"):
            os.chdir(destination+merge_fields["識別碼"]+merge_fields["設置者名稱"])
            mergeFile.choice(merge_fields)
            mergeFile.count_sell(diff_period_list,period)
            mergeFile.merge(**merge_fields)
            mergeFile.write(files[file])
        elif(files[file] in agreement_files):
            os.chdir(destination+merge_fields["識別碼"]+merge_fields["設置者名稱"])
            if(not(os.path.isdir("同意備案"))):
                os.mkdir("同意備案")
            os.chdir(destination+merge_fields["識別碼"]+merge_fields["設置者名稱"]+'\\同意備案')
            mergeFile.choice(merge_fields)
            mergeFile.judge(merge_fields)
            mergeFile.merge(**merge_fields)
            mergeFile.write(files[file])#存在cmd當前路徑下
        elif(files[file] in device_files):
            os.chdir(destination+merge_fields["識別碼"]+merge_fields["設置者名稱"])
            if(not(os.path.isdir("設備登記"))):
                os.mkdir("設備登記")
            os.chdir(destination+merge_fields["識別碼"]+merge_fields["設置者名稱"]+'\\設備登記')
            mergeFile.field_concate("逆變器廠牌",diff_inv_list,"/")
            mergeFile.field_concate("逆變器型號",diff_inv_list,"/")
            mergeFile.field_concate("逆變器額定輸出功率瓦",diff_inv_list,"W/")
            mergeFile.field_concate("逆變器輸出最大電流",diff_inv_list,"A/")
            mergeFile.field_concate("逆變器數量",diff_inv_list,"台/")
            mergeFile.p_concate("逆變器PV輸入操作電壓最低",diff_inv_list,"/")
            mergeFile.choice(merge_fields)
            mergeFile.merge(**merge_fields)
            mergeFile.write(files[file])#存在cmd當前路徑下
        elif(files[file]=="01 併聯登記單(加強電力網)_OK.xlsx"):
            workbook=openpyxl.load_workbook(origin_file_path+files[file])
            worksheet=workbook["併連躉售"]
            worksheet["F8"]=merge_fields["設置者名稱"]
            if(merge_fields["設置地址"]!=""):
                worksheet["F10"]=merge_fields["設置地址"]
            else:
                worksheet["F10"]=merge_fields["設置地號"]
            worksheet["F6"]="台灣電力公司"+electricity.elecAddr(merge_fields["設置地號"],merge_fields["設置者名稱"],merge_fields["期別"])[0:3]+"營業處"
            worksheet["AB9"]=merge_fields["維護者地址"]
            worksheet["AM7"]=merge_fields["維護者電話"]
            worksheet["F12"]=merge_fields["新設或增設"]+"太陽光電系統"+merge_fields["躉售容量"]+"kWp\n"+merge_fields["逆變器規格"]+"一套，併"+merge_fields["併聯方式"]+"，"+merge_fields["躉售方式"]+"躉售"
            workbook.save(filename=destination+merge_fields["識別碼"]+merge_fields["設置者名稱"]+"\\"+files[file])
        else:
            os.chdir(destination+merge_fields["識別碼"]+merge_fields["設置者名稱"])
            mergeFile.choice(merge_fields)
            mergeFile.merge(**merge_fields)
            mergeFile.write(files[file])#存在cmd當前路徑下
        #if(os.path.isdir("設備登記")):
            #comRegist(merge_fields,dst=destination+merge_fields["識別碼"]+merge_fields["設置者名稱"]+'\\設備登記\\')
origin_file_path = "C:\\Users\\林庭君\\Downloads\\程式\\原始資料區\\"
pdf_file_path="Y:\\tsaifan\\送件資料區\\"
files = os.listdir(origin_file_path)
diff_inv_list=[]#存同個設置者要合併的資料
diff_period_list=[]
case_data_list=[]#案場資料查詢
capacity_data_list=[]#躉售容量查詢
destination="C:\\Users\\林庭君\\Downloads\\程式\\生成的檔案\\"
sell=0
agreement_files=["第7條附件1-再生能源發電設備同意備案申請表(113.01.04).docx","04足資辨識設置場址及位置照片20240419.docx"]
device_files=["02 再生能源發電設備設置聲明書.docx","02-2 竣工試驗報告.docx","04-2 採購委託書(發票不是開給設置者才需要).docx","第11條附件4-再生能源發電設備設備登記申請表及設置聲明書 (113.01.04).docx","第11條附件4-範例1再生能源發電設備完工照片(113.01.04).docx","第11條附件4-範例2再生能源發電設備支出憑證(113.01.04).docx"]

conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\林庭君\Downloads\程式\正式.accdb')
cursor = conn.cursor()
cursor.execute('SELECT * FROM "案場資料查詢"')

record=cursor.fetchall()
record.insert(0,())
for row in cursor.columns(table='案場資料查詢'):
    record[0]+=(row.column_name,)
for i in range(1,len(record)):
    data_dict={}
    for j in range(len(record[0])):
        data_dict[record[0][j]]=record[i][j]
    case_data_list.append(data_dict)
'''
cursor.execute('SELECT * FROM "躉售容量查詢"')
record=cursor.fetchall()
record.insert(0,())
for row in cursor.columns(table='躉售容量查詢'):
    record[0]+=(row.column_name,)
for i in range(1,len(record)):
    data_dict={}
    for j in range(len(record[0])):
        data_dict[record[0][j]]=record[i][j]
    capacity_data_list.append(data_dict)
'''
cursor.close()
conn.close()

        
for i in range(len(case_data_list)):  # 循環逐行打印
    period_dict={}
    period=1
    #若上面改成 table[i][column].value，為何下面都印出相同的東西 
    #print(table[i][3])
    #創建一个空字典来存储键值对参数
    merge_fields = {}
    for column in list(case_data_list[i]):
        #將小數點後是0的去掉0
        if(type(case_data_list[i][column]) == int or type(case_data_list[i][column]) == float or type(case_data_list[i][column]) ==decimal.Decimal):
            if(case_data_list[i][column]%1==0):
                case_data_list[i][column] = str(int(case_data_list[i][column])) 
                merge_fields[column]=case_data_list[i][column]
            else:
                case_data_list[i][column] = str(float(case_data_list[i][column]))
                merge_fields[column]=case_data_list[i][column]
        else:
            if(column=="簽約日期" or column=="併網日期" and case_data_list[i][column] is not None):
                case_data_list[i][column] = str(case_data_list[i][column]).strip(" 00:00:00")
            case_data_list[i][column] = str(case_data_list[i][column])
            merge_fields[column]=case_data_list[i][column]
    #處理diff_inv
    diff_inv_list.append(merge_fields)
    if(case_data_list[i]["是否要匯出"]=="True"):
        try:
            if(str(case_data_list[i+1]["識別碼"])==diff_inv_list[0]["識別碼"]):
                continue
            else:
                #處理diff_period
                #二期以上
                period=cn2an.cn2an(case_data_list[i]["期別"][0])
                for repeat_row in range(len(case_data_list)):
                    # 設置者名稱跟地號相同
                    if(case_data_list[repeat_row]["設置者名稱"]==diff_inv_list[0]["設置者名稱"] and case_data_list[repeat_row]["設置地址"]==diff_inv_list[0]["設置地址"]):
                        period_dict={}
                        for column in list(case_data_list[repeat_row]):
                            period_dict[column]=case_data_list[repeat_row][column]
                        diff_period_list.append(period_dict)
                generFile(merge_fields,diff_inv_list,diff_period_list,period)
                sell=0
                diff_inv_list=[]
                diff_period_list=[]        #最後一筆
        except:
            period=cn2an.cn2an(case_data_list[i]["期別"][0])
            for repeat_row in range(len(case_data_list)):
                # 設置者名稱跟地號相同
                if(case_data_list[repeat_row]["設置者名稱"]==diff_inv_list[0]["設置者名稱"] and case_data_list[repeat_row]["設置地址"]==diff_inv_list[0]["設置地址"]):
                    period_dict={}
                    for column in list(case_data_list[repeat_row]):
                        period_dict[column]=case_data_list[repeat_row][column]
                    diff_period_list.append(period_dict)
            generFile(merge_fields,diff_inv_list,diff_period_list,period)
            sell=0
            diff_inv_list=[]
            diff_period_list=[]
    else:
        diff_inv_list=[]
        continue