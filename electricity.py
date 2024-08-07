import openpyxl
from tkinter import messagebox
workbook=openpyxl.load_workbook(r"C:\Users\林庭君\Downloads\台電區處.xlsx")
worksheet=workbook["台電區處"]
def elecAddr(site,constructor,period):
    region=""
    for row in worksheet.iter_rows(min_row=1, max_col=25, max_row=56):
        if(site[0:3]=="臺中市" and site[3:8]!="大甲區日南" and site[3:5]!="霧峰"):
            region="臺中區處"
            return region
        if(site[0:3]==row[0].value):
            if("區處" not in row[0].value and row[1].value!=None):
                for cell in row[2:]:
                    if(site[3:site.index("區")]==cell.value):
                        if(row[1].value=="全部"):
                            region=worksheet[cell.row-(cell.row%3-1)][0].value
                            return region
                        elif(row[1].value=="部分"):
                            messagebox.showinfo("警告",constructor+period+"需要確認台電區處\n請自行填寫檔案'01 併聯登記單(加強電力網)_OK.xlsx'之台電區處")
                            return region
                    else:
                        continue
                    
                
            elif("區處" not in row[0].value and row[1].value==None):
                region=worksheet[row[0].row-(row[0].row%3-1)][0].value
                return region
            else:
                continue
        else:
            continue
