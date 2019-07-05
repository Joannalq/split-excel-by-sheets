import pandas as pd
import os
import xlrd
import tkinter as tk

root = tk.Tk()
root.title("divide the excel by sheet")
root.geometry('600x200')

#get all excel files under the folder
"""
os.walk(top, topdown=Ture, onerror=None, followlinks=False): get all files
os.listdir()
"""
def collect_excelfiles(folderpath,xtype):
    dirs = os.listdir(folderpath)
    file_name = []
    for filename in dirs:
        #os.path.splitext() divide the filename into file name and extension name
        if os.path.splitext(filename)[1] == xtype:
            file_name.append(filename)
    return file_name

#according to filepath to get new folder name in order to put divided excels into this folder
def newfilename(filepath):
    firstPart = filepath.split("/")
    filename = firstPart[-1]
    newfoldername = filename.split(".")[0]
    return newfoldername

#divide excel based on sheets
def didive_excel_by_sheet(filepath, folderpath):
    #read all sheet
    address = filepath
    contentread = pd.read_excel(address, None)
    names = contentread.keys()

    #set position to save
    #newaddress = "D:/HYGZ-USA MATH/ZCJB-20190620T030028Z-001/ZCJB/85_all/"
    newfoldername = newfilename(filepath)
    newaddress = folderpath+"/"+ newfoldername+"/"
    os.mkdir(newaddress)
    os.chdir(newaddress)
    tdir = newaddress

    #iterate all the sheets
    for name in names:
        tempsheet = pd.read_excel(address, sheet_name = name)
        writer = pd.ExcelWriter(tdir+name+".xlsx")
        tempsheet.to_excel(writer, sheet_name = name, index=False, header=None)
        writer.save()
        writer.close()

def divide(folderpath,xtype):
    #folderpath = "D:/HYGZ-USA MATH/ZCJB-20190620T030028Z-001/ZCJB"
    folderpath = folderpath
    xtype = xtype
    for singlefile in collect_excelfiles(folderpath,xtype):
        filepath = folderpath + "/" + singlefile
        didive_excel_by_sheet(filepath, folderpath)

def run_event(event):
    folderpath = entry_folderpath.get()
    xtype = entry_type.get()
    divide(folderpath,xtype)


if __name__ == "__main__":
    add_label = tk.Label(root, text="please put folder path. eg.D:/HYGZ-USA MATH/ZCJB-20190620T030028Z-001/ZCJB", fg='black')
    add_label.pack(pady=10)
    entry_folderpath = tk.Entry(root, width=44, font=('Arial',8), foreground='gray')
    entry_folderpath.pack()

    type_label = tk.Label(root, text="please put suffix with dot. example: '.xlsx'", fg='black')
    type_label.pack(pady=10)
    entry_type = tk.Entry(root, width=20, font=('Arial',8), foreground='gray')
    entry_type.pack()

    start_button = tk.Button(root, text="start", width=25)
    start_button.pack(pady=20)
    start_button.bind("<Button-1>", run_event)#绑定左键鼠标事件
    root.mainloop()