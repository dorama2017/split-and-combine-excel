import tkinter as tk
from tkinter import filedialog

# 创建窗口对象
window = tk.Tk()
window.title("excel小程序")
window.geometry('500x100')


import xlwings as xw
import os

def combineexcel(path):
    if len(path) > 0 :
        app = xw.App(visible=True, add_book=False)
        files = [i  for root, dirs, files in os.walk(path, topdown=False) for i in files]
        #print(files)
        new_workbook = app.books.add()
        for i in range(len(files)):
            if files[i].endswith(".xls") or files[i].endswith(".xlsx"):
                wbook02 = app.books.open(path+'\\'+files[i])
                sheet = wbook02.sheets[0]
               # for sheet in wbook02.sheets:
              #  print(wbook02.sheets)
             #   if len(wbook02.sheets) ==1:
              #  sheet  =wbook02.sheets[0]
                if files[i].endswith(".xls"):
                    sheet.name = files[i].replace('.xls','')
                if files[i].endswith(".xls"):
                    sheet.name = files[i].replace('.xlsx','')
                 #   print(new_workbook.sheets)
                sheet.copy(before=new_workbook.sheets[0])
                wbook02.close()


        new_workbook.save('{}\\{}.xlsx'.format(path,'合并'))  # 保存新工作簿
        new_workbook.close()  # 关闭新建工作簿
        app.quit()  # 退出Excel程序

def splitexcel(file_path):
   # file_path =
   if len(file_path) > 0 :
        app = xw.App(visible=True, add_book=False)
        workbook = app.books.open(file_path)
        worksheet = workbook.sheets
        for i in worksheet:  # 遍历工作簿中所有工作表
            new_workbook = app.books.add()  # 新建工作簿
            new_worksheet = new_workbook.sheets[0]  # 选中新建工作簿中的第1张工作表
            i.copy(before=new_worksheet)  # 将原来工作簿中的当前工作表复制到新建工作簿的第1张工作表之前if __name__ == '__main__':

            path = os.path.dirname(file_path)
            if os.path.exists('{}\拆分后'.format(path)) == False:
                os.mkdir('{}\拆分后'.format(path))
       #     print('{}\拆分后\{}.xlsx'.format(path,i.name))
            new_workbook.save('{}\拆分后\{}.xlsx'.format(path,i.name))  # 保存新工作簿
            new_workbook.close()  # 关闭新建工作簿
        app.quit()  # 退出Excel程序
def select_folder():
    # 打开文件选择对话框
    folderpath = filedialog.askdirectory()
   # print("选择的文件夹：", folderpath)
    combineexcel(folderpath)
def select_file():
    # 打开文件选择对话
    filepath = filedialog.askopenfilename()
  #  print(filepath)
    splitexcel(filepath)



Label= tk.Label(window, text="合并文件夹中的所有xls或者xlsx文件的第一个sheet为1个表格!")
Label.pack()
select_button = tk.Button(window, text="合并表格，请选择文件夹", command=select_folder)
select_button.pack()


Label= tk.Label(window, text="分割一个表格的多个sheet,按sheet名字保存!")
Label.pack()
select_button = tk.Button(window, text="分割表格，请选择文件", command=select_file)
select_button.pack()
# 启动主循环
window.mainloop()
