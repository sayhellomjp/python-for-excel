from tkinter import *
import tkinter.filedialog
import tkinter.messagebox
import os
import xlrd, xlwt
from xlutils.copy import copy

# excel样式
alignment = xlwt.Alignment() #创建居中
alignment.horz = xlwt.Alignment.HORZ_CENTER #可取值: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
alignment.vert = xlwt.Alignment.VERT_CENTER # 可取值: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED

font_red = xlwt.Font()
font_red.colour_index = 2

style1 = xlwt.XFStyle() # Create the Style
style1.font = font_red # Apply the Font to the Style
style1.alignment = alignment

style2 = xlwt.easyxf('align: wrap on') # Create the Style
style2.alignment = alignment

# 对比字符串
def check_str(check_length, after, before):
    result = False
    after_len = len(after)
    for i in range(after_len):
        if i <= after_len - check_length:
            last = i + check_length
            check_str = after[i:last]
            if before.find(check_str) != -1:
                # 找到匹配项
                result = True
    return result

book = None
before_book = None

# 上传上周计划
def upload_before():
    selectFileName = tkinter.filedialog.askopenfilename(title='选择文件')
    if selectFileName:
        suffix = os.path.splitext(selectFileName)[1]
        if suffix == '.xls' or suffix == '.xlsx':
            global before_book
            before_book = xlrd.open_workbook(selectFileName)
            e1.delete(0, END)
            e1.insert(0, selectFileName)
        else:
            tkinter.messagebox.showinfo('提示', '请选择xls或xlsx格式的文件')

# 上传本周计划
def upload_after():
    selectFileName = tkinter.filedialog.askopenfilename(title='选择文件')
    if selectFileName:
        suffix = os.path.splitext(selectFileName)[1]
        if suffix == '.xls':
            global book
            book = xlrd.open_workbook(selectFileName, formatting_info=True)
            e2.delete(0, END)
            e2.insert(0, selectFileName)
        else:
            tkinter.messagebox.showinfo('提示', '请选择xls格式的excel，xlsx格式文件可尝试另存为xls格式后再尝试')

# 对比并下载
def download():
    if not before_book:
        tkinter.messagebox.showinfo('提示', '请上传上周计划')
        return
    if not book:
        tkinter.messagebox.showinfo('提示', '请上传本周计划')
        return

    # copy to new file
    save_book = copy(book)
    # get new file sheet 0
    save_sheet = save_book.get_sheet(0)

    # init data
    sh = book.sheet_by_index(0)
    before_sh = before_book.sheet_by_index(0)

    # 列宽
    col1 = save_sheet.col(8)  # 获取第0列
    col1.width = 800 * 20  # 设置第0列的宽为380，高为20

    for r in range(2, sh.nrows):
        after_word = sh.cell_value(r, 4)
        is_find_f = False
        before_checked_word_list = []
        # before cell
        for br in range(2, before_sh.nrows):
            before_word = before_sh.cell_value(br, 4)
            isFind = check_str(4, after_word, before_word)
            if isFind:
                before_checked_word_list.append(before_word)
                is_find_f = True
        if not is_find_f:
            # 未找到匹配项 则为新增标红
            save_sheet.write(r, 4, after_word, style1)
        else:
            # 找到匹配项 得到匹配的内容
            out_value = ''
            for i in range(len(before_checked_word_list)):
                out_value += before_checked_word_list[i]
                if i < len(before_checked_word_list) - 1:
                    out_value += '\r\n'
            save_sheet.write(r, 8, out_value, style2)
    path = tkinter.filedialog.asksaveasfilename(title='保存文件', initialfile='对比结果.xls')
    print(path)
    save_book.save(path)


root = Tk()
root.title('工作计划excel分析程序')
root.resizable(0,0)
root.geometry('500x150')

Label(root, text="上周计划").grid(row=0, ipadx=5, ipady=10)
Label(root, text="本周计划").grid(row=1, ipadx=5, ipady=10)

e1 = Entry(root, width=50)
e1.grid(row=0, column=1)
e2 = Entry(root, width=50)
e2.grid(row=1, column=1)

btn1 = Button(text='上传', command=upload_before).grid(row=0, column=2)
btn2 = Button(text='上传', command=upload_after).grid(row=1, column=2)
btn3 = Button(text='下载对比结果', command=download).grid(row=2, column=0, columnspan=3, ipady=10)

root.mainloop()