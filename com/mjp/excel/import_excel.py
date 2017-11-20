import xlrd, xlwt
from xlutils.copy import copy

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

# alignment
alignment = xlwt.Alignment() #创建居中
alignment.horz = xlwt.Alignment.HORZ_CENTER #可取值: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
alignment.vert = xlwt.Alignment.VERT_CENTER # 可取值: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED

# font style
font_red = xlwt.Font()
font_red.colour_index = 2

style1 = xlwt.XFStyle() # Create the Style
style1.font = font_red # Apply the Font to the Style
style1.alignment = alignment

style2 = xlwt.easyxf('align: wrap on') # Create the Style
style2.alignment = alignment

# read origin data
book = xlrd.open_workbook("D:/rtt/after111.xls", formatting_info=True)
before_book = xlrd.open_workbook("D:/rtt/before111.xls", formatting_info=True)
# copy to new file
save_book = copy(book)
# get new file sheet 0
save_sheet = save_book.get_sheet(0)

# init data
sh = book.sheet_by_index(0)
before_sh = before_book.sheet_by_index(0)

# 列宽
col1=save_sheet.col(8) #获取第0列
col1.width=800*20 #设置第0列的宽为380，高为20

list = []
show_times = 0
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

    # print(after_word, is_find_f)

save_book.save('D:/rtt/ttt.xls')

# check_str(3, '查明爱科特公司回函情况，形成书面说明报办公室', None)

