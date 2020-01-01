import os
import xlrd
import xlwt
from tkinter import *
import tkinter.messagebox


# Value calculation
def dataAnalysis():
    # Get file path based on the filename provided
    print(os.path.abspath('.'))
    filepath = os.path.abspath('.')
    fpath_input = fpath_value.get()
    filename = filepath + "\\" + fpath_input
    # print(filename)

    # Grab necessary datas from user
    fd = xlrd.open_workbook(filename)
    worksheet = fd.sheet_by_index(int(fsheet_value.get()) - 1)
    nrows = worksheet.nrows
    ftype1 = int(ftype1_value.get()) - 1
    ftype2 = int(ftype2_value.get()) - 1
    fnum = int(fnum_value.get()) - 1

    items = {}
    if fpath_input == '':
        tkinter.messagebox.showinfo("提示", "请输入文件名（包含后缀名）")
    if ftype1_value.get() == '' or ftype2_value.get() == '' or fnum_value.get() == '':
        tkinter.messagebox.showinfo("提示", "请提供必要的数值！！！")

    print(nrows)
    # print(worksheet.cell_value(3 ,8))

    for x in range(2, nrows):
        data = worksheet.cell_value(x, ftype2)
        type = worksheet.cell_value(x, ftype1)
        # print(data)
        if type in items:
            if data in items[type]:
                value = worksheet.cell_value(x, fnum)
                if value != '':
                    if data:
                        if '\u4e00' <= str(value) <= '\u9fff':
                            print("跳过中文类型名")
                            print("")
                        else:
                            print("type: " + type + " " + data)
                            print("length: ", value)
                            # print("A")
                            # print(items[type][data])
                            # print("B")
                            # print(float(value))
                            items[type][data] += float(value)

                            print("count: ", items[type][data])
                            print('')
                    else:
                        print("跳过空元素")
                        print("")
                else:
                    print("跳过空元素")
                    print("")
            else:
                print("type: " + type + " " + data)
                print("length: ", worksheet.cell_value(x, fnum))
                items[type][data] = worksheet.cell_value(x, fnum)
                print("count: ", items[type][data])
                print("")


        else:
            items[type] = {}
            items[type][data] = worksheet.cell_value(x, fnum)
            print("type: " + type + " " + data)
            print("length: ", worksheet.cell_value(x, fnum))
            items[type][data] = worksheet.cell_value(x, fnum)
            print("count: ", items[type][data])
            print("")

    # Result output
    for x in items:
        print("电缆型号: ", x)
        for y in items[x]:
            print("芯数X截面: ", y, " 电缆总长: ", items[x][y], "米")

        print("")

    result_box.insert(INSERT, str(items))


# GUI section
window = Tk(className='简单图表统计器')
window.geometry('400x650')

fpath_label = Label(window, text="文件名：")
fpath_label.grid()
fpath_value = Entry(window, bd=3)
fpath_value.grid(row=0, column=1)

ftype1_label = Label(window, text="母类型列数：")
ftype1_label.grid()
ftype1_value = Entry(window, bd=3)
ftype1_value.grid(row=1, column=1)

ftype2_label = Label(window, text="子类型列数：")
ftype2_label.grid()
ftype2_value = Entry(window, bd=3)
ftype2_value.grid(row=2, column=1)

fnum_label = Label(window, text="统计列数：")
fnum_label.grid()
fnum_value = Entry(window, bd=3)
fnum_value.grid(row=3, column=1)

fsheet_label = Label(window, text="需统计分页：")
fsheet_label.grid()
fsheet_value = Entry(window, bd=3)
fsheet_value.grid(row=4, column=1)

start_button = Button(window, text="开始统计", width=10, command=dataAnalysis)
start_button.grid(row=5, column=1)

res = StringVar()
result_box = Text(window, width=30, height=30, bd=5)
result_box.grid(row=6, column=1)

window.mainloop()


