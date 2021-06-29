import tkinter as tk
from tkinter.filedialog import askopenfilename
import data 

my_workbook = None

def select_path():
    global my_workbook
    _path = askopenfilename()
    if len(_path) == 0:
        print("path is null !!!")
    else:
        path.set(_path)
        path_str = path.get()
        if path_str == "" or len(path_str) < 5:
            label2_text.set("未选择Excel文件")
        elif path_str[-4:] != ".xls" and path_str[-5:] != ".xlsx":
            label2_text.set("选择文件类型错误")
        else:
            label2_text.set("文件导入成功")
            my_workbook = t.WorkBook(_path)
            print(path.get())

def statistical_chart():
    if path.get() != "":
        data_list = my_workbook.get_statistical_chart_data()
        data.draw_init()
        data.draw_one(data_list[0])
        data.draw_two(data_list[1])
        data.draw_three(data_list[2])
        data.draw_four(data_list[3])
        data.draw_show()
        label2_text.set("文件处理完成")
    else:
        label2_text.set("未选择Excel文件")



def _sort_by_store():
    if my_workbook == None:
        label2_text.set("未选择Excel文件")
        return
    my_workbook.sort_by_store()
    my_workbook.create_workbook(True)
    label2_text.set("按店铺排序导出完成")

def _sort_by_rider():
    if my_workbook == None:
        label2_text.set("未选择Excel文件")
        return
    my_workbook.sort_by_rider()
    my_workbook.create_workbook(True)
    label2_text.set("按骑手排序导出完成")

def _store_info():
    if my_workbook == None:
        label2_text.set("未选择Excel文件")
        return
    my_workbook.store_info()
    my_workbook.create_workbook(False)
    label2_text.set("店铺信息导出完成")

def _rider_info():
    if my_workbook == None:
        label2_text.set("未选择Excel文件")
        return
    my_workbook.rider_info()
    my_workbook.create_workbook(False)
    label2_text.set("骑手信息导出完成")

def _cancelled_orders():
    if my_workbook == None:
        label2_text.set("未选择Excel文件")
        return
    my_workbook.cancelled_orders()
    my_workbook.create_workbook(False)
    label2_text.set("已取消订单导出完成")


root_window = tk.Tk()
root_window.geometry("260x350")
root_window.resizable(0, 0)
root_window.title('Beta V0.1')
path = tk.StringVar()

label1 = tk.Label(root_window, text = "文件路径:")
label1.grid(row = 0, column = 0)

entry1 = tk.Entry(root_window, textvariable = path)
entry1.grid(row = 0, column = 1)

button1 = tk.Button(root_window, text = "文件选择", command = select_path)
button1.grid(row = 0, column = 2)

button2 = tk.Button(root_window, text = "统计表格",  width=15, height=2, command = statistical_chart)
button2.grid(row = 2, column = 1)

button4 = tk.Button(root_window, text = "店铺名归类导出", width=15, height=2, command = _sort_by_store)
button4.grid(row = 3, column = 1)

button5 = tk.Button(root_window, text = "骑手名归类导出",  width=15, height=2,command = _sort_by_rider)
button5.grid(row = 4, column = 1)

button6 = tk.Button(root_window, text = "已取消订单信息导出",  width=15, height=2,command = _cancelled_orders)
button6.grid(row = 5, column = 1)

button6 = tk.Button(root_window, text = "店铺生产信息导出",  width=15, height=2,command = _store_info)
button6.grid(row = 6, column = 1)

button6 = tk.Button(root_window, text = "骑手派单信息导出",  width=15, height=2,command = _rider_info)
button6.grid(row = 7, column = 1)

#处理会改变的文字
label2_text = tk.StringVar()
label2_text.set("未选择Excel文件")
label2 = tk.Label(root_window, textvariable = label2_text)
label2.grid(row = 1, column = 1)

root_window.mainloop()



