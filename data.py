#!/usr/bin/env python3

from os import execl, name
from time import time
from matplotlib.pyplot import draw
import xlrd
import xlwt
import datetime

class WorkBook:
    sheet_data = None
    store_set = set()
    rider_set = set()
    total_order = 0
    cancel_order_id = []
    overtime_order_id = []
    new_data = None
    src_title = ['订单号', '运单号', '商家流水号', '商家名称', '商家ID', \
            '城市', '骑手', '骑手ID', '站点', '站点ID', '商圈', '预订单',\
            '状态', '组织类型', '众包类型', '配送时效', '等待时长', '送达时长', \
            '连击时长', '导航距离', '折线距离', '商家地址', '商家配送评分', \
            '商家配送评价', '订单原价', '订单金额', '折扣金额', '付商家款', \
            '实际付款', '收用户款', '实际收款', '配送费', '下单时间', '支付时间', \
            '期望送达时间', '商家推单时间', '调度时间', '接单时间', '到店时间', '取货时间', \
            '送达时间', '取消时间', '取消原因', '取消操作人', '请退款原因', '申请退款操作人', \
            '驻点订单类型', '业务类型', '是否跨区单', '是否乐跑单', '商户点击出餐时间', '新预计送达时间']

    store_info_title = ["商家名", "商家ID", "订单总数", "完成订单量", "取消订单量", "交易总额"]
    rider_info_title = ['骑手名', '骑手ID', '订单总量', '完成订单量', '取消订单量', '超时订单量', '超时比例']

    def __init__(self, _path):
        workbook = xlrd.open_workbook(_path)
        self.sheet_data= workbook.sheet_by_index(0)
        self.total_order = self.sheet_data.nrows
        print(self.sheet_data.row_values(0))
        for i in range(1, self.total_order):
            self.store_set.add(self.sheet_data.cell_value(i, 3))
            # if self.sheet_data.cell_value(i, 6) == '':
            #     continue
            self.rider_set.add(self.sheet_data.cell_value(i, 6))
            if self.sheet_data.cell_value(i, 12) == "已取消":
                self.cancel_order_id.append(self.sheet_data.cell_value(i, 0))
            else:
                time1 = self.sheet_data.cell_value(i, 34)
                time2 = self.sheet_data.cell_value(i, 40)
                expected_time = datetime.datetime.strptime(time1, '%Y-%m-%d %H:%M:%S')
                actual_time = datetime.datetime.strptime(time2, '%Y-%m-%d %H:%M:%S')
                if actual_time > expected_time and (actual_time - expected_time).seconds / 60 > 8:
                    self.overtime_order_id.append(self.sheet_data.cell_value(i, 0))

    def sort_by_store(self):
        store_sort_index = 3
        self.sort(store_sort_index, self.store_set)

    def sort_by_rider(self):
        rider_sort_index = 6
        self.sort(rider_sort_index, self.rider_set)

    def store_info(self):
        self.new_data = [""] * (len(self.store_set) + 1)
        self.new_data[0] = self.store_info_title
        index = 1
        for name in self.store_set:
            count_order = 0
            store_id = ''
            count_finshed = 0
            count_unfinshed = 0
            amount = 0
            for i in range(1, self.total_order):
                if self.sheet_data.cell_value(i, 3) == name:
                    count_order += 1
                    store_id = self.sheet_data.cell_value(i, 4)
                    amount += float(self.sheet_data.cell_value(i, 25))
                    if self.sheet_data.cell_value(i, 12) != "已取消":
                        count_finshed += 1
                    else:
                        count_unfinshed += 1
            self.new_data[index] = [name, store_id, count_order, count_finshed, count_unfinshed, amount]
            index += 1

    def rider_info(self):
        self.new_data = [""] * (len(self.rider_set) + 1)
        self.new_data[0] = self.rider_info_title
        index = 1
        for name in self.rider_set:
            count_order = 0
            rider_id = ''
            count_finshed = 0
            count_unfinshed = 0
            overtime_order = 0
            overtime_proportion = ""
            for i in range(1, self.total_order):
                if self.sheet_data.cell_value(i, 6) == name:
                    count_order += 1
                    rider_id = self.sheet_data.cell_value(i, 7)
                    if self.sheet_data.cell_value(i, 12) != "已取消":
                        count_finshed += 1
                    else:
                        count_unfinshed += 1

                    if self.sheet_data.cell_value(i, 0) in self.overtime_order_id:
                        overtime_order += 1

            if overtime_order != 0:
                overtime_proportion = str(int((overtime_order / count_finshed) * 100 * 100) / 100) + "%"

            self.new_data[index] = [name, rider_id, count_order, count_finshed, count_unfinshed, overtime_order, overtime_proportion]
            index += 1

    def sort(self, sort_index, name_set):
        self.new_data = [""] * self.total_order
        self.new_data[0] = self.src_title
        data_index = 1
        for name in name_set:
            for i in range(1, self.total_order):
                if self.sheet_data.cell_value(i, sort_index) == name:
                    self.new_data[data_index] = self.sheet_data.row_values(i)
                    data_index += 1

    def cancelled_orders(self):
        self.new_data = [""] * self.total_order
        self.new_data[0] = self.src_title
        data_index = 1
        for i in range(1, self.total_order):
            if self.sheet_data.cell_value(i, 12) == "已取消":
                self.new_data[data_index] = self.sheet_data.row_values(i)
                data_index += 1

    def get_statistical_chart_data(self):
        orders_of_every_hours = [0] * 24
        amount_of_every_hours = [0] * 24
        for i in range(1, self.total_order):
            if self.sheet_data.cell_value(i, 12) != "已取消":
                order_time = self.sheet_data.cell_value(i, 32)
                index = int(datetime.datetime.strptime(order_time, '%Y-%m-%d %H:%M:%S').hour)
                orders_of_every_hours[index] += 1
                amount_of_every_hours[index] += float(self.sheet_data.cell_value(i, 24))

        cancel_order_quantity = len(self.cancel_order_id)
        overtime_order_quantity = len(self.overtime_order_id)
        normal_order_quantity = self.total_order - 1 - cancel_order_quantity - overtime_order_quantity

        dict_of_orders_by_rider = {}
        for rider_name in self.rider_set:
            dict_of_orders_by_rider[rider_name] = 0
            for i in range(1, self.total_order):
                if rider_name == self.sheet_data.cell_value(i, 6) and self.sheet_data.cell_value(i, 12) != "已取消":
                    dict_of_orders_by_rider[rider_name] += 1

        dict_rider_rank = {}
        for order_num in list(set(sorted(dict_of_orders_by_rider.values())))[-5:]:
            dict_rider_rank[str(order_num)] = ''
            for rider_name in dict_of_orders_by_rider:
                if dict_of_orders_by_rider[rider_name] == order_num:
                    dict_rider_rank[str(order_num)] += dict_rider_rank[str(order_num)] + rider_name + '、' 
            dict_rider_rank[str(order_num)] = dict_rider_rank[str(order_num)][:-1]

        dict_store = {}
        for store_name in self.store_set:
            dict_store[store_name] = [0, 0]
            for i in range(1, self.total_order):
                if store_name == self.sheet_data.cell_value(i, 3) and self.sheet_data.cell_value(i, 12) != "已取消":
                    dict_store[store_name][0] += 1
                    dict_store[store_name][1] += self.sheet_data.cell_value(i, 25)
        store_rank = sorted(dict_store.items(), key = lambda i : i[1][1], reverse = True)[0:5]
        return [(orders_of_every_hours, amount_of_every_hours), (normal_order_quantity, overtime_order_quantity, cancel_order_quantity), dict_rider_rank, store_rank]

    def create_workbook(self, flag):
        new_execl = xlwt.Workbook(encoding = 'utf-8')  # 设置文件编码格式
        new_sheet_data = new_execl.add_sheet('My Worksheet')  # 添加sheet页

        style_yellow = xlwt.XFStyle()
        pattern_yellow = xlwt.Pattern()
        pattern_yellow.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern_yellow.pattern_fore_colour = xlwt.Style.colour_map['yellow'] #设置单元格背景色为黄色
        style_yellow.pattern = pattern_yellow

        style_red = xlwt.XFStyle()
        pattern_red = xlwt.Pattern()
        pattern_red.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern_red.pattern_fore_colour = xlwt.Style.colour_map['red'] #设置单元格背景色为黄色
        style_red.pattern = pattern_red

        row_num = 0
        for row_data in self.new_data:
            style = None
            if flag:
                if row_data[0] in self.cancel_order_id:
                    style = style_yellow
                elif row_data[0] in self.overtime_order_id:
                    style = style_red

            col_num = 0
            for data in row_data:
                if style != None:
                    new_sheet_data.write(row_num, col_num, data, style) # 带样式的写入
                else:
                    new_sheet_data.write(row_num, col_num, data) # 不带样式的写入
                col_num += 1
            row_num += 1

        new_execl.save(str(datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")) + ".xls")


import matplotlib.pylab as plt

fig = None

def draw_init():
    global fig
    fig = plt.figure(figsize=(25, 15))
    fig.subplots_adjust(wspace=0.3,hspace=0.3)

def draw_one(_data):
    print(_data)
    x = [i + 0.5 for i in range(24)]
    y1 = _data[0]
    y2 = _data[1]
    ax1 = fig.add_subplot(221)
    x_major_locator = plt.MultipleLocator(1)
    # y_major_locator =plt.MultipleLocator(10)
    ax1.xaxis.set_major_locator(x_major_locator)
    ax1.yaxis.grid(color='b', linestyle='--', linewidth=1, alpha=0.3)
    #plt.legend(loc="upper left") #bbox_to_anchor=(0.6,0.95)
    plt.xticks(rotation = 45)
    ax1.set_ylabel('订单量（条）')
    ax1.set_title("全天订单量与交易额")
    ax1.set_xlim([0, 24])
    ax1.set_xlabel('时间段（小时）')
    ax2 = ax1.twinx()
    ax2.set_ylabel('交易金额（元）')
    l1 = ax1.plot(x, y1, '-*',label='订单量（条）') #alpha=0.5
    l2 = ax2.plot(x, y2,'r-o', label = '交易金额（元）')#alpha=0.5, 
    l = l1 + l2
    labs = [i.get_label()for i in l]
    plt.legend(l, labs, loc="upper right")

def draw_two(_data):
    print(_data)
    labels = ("未超时订单", "超时订单", "取消订单")
    num = list(_data)
    ax = fig.add_subplot(222)
    ax.set_title("订单统计")
    explode = (0, 0, 0.1)
    ax.pie(x = num, labels = labels, explode = explode, autopct = "%1.2f%%")

def draw_three(_data):
    print(_data)
    x = [_data[name] for name in _data]
    y = [int(name) for name in _data]
    colors = ['#98F5FF', '#7FFFD4', '#54FF9F', '#FFF68F', '#FF3030']
    ax = fig.add_subplot(223)
    ax.set_title("骑手送单量排行（前五）")
    _barh = ax.barh(x,y,color = colors)

    for rect in _barh:
        w = rect.get_width()
        ax.text(w, rect.get_y()+rect.get_height()/2, '%d' %
                int(w), ha='left', va='center')

def draw_four(_data):
    print(_data)
    x = [i[0] for i in _data]
    x1 = [i for i in range(5)]
    x2 = [i + 0.45 for i in range(5)]
    temp = [i + 0.45 / 2 for i in range(5)]
    y1 = [i[1][0] for i in _data]
    y2 = [i[1][1] for i in _data]
    ax1 = fig.add_subplot(224)
    plt.xticks(rotation = 15)
    l1 = ax1.bar(x1, y1, 0.45, label='订单量（条）') #alpha=0.5
    ax1.set_xticks(temp)#将坐标设置在指定位置
    ax1.set_xticklabels(x)#将横坐标替换成
    ax1.set_ylabel('订单量（条）')
    ax1.set_title(" 店铺交易额排名（前五）")
    ax2 = ax1.twinx()
    ax2.set_ylabel('交易金额（元）')
    ax2.yaxis.grid(color='b', linestyle='--', linewidth=1, alpha=0.3)
    l2 = ax2.bar(x2, y2, 0.45, color = "red", label='交易金额（元）') #alpha=0.5
    plt.legend([l1,l2],["订单量（条）","交易金额（元）"])

def draw_show():
    plt.show()
def draw_close():
    plt.close()
if __name__ == "__main__":
    my_workbook = WorkBook('20210617_145450_8.xls')
    data_list = my_workbook.get_statistical_chart_data()
    draw_one(data_list[0])
    draw_two(data_list[1])
    draw_three(data_list[2])
    draw_four(data_list[3])
    plt.show()
