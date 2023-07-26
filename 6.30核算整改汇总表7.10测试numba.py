#coding:utf-8
import zipfile
####################################################抬头
import threading
import pandas as pd
import numpy as np
import time
import os
import xlsxwriter
from pathlib import Path
import warnings
import tkinter as tk
import tkinter.messagebox #弹出框
from tkinter import *
from PIL import Image, ImageTk, ImageSequence
import xlsxwriter
import traceback
import datetime
from datetime import datetime
import numba
from numba import jit
warnings.filterwarnings('ignore')
#print('请确认相關数据源是否都已放入对应文件夹，程序执行完成会自动关闭，请耐心等待')

###自定义函数
def vlookup(find_col, start_col, object_col):
    '''类似EXCEL的VLOOKUP,未找到返回空文本
    find_col:要查找的序列，需要是文本。其中每个值都将返回一个对象，
    start_col:起始列,不能含有空值
    object_col:目标列
    '''
    temp = []
    find_col = [str.strip(str.upper(i)) for i in find_col]
    startlist = [str.strip(str.upper(i)) for i in start_col]
    for item in find_col:
        if item in startlist:
            num = startlist.index(item)
            temp.append(object_col[num])
        else:
            temp.append('')
    return temp
###自定义函数
####################################################软件封面设计
window = tk.Tk()
window.title('项目成本汇总工具6.30-1')  # 窗口的标题
window.geometry('600x900')  # 窗口的大小
window.iconbitmap('软件附带文件\头像.ico')
frame=tk.Canvas(window,width=620,height=800,background='silver',scrollregion=(0,0,1500,1000))
roll=Scrollbar(window,orient='vertical',command=frame.yview)
#frame.pack(fill="both",side='right')
frame['yscrollcommand']=roll.set
roll.pack(side=RIGHT, fill=Y)
frame.pack(side=TOP, fill=Y, expand=True)
image_file = ImageTk.PhotoImage(file=r'软件附带文件\背景.jpg')
image =frame.create_image(300, 0, anchor='n', image=image_file)
image1 =frame.create_image(300, 450, anchor='n', image=image_file)
image_file1 = ImageTk.PhotoImage(file=r'软件附带文件\海目星.jpg')
image_new =frame.create_image(50, 10, anchor='n', image=image_file1)

lab_choice = tk.Label(frame,
             text='请确认历史PO剔除起始时间：',  # 标签的文字
             bg='indianred',  # 标签背景颜色
             font=('华文行楷', 12),  # 字体和字体大小
             width=25, height=1,  # 标签长宽(以字符长度计算)
             padx=2,pady=4,anchor='s',fg='black')
frame.create_window(150,55,window=lab_choice)
now_day1= (datetime.strptime(time.strftime("%Y-%m",time.localtime(time.time()))+'-01','%Y-%m-%d')-pd.to_timedelta(63,unit='D')).strftime("%Y-%m")
now_day1=now_day1+'-01'
caldate1= tk.StringVar(value=now_day1)
po_time=tk.Entry(frame,show=None,width=10,bd=4,cursor='cross',textvariable=caldate1)
frame.create_window(300,55,window=po_time)
lab_choice1 = tk.Label(frame,
             text='请确认历史料剔除起始时间：',  # 标签的文字
             bg='lightgreen',  # 标签背景颜色
             font=('华文行楷', 12),  # 字体和字体大小
             width=25, height=1,  # 标签长宽(以字符长度计算)
             padx=2,pady=4,anchor='s',fg='black')
frame.create_window(150,95,window=lab_choice1)
now_day2= (datetime.strptime(time.strftime("%Y-%m",time.localtime(time.time()))+'-01','%Y-%m-%d')-pd.to_timedelta(33,unit='D')).strftime("%Y-%m")
now_day2=now_day2+'-01'
caldate2= tk.StringVar(value=now_day2)
mater_time=tk.Entry(frame,show=None,width=10,bd=4,cursor='cross',textvariable=caldate2)
frame.create_window(300,95,window=mater_time)

lab_choice2= tk.Label(frame,
             text='请确认历史工剔除起始时间：',  # 标签的文字
             bg='khaki',  # 标签背景颜色
             font=('华文行楷', 12),  # 字体和字体大小
             width=25, height=1,  # 标签长宽(以字符长度计算)
             padx=2,pady=4,anchor='s',fg='black')
frame.create_window(150,135,window=lab_choice2)
now_day3= (datetime.strptime(time.strftime("%Y-%m",time.localtime(time.time()))+'-01','%Y-%m-%d')-pd.to_timedelta(63,unit='D')).strftime("%Y-%m")
now_day3=now_day3+'-01'
caldate3= tk.StringVar(value=now_day3)
work_time=tk.Entry(frame,show=None,width=10,bd=4,cursor='cross',textvariable=caldate3)
frame.create_window(300,135,window=work_time)

lab_choice3= tk.Label(frame,
             text='请确认历史费剔除起始时间：',  # 标签的文字
             bg='lightskyblue',  # 标签背景颜色
             font=('华文行楷', 12),  # 字体和字体大小
             width=25, height=1,  # 标签长宽(以字符长度计算)
             padx=2,pady=4,anchor='s',fg='black')
frame.create_window(150,175,window=lab_choice3)
now_day4= (datetime.strptime(time.strftime("%Y-%m",time.localtime(time.time()))+'-01','%Y-%m-%d')).strftime("%Y-%m-%d")
caldate4= tk.StringVar(value=now_day4)
cost_time=tk.Entry(frame,show=None,width=10,bd=4,cursor='cross',textvariable=caldate4)
frame.create_window(300,175,window=cost_time)
button_execute= tk.Button(frame, text='执行', width=4, height=1, fg='darkred',bd=6,font=('华文行楷', 18))
frame.create_window(180,230,window=button_execute)
time_ = time.time()
 #这个事件是鼠标进入组件，用什么事件
def pick():
    global a,flag
    while 1<2:
        im = Image.open(r'软件附带文件\机器.gif')
        # GIF图片流的迭代器
        iter = ImageSequence.Iterator(im)
        #frame就是gif的每一帧，转换一下格式就能显示了
        for jp in iter:
            pic=ImageTk.PhotoImage(jp)
            pic_act=frame.create_image((470,130), image=pic)
            time.sleep(0.1)
            window.update_idletasks()  #刷新
            window.update()
def t1():
    frame.bind("<Enter>",pick)

@jit(nopython=True) # jit，numba装饰器中的一种10^5
def execute():
    time_start = time.time()
    image_fil= ImageTk.PhotoImage(file=r'软件附带文件\加.jpg')
    window.update()
    #frame.create_image(420, 100,image=image_fil)
    screm = tk.Text(frame, bg='white',  # 标签背景颜色
                    font=('微软雅黑', 12),  # 字体和字体大小
                    width=60, height=35,  # 标签长宽(以字符长度计算)
                    )
    frame.create_window(300, 620, window=screm)
####################################################第一阶段
    screm.insert(INSERT, '\n请确认相關数据源是否都已放入对应文件夹', '\n')
    window.update()
    try:
        screm.insert(INSERT, '\n一、第一阶段--读取核算相关明细数据源并生成四个表明细数据', '\n')
        window.update()
##################################################################################################################################################################读取核算底表
        screm.insert(INSERT, '\n1.1：正在读取核算底表...', '\n')
        window.update()
        Path_base = r'数据源\核算底表'
        filename_base = os.listdir(Path_base)
        for i in range(len(filename_base)):
            if str(filename_base[i]).count('~$') == 0:
                # order_new = pd.read_excel(filePathT_1 + '/' + str(file_name1[i]), header=2)
                item = pd.read_excel(Path_base + '/' + str(filename_base[i]))
        item_str=["序列号","区域","行业中心","设备类型","客户简称","大项目名称","大项目号","产品线名称",'产品线编码',"核算项目号","设备名称","生产状态","一般工单号601/608","返工工单号603","项目号整理","成品料号","是否预验收","全面预算有无",'自制/外包',"OA状态","项目财经","项目财经再分类"]
        item[item_str] = item[item_str].fillna('')
        item=item[item['设备类型'].str.contains('增值改造|纯人力')==False].reset_index(drop=True)
        item[["项目数量","已出货数量","在产数量","集团收入","软件收入","硬件收入"]]=item[["项目数量","已出货数量","在产数量","集团收入","软件收入","硬件收入"]].fillna(0)
        default_date1 = pd.Timestamp(2090, 1, 1)
        item[["工单开立时间","工单完工时间","系统出货时间","实际出货时间","系统验收时间","实际验收时间"]]=item[["工单开立时间","工单完工时间","系统出货时间","实际出货时间","系统验收时间","实际验收时间"]].fillna(default_date1)
        item['工单开立时间'] = pd.to_datetime(item['工单开立时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        item['工单完工时间'] = pd.to_datetime(item['工单完工时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        item['系统出货时间'] = pd.to_datetime(item['系统出货时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        item['实际出货时间'] = pd.to_datetime(item['实际出货时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        item['系统验收时间'] = pd.to_datetime(item['系统验收时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        item['实际验收时间'] = pd.to_datetime(item['实际验收时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        item['工单开立时间'] = pd.to_datetime(item['工单开立时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        item['工单开立时间'] = ['' if i == '2090-01-01' else i for i in item['工单开立时间']]
        item['工单完工时间'] = ['' if i == '2090-01-01' else i for i in item['工单完工时间']]
        item['系统出货时间'] = ['' if i == '2090-01-01' else i for i in item['系统出货时间']]
        item['系统验收时间'] = ['' if i == '2090-01-01' else i for i in item['系统验收时间']]
        item['实际出货时间'] = ['' if i == '2090-01-01' else i for i in item['实际出货时间']]
        item['实际验收时间'] = ['' if i == '2090-01-01' else i for i in item['实际验收时间']]

        item=item[["序列号","区域","行业中心","设备类型","客户简称","大项目名称","大项目号","产品线名称",'产品线编码',"核算项目号","设备名称","项目数量","已出货数量","在产数量","生产状态","集团收入","软件收入","硬件收入","一般工单号601/608","工单开立时间","工单完工时间","系统出货时间","实际出货时间","返工工单号603","系统验收时间","实际验收时间","项目号整理","成品料号","是否预验收","全面预算有无","OA状态",'自制/外包',"项目财经","项目财经再分类"]]
        report_item = item.drop_duplicates(subset=['项目号整理']).reset_index(drop=True)[["项目号整理", "核算项目号",  "大项目号", "大项目名称", "设备名称", "产品线名称",'产品线编码'
            ,'客户简称','自制/外包', '项目财经', '项目财经再分类', '成品料号']]
        report_item_cal_out = item.copy()
        time_base=time.time()

##################################################################################################################################################################读取采购PO
        screm.insert(INSERT, '\n1.2：采购PO...', '\n')
        window.update()
        screm.insert(INSERT, '\n1.2.1：正在读取采购PO...', '\n')
        window.update()
        dir = r"数据源\采购明细-数据源"
        if os.listdir(dir):
            index = 0
            report2 = []
            filenames = os.listdir(dir)
            for i in filenames:
                if '~$' in i:
                    filenames.remove(i)
            for name in filenames:
                if 'csv' in name and '~$' not in name:
                    with open(os.path.join(dir, name), encoding='utf-8', errors='ignore') as f:
                        df3 = pd.read_csv(f,  sep=',', error_bad_lines=False, low_memory=False,thousands=',')  #, header=2
                    report2.append(df3)
                    index += 1  # 为了查看合并到第几个表格了
            report002 = pd.concat(report2)
#######################
            screm.insert(INSERT, '\n1.2.2：正在处理采购PO格式...', '\n')
            window.update()
            default_date = pd.Timestamp(1990, 1, 1)
            report002['采购日期']=report002['采购日期'].fillna(default_date)
            report002['采购日期']=pd.to_datetime(report002['采购日期'], errors='coerce')
            report002 = report002.reset_index(drop=True)
            report002['数据-来源'] = '新增数据'
            if "状态" in report002.columns:
                del report002["状态"]
            if "备注" in report002.columns:
                del report002["备注"]
            col1=['产能','工艺']
            for i in col1:
                if i not in report002.columns:
                    report002[i]=''
            report002 = report002.rename(columns={'采购单据状态': '状态', '采购行备注': '备注','采购单采购人员':'采购人员'})
            report002 = report002.rename(columns={'采购单据状态': '状态'})
            report002 = report002.rename(columns={'请购单号': '来源单号'})
            report002 = report002.rename(columns={'产品线': '产品线编码','大项目':'大项目号'})
            need_col = ["采购单号", "采购日期", "采购单类型", "料件编号", "品名", "规格", "采购数量", "单价", "含税金额",'零件类型','模组名称'
                , "税率", "已收货量", "未交量", "备注", "项目编号"
                ,  "采购供应商", "库存管理特征", "行状态", "状态", "作业编号"
                , "品牌","验退量", "仓退换货量", "仓退量", "留置原因说明"
                , "采购部门", '采购人员', "料件分类", "来源单号",'大项目号','大项目名称','产品线编码','产品线名称','数据-来源','产能','工艺']
            # 删除不需要的字段
            for col in report002.columns:
                if col not in need_col:
                    del report002[col]
            # 新增字段并重置列顺序
            report002 = report002.reindex(columns=[ "核算项目号", "采购单号", "采购日期", "采购单类型", "料件编号", "品名", "规格", "采购数量"
                , "未税单价", "采购金额-未税", "单价", "含税金额", "税率", "已收货量"
                , "未交量", "备注", "项目编号", "采购人员", "采购供应商", "库存管理特征"
                , "行状态", "状态", "作业编号", "模组标识", "模组名称", "品牌", "零件类型"
                , "产能", "标准/非标", "客户简称", "大项目名称", "大项目号", "设备名称"
                , "产品线编码", "产品线名称", "工艺", "验退量", "仓退换货量", "仓退量"
                , "项目号整理", "留置原因说明", "采购部门", "料件分类", "来源单号"
                , "数据-来源", "数据分布", "项目财经"])
            date_col = ['采购日期']
            default_date = pd.Timestamp(1990, 1, 1)
            report002[date_col] = report002[date_col].fillna(default_date)
            report002['采购日期'] = pd.to_datetime(report002['采购日期'], errors='coerce')
            # 文本
            text_col = [ "核算项目号", "采购单号",  "采购单类型", "料件编号", "品名", "规格"
                ,  "备注", "项目编号", "采购人员", "采购供应商", "库存管理特征"
                , "行状态", "状态", "作业编号", "模组标识", "模组名称", "品牌", "零件类型"
                , "产能", "标准/非标", "客户简称", "大项目名称", "大项目号", "设备名称"
                , "产品线编码", "产品线名称", "工艺", "项目号整理", "留置原因说明", "采购部门", "料件分类", "来源单号"
                , "数据-来源", "数据分布", "项目财经"]
            report002[text_col] = report002[text_col].fillna('')
            report002["项目编号"] = report002["项目编号"].fillna('').astype(str)
            report002["项目编号"] =report002["项目编号"].replace(' ', '', regex=True)
            report002["料件编号"] =report002["料件编号"].replace('=\"','', regex=True)
            report002["料件编号"] = report002["料件编号"].replace('\"', '', regex=True)
            # 数值
            int_col = ["采购数量", "未税单价", "采购金额-未税", "单价"
                , "含税金额", "税率", "已收货量", "未交量"
                ,"验退量", "仓退换货量", "仓退量"]
            report002[int_col] = report002[int_col].fillna(0)
####################数据处理
            screm.insert(INSERT, '\n1.2.3：正在处理采购PO字段...', '\n')
            window.update()
            if len(report002[report002['项目编号'].str.contains('-')]) > 0:
                report002['项目号整理'] = ''
                report002['项目整'] = report002['项目编号'].str.split('-', expand=True)[0]
                report002['项目整1'] = report002['项目编号'].str.split('-', expand=True)[1]
                report002['项目整1'] = report002['项目整1'].fillna('空值')
                report002['项目号整理'] = report002['项目整']
                report002.loc[(report002['项目整1'].str.isdigit())|(report002['项目整1'].str.contains('SH')), '项目号整理'] = report002['项目号整理'] + '-' + report002['项目整1']
            if len(report002[report002['项目编号'].str.contains('-')]) == 0:
                report002['项目号整理'] = report002['项目编号']
            report002.loc[(report002['项目号整理'].str[0] == 'F') & (report002['项目号整理'].str[:3].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX')), '项目号整理'] = report002['项目号整理'].str[3:]
            report002.loc[(report002['项目号整理'].str[:2].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX', na=False)), '项目号整理'] =report002['项目号整理'].str[2:]
####################数据处理
            process=report002.copy()
            process['标准/非标'] = ''
            process['料件编号'] = process['料件编号'].fillna('\\')
            process['零件类型']= '非标件'
            process.loc[process['料件编号'].str[:1].str.isdigit(), '标准/非标'] = '标准件'
            process['零件类型'] = process['零件类型'].fillna('空值')
            process.loc[process['零件类型'].str.contains('标准'), '标准/非标'] = '标准件'
            process.loc[(process['零件类型'].str.contains('空值|标准') == False), '标准/非标'] = '非标件'
            process['零件类型'] = process['零件类型'].replace('空值', '', regex=True)
            # 作业编号更新
            process = process.reset_index(drop=True)
            process['作业编号'] = process['作业编号'].fillna('无模组号').astype(str)
            process['作业编号'] = process['作业编号'].astype(str)
            process.loc[process['作业编号'] == "", '作业编号'] = '无模组号'
            process.loc[process['作业编号'] == " ", '作业编号'] = '无模组号'
            process.loc[process['作业编号'] == 0, '作业编号'] = '无模组号'
            process['作业编号'] = process['作业编号'].replace('无模组号', '', regex=True)
            # 模组名称
            process['模组标识'] = process['核算项目号'].astype(str) + '-' + process['作业编号'].astype(str)
            process['未交量'] = pd.to_numeric(process['未交量'])
            process['采购数量'] = pd.to_numeric(process['采购数量'])
            process['单价'] = pd.to_numeric(process['单价'])
            process['百分比'] = process['税率'].astype(str).str.replace('%', '', regex=True).astype(float) / 100
            # 未税单价
            process['未税单价'] = process['单价'] / (1 + process['百分比'])
            process['未税单价'] = process['未税单价'].fillna(0)
            del process['百分比']

            #################采购PO剔除
            screm.insert(INSERT, '\n1.2.4：正在剔除采购PO金额...', '\n')
            window.update()

            process.loc[(process['行状态'].str.contains('短结'))&(process['未交量']<process['采购数量']), '采购数量'] =process['采购数量'] - process['未交量']
            # 采购金额-未税
            process['采购金额-未税'] = process['采购数量'] * process['未税单价']
            ###剔除
            process.loc[(process['状态'].str.contains('作废')), '采购金额-未税'] = 0
            process['备注'] = process['备注'].fillna('')
            process.loc[(process['行状态'].str.contains('短结|已拒绝')) & (process['采购数量'] == process['未交量'])&(process['采购供应商'].str.contains('001') == False) & (process['备注'].str.contains('cint350调拨拨入过账-结案采购单') == False), '采购金额-未税'] = 0

            process.loc[(process['采购供应商'].str.contains('海目星激光科技集团股份有限')) & (process['品名'].str[-4:].str.contains('激光器')), '采购金额-未税'] = 0
            process.loc[(process['采购单类型'].str.contains('期初采购单')), '采购金额-未税'] = 0
            process.loc[process['采购单类型'].str.contains("多角采购a\(LEBG-->LEBGJM销售\)|多角采购m\(LEBG-->LEBGJM销售\)内外价|多角采购p\(LEBGJS-->LEBGJM销售\)|多角采购n\(LEBG-->LEBGJS销售\)内外价|多角采购e\(LEBG-->LEBGJS销售\)|多角采购t\(LEBG-->LEBGHX销售\)人民币"), '采购金额-未税'] = 0
            process.loc[process['采购单类型'].str.contains("多角采购t\(LEBG-->LEBGHX销售\)人民币|关联交易采购订单|多角采购u\(LEBG-->LEBGHX销售\)外币|多角采购b\(LEBGJS-->LEBGHX销售\)人民币|多角采购p\(LEBGJS-->LEBGJM销售\)内外价|多角采购s\(LEBGJS-->LEBG销售\)|集采多角采购q\(LEBG-->LEBGJM销售\)"), '采购金额-未税'] = 0
            ####输出剔除项次
            process['剔除'] = ''
            process.loc[(process['行状态'].str.contains('短结')) & (process['未交量'] < process['采购数量']), '剔除'] ='是'
            process.loc[(process['状态'].str.contains('作废')), '剔除'] ='是'
            process.loc[(process['行状态'].str.contains('短结|已拒绝')) & (process['采购数量'] == process['未交量']) & (process['采购供应商'].str.contains('001') == False)&(process['备注'].str.contains('cint350调拨拨入过账-结案采购单') == False), '剔除'] ='是'
            process.loc[(process['采购供应商'].str.contains('海目星激光科技集团股份有限')) & (process['品名'].str[-4:].str.contains('激光器')), '剔除'] ='是'
            process.loc[(process['采购单类型'].str.contains('期初采购单')), '剔除'] ='是'
            process.loc[process['采购单类型'].str.contains("多角采购a\(LEBG-->LEBGJM销售\)|多角采购m\(LEBG-->LEBGJM销售\)内外价|多角采购p\(LEBGJS-->LEBGJM销售\)|多角采购n\(LEBG-->LEBGJS销售\)内外价|多角采购e\(LEBG-->LEBGJS销售\)|多角采购t\(LEBG-->LEBGHX销售\)人民币"), '剔除'] ='是'
            process.loc[process['采购单类型'].str.contains("多角采购t\(LEBG-->LEBGHX销售\)人民币|关联交易采购订单|多角采购u\(LEBG-->LEBGHX销售\)外币|多角采购b\(LEBGJS-->LEBGHX销售\)人民币|多角采购p\(LEBGJS-->LEBGJM销售\)内外价|多角采购s\(LEBGJS-->LEBG销售\)|集采多角采购q\(LEBG-->LEBGJM销售\)"), '剔除'] ='是'
            ####输出剔除项次
#################对照表数据
            screm.insert(INSERT, '\n1.2.5：正在拉取采购PO对照表数据...', '\n')
            window.update()
            for i in ["核算项目号",  "设备名称",'客户简称','项目财经','项目财经再分类']:
                if i in process.columns:
                    del process[i]

            process[["核算项目号",  "设备名称",'客户简称','项目财经','项目财经再分类']] = pd.merge(process, report_item, left_on='项目号整理', right_on='项目号整理', how='left')[["核算项目号",  "设备名称",'客户简称','项目财经','项目财经再分类']]
            ##########整顿格式

            process = process.rename(columns={'单价': '含税单价'})
            process = process.reindex(columns=
                                      ["核算项目号", "采购单号", "采购日期", "采购单类型", "料件编号", "品名", "规格"
                                          , "采购数量", "未税单价", "采购金额-未税","含税单价"
                                          , "含税金额", "税率", "已收货量", "未交量", "备注", "项目编号", "采购人员", "采购供应商"
                                          , "库存管理特征", "行状态", "状态", "作业编号","模组标识", "模组名称", "品牌"
                                          , "零件类型", "产能","标准/非标", "客户简称", "大项目名称", "大项目号", "设备名称",'产品线编码'
                                          , "产品线名称",'工艺',"项目财经","项目财经再分类"
                                          ,"验退量", "仓退换货量", "仓退量", "项目号整理","留置原因说明"
                                          , "采购部门", "料件分类", "来源单号",'数据-来源','剔除'])
            process=process.fillna("")
            process_delete = process[process['剔除'].str.contains('是')].copy().reset_index(drop=True)
            del process['剔除']
        else:
            #print('采购明细当前无文件，请核查')
            screm.insert(INSERT, '\n采购明细当前无文件，请核查', '\n')
            process = pd.DataFrame(
                columns=["核算项目号", "采购单号", "采购日期", "采购单类型", "料件编号", "品名", "规格"
                                          , "采购数量", "未税单价", "采购金额-未税","含税单价"
                                          , "含税金额", "税率", "已收货量", "未交量", "备注", "项目编号", "采购人员", "采购供应商"
                                          , "库存管理特征", "行状态", "状态", "作业编号","模组标识", "模组名称", "品牌"
                                          , "零件类型", "产能","标准/非标", "客户简称", "大项目名称", "大项目号", "设备名称",'产品线编码'
                                          , "产品线名称",'工艺',"项目财经","项目财经再分类"
                                          ,"验退量", "仓退换货量", "仓退量", "项目号整理","留置原因说明"
                                          , "采购部门","料件分类", "来源单号"])
            window.update()

        ###输出采购PO异常
        process['核算项目号'] = process['核算项目号'].fillna('').astype(str)
        process_error = process[(process['核算项目号'].str.len()<2)].drop_duplicates(subset=['项目编号']).reset_index(drop=True)

        screm.insert(INSERT, '\n*****新增采购PO-存在未参与核算项目号' + str(len(process_error)) + '个', '\n')
        window.update()

        time_po = time.time()
        screm.insert(INSERT, '\n处理采购明细数据:%d秒' % (time_po - time_base), '\n')
        window.update()

###############################################################################################################################################################料
        screm.insert(INSERT, '\n1.3：料...', '\n')
        window.update()
        need_col=["项目编号","扣账日期","发退料单号","料号","品名","规格","仓库说明","单位","数量","未税单价",'领料类型','大项目号','大项目名称','产品线名称','产品线编码','单别']
        #杂发
        #print("正在处理杂发")
##############################杂收发
        screm.insert(INSERT, '\n1.3.1：正在处理杂收发...', '\n')
        window.update()
        path_send= r'数据源\料-数据源\杂收发'
        index = 0
        report_send = []
        file_send= os.listdir(path_send)
        for i in file_send:
            if '~$' in i:
                file_send.remove(i)
        if os.listdir(path_send):
            for name in file_send:
                if 'csv' in name and '~$' not in name:
                    #print('读取杂发第'+str(index+1)+'份'+name)
                    screm.insert(INSERT, '\n读取杂收发第'+str(index+1)+'份'+name, '\n')
                    window.update()
                    with open(os.path.join(path_send,name), encoding='utf-8', errors='ignore') as f:
                        df4= pd.read_csv(f, sep=',', error_bad_lines=False, low_memory=False,thousands=',')
                    df4['项目编号']=df4['项目编号'].fillna("空值")
                    df4=df4[df4['项目编号']!='空值']
                    df4['领料类型']='杂收发'
                    df4['数量']=-1*df4['本期异动数量']
                    df4['未税单价']=df4['本期异动单价']
                    if '单位名称' in df4.columns:
                        df4 = df4.rename(columns={'单位名称': '单位'})
                    if '成本单位名称' in df4.columns:
                        df4 = df4.rename(columns={'成本单位名称': '单位'})
                    df4 = df4.rename(columns={'单位名称': '单位','库位名称':'仓库说明','产品线':'产品线编码'})
                    for col   in df4.columns:
                        if col not in need_col:
                             del df4[col]
                    report_send.append(df4)
                    index += 1  # 为了查看合并到第几个表格了
            reportsend = pd.concat(report_send).reset_index(drop=True)
            default_date = pd.Timestamp(1990, 1, 1)
            reportsend["扣账日期"] = reportsend["扣账日期"].fillna(default_date)
            reportsend["扣账日期"] = pd.to_datetime(reportsend["扣账日期"], errors='coerce')
            reportsend['数据-来源'] = '新增数据'
        else:
            #print('杂发无文件')
            screm.insert(INSERT, '\n杂收发无文件', '\n')
            window.update()
            reportsend= pd.DataFrame(
                columns=["项目编号", "扣账日期", "发退料单号", "料号", "品名", "规格", "仓库说明", "单位", "数量", "未税单价",
                         "单别", "项目号整理", '领料类型','成品料号','大项目号','大项目名称','产品线名称','产品线编码','数据-来源'])

            reportsend[["数量", '未税单价']] = reportsend[["数量", '未税单价']].fillna(0)
            reportsend= reportsend.fillna('')
        #print("正在处理在制")
##############################在制
        screm.insert(INSERT, '\n1.3.2：正在处理在制...', '\n')
        window.update()
        ###########在制
        need_col = ["项目编号", "扣账日期", "发退料单号", "料号", "品名", "规格", "仓库说明", "单位",  "数量", "未税单价", "工单号码", '工单单号',
                     "单别", "作业编号","项目号整理", '领料类型', '品牌', '成品料号','大项目号','大项目名称','产品线名称','产品线编码','工单单据类别','成本单位名称','模组名称','零件类型']
        path_in= r'数据源\料-数据源\在制'
        index = 0
        report_in = []
        file_in= os.listdir(path_in)
        for i in file_in:
            if '~$' in i:
                file_in.remove(i)
        if os.listdir(path_in):
            for name in file_in:
                if 'csv' in name and '~$' not in name:
                    #print('读取在制第'+str(index+1)+'份'+name)
                    screm.insert(INSERT, '\n读取在制第'+str(index+1)+'份'+name, '\n')
                    window.update()
                    with open(os.path.join(path_in,name), encoding='gb18030', errors='ignore') as f:
                        # 再解决部分报错行如 ParserError：Error tokenizing data.C error:Expected 2 fields in line 407,saw 3.
                        df6= pd.read_csv(f, sep=',', error_bad_lines=False, low_memory=False,thousands=',')
                        if '项目号' in list(df6.columns):
                            df6 = df6.rename(columns={'项目号': '项目编号','工单单号': '工单号码'})
                        if '大项目号名称' in list(df6.columns):
                            df6 = df6.rename(columns={'大项目号名称': '大项目名称'})
                        if '产品线' in list(df6.columns):
                            df6 = df6.rename(columns={'产品线': '产品线编码'})
                        if '工单单号' in list(df6.columns):
                            df6 = df6.rename(columns={'工单单号': '工单号码'})
                        if '库位名称' in list(df6.columns):
                            df6 = df6.rename(columns={'库位名称': '仓库说明'})
                        if '母件料号' in list(df6.columns):
                            df6 = df6.rename(columns={'母件料号': '成品料号'})
                    df6['项目编号']=df6['项目编号'].fillna("空值")
                    df6[['本期异动单价']]=df6[['本期异动单价']].fillna(0)
                    df6=df6[df6['项目编号']!='空值']
                    df6['领料类型']='在制'
                    df6['数量']=-1*df6['本期异动数量']
                    df6['未税单价']=df6['本期异动单价']
                    if '单位' in df6.columns:
                        del df6['单位']
                    df6 = df6.rename(columns={'成本单位名称': '单位'})
                    for col   in df6.columns:
                        if col not in need_col:
                             del df6[col]
                    report_in.append(df6)
                    index += 1  # 为了查看合并到第几个表格了
            reportin = pd.concat(report_in,ignore_index=True).reset_index(drop=True)
            default_date = pd.Timestamp(1990, 1, 1)
            reportin['扣账日期'] = pd.to_datetime(reportin['扣账日期'], errors='coerce')
            reportin['扣账日期'] = reportin['扣账日期'].fillna(default_date)
            reportin['数据-来源'] = '新增数据'
        else:
            #print('在制无文件')
            screm.insert(INSERT, '\n在制无文件', '\n')
            window.update()
            reportin=pd.DataFrame(columns=["项目编号","扣账日期",'零件类型','模组名称',"发退料单号","料号","品名","规格","仓库说明","单位","数量","未税单价","工单号码","单别","作业编号","项目号整理",'领料类型','品牌','成品料号','大项目号','大项目名称','产品线名称','产品线编码','新增数据'])
            default_date = pd.Timestamp(1990, 1, 1)
            reportin["扣账日期"] = reportin["扣账日期"].fillna(default_date)
            reportin[["数量", '未税单价']] = reportin[["数量", '未税单价']].fillna(0)
            reportin=reportin.fillna('')
        screm.insert(INSERT, '\n1.3.3：正在合并杂收发和在制...', '\n')
        window.update()
        report_material1 = pd.concat([reportsend,  reportin], ignore_index=True).reset_index(drop=True)
#############################
        screm.insert(INSERT, '\n1.3.4：正在处理料的字段数据...', '\n')
        window.update()
        #######3料项目号整理
        if len(report_material1[report_material1['项目编号'].str.contains('-')]) > 0:
            report_material1['项目号整理'] = ''
            report_material1['项目整'] = report_material1['项目编号'].str.split('-', expand=True)[0]
            report_material1['项目整1'] = report_material1['项目编号'].str.split('-', expand=True)[1]
            report_material1['项目整1'] = report_material1['项目整1'].fillna('空值')
            report_material1['项目号整理'] = report_material1['项目整']
            report_material1.loc[(report_material1['项目整1'].str.isdigit())|(report_material1['项目整1'].str.contains('SH')), '项目号整理'] = report_material1['项目号整理'] + '-' + report_material1['项目整1']
        else:
            report_material1['项目号整理'] = report_material1['项目编号']
        report_material1.loc[(report_material1['项目号整理'].str[0] == 'F') & (report_material1['项目号整理'].str[:3].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX')), '项目号整理'] = \
        report_material1['项目号整理'].str[3:]
        report_material1.loc[(report_material1['项目号整理'].str[:2].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX', na=False)), '项目号整理'] = \
        report_material1['项目号整理'].str[2:]
        #print("3料—项目号整理完成")
        screm.insert(INSERT, '\n3料—项目号整理完成', '\n')
        window.update()
        # 3料核算项目号
        report_material1['核算项目号'] = pd.merge(report_material1, report_item, left_on='项目号整理', right_on='项目号整理', how='left')['核算项目号']
        #4料合一
        report_material =report_material1.copy().reset_index(drop=True)
        # 模组标志
        report_material['核算项目号']=report_material['核算项目号'].fillna('')
        if '作业编号' not in report_material.columns:
            report_material['作业编号'] =''
        report_material['作业编号']=report_material['作业编号'].fillna('')
        report_material['模组标识'] = report_material['核算项目号'] + '-' + report_material['作业编号']
        # 零件类型
        #report_material['零件类型'] = pd.merge(report_material, report_20, left_on='料号', right_on='元件料号', how='left')['零件类型']
        # 标准\非标
        if '零件类型' not in report_material.columns:
            report_material['零件类型']='空值'
        report_material['标件/非标件'] = '非标件'
        report_material['料号']=report_material['料号'].fillna('\\')
        #report_material.loc[report_material['料号'].str[:1].str.isalpha(), '标件/非标件'] = '非标件'
        report_material.loc[report_material['料号'].str[:1].str.isdigit() , '标件/非标件'] = '标准件'
        report_material['零件类型'] = report_material['零件类型'].fillna('空值')
        report_material.loc[report_material['零件类型'].str.contains('标准'), '标件/非标件'] = '标准件'
        report_material.loc[(report_material['零件类型'].str.contains('空值|标准') == False), '标件/非标件'] = '非标件'
        report_material['零件类型'] = report_material['零件类型'].replace('空值', '', regex=True)
        default_date = pd.Timestamp(1990, 1, 1)
        report_material['扣账日期'] = report_material['扣账日期'].fillna(default_date)
        # 是否2022年
        report_material['扣账日期'] = pd.to_datetime(report_material['扣账日期'], errors='coerce')
        # 数值填充
        report_material[['数量', '未税单价']] = report_material[['数量', '未税单价']].fillna(0)
        report_material.loc[report_material['未税单价']!=19950222,'未税金额']=report_material['未税单价'] * report_material['数量']
        report_material['未税单价'] = report_material['未税单价'].replace(19950222, '', regex=True)
        #是否有工单
        report_material['工单号码']=report_material['工单号码'].fillna('')
        report_material['是否有工单']='否'
        report_material=report_material.reset_index(drop=True)
        report_material.loc[report_material['工单号码'].str.contains('601-|603-|608-'), '是否有工单'] = '是'
        report_material  = report_material.rename(columns={'品牌': '品牌/供应商'})
################################
        screm.insert(INSERT, '\n1.3.5：料正在拉取对照表数据...', '\n')
        window.update()
        report_mat = pd.merge(report_material, report_item, on='核算项目号', how='left')
        report_material[['客户简称', '设备名称', '项目财经', '项目财经再分类']] = report_mat[['客户简称', '设备名称', '项目财经', '项目财经再分类']]
        report_material['项目财经'] = report_material['项目财经'].fillna('找不到')
        report_material['项目财经再分类'] = report_material['项目财经再分类'].fillna('找不到')

        report_material=report_material[~report_material.index.duplicated()]##############有一个重复索引找不到
        report_material = report_material.reindex(columns=["核算项目号", "项目编号", "扣账日期", "发退料单号", "料号", "品名"
                                                      , "规格", '仓库说明', '单位', "数量", "未税单价", "未税金额"
                                                      , "工单号码",'单别',"作业编号","模组标识", "模组名称"
                                                      , "品牌/供应商", "零件类型", "产能", "标件/非标件", "客户简称",'产品线编码', "产品线名称","大项目号",'大项目名称'
                                                      , "设备名称",'工艺',"项目号整理",'成品料号','是否有工单','项目财经',"领料类型",'项目财经再分类','数据-来源'])

        report_material['扣账日期']= pd.to_datetime(report_material['扣账日期'], errors='coerce')
        report_material['扣账日期'] = report_material['扣账日期'].dt.strftime('%Y-%m-%d').astype(str)
        report_material['扣账日期'] = ['' if i == '1990-01-01' else i for i in report_material['扣账日期']]
        report_material=report_material.fillna('')
####################剔除金额
        screm.insert(INSERT, '\n1.3.6：料正在剔除金额...', '\n')
        window.update()
        path_mater_602 = r'数据源\602立旧项目'
        file_mater_602 = os.listdir(path_mater_602)
        for i in file_mater_602:
            if '~$' in i:
                file_mater_602.remove(i)
        if os.listdir(path_mater_602 ):
            for name in file_mater_602:
                if 'xlsx' in name and '~$' not in name:
                    mater_602 = pd.read_excel(path_mater_602 + '\\' + name)[['工单']]
        mater_602=mater_602.fillna('')
        report_material.loc[(report_material['品名'].str.contains('海目星'))&(report_material['品名'].str.contains('软件')), '未税金额'] = 0
        report_material.loc[(report_material['料号'].str[:5].str.contains('F5-'))&(report_material['料号'].str.len() <= 15)&(report_material['工单号码'].str.contains('601-')==False),'未税金额']=0
        report_material.loc[ (report_material['料号'].str[:6].str.contains('B01-'))&(report_material['料号'].str.len() <= 11), '未税金额'] = 0
        report_material.loc[(report_material['料号'].str[:5].str.contains('F8-')), '未税金额'] = 0
        ##保留LEW不带-
        '''
        report_material.loc[(~report_material['工单号码'].isin(mater_602['工单']))&(report_material['料号'].str[:6].str.contains('F8-|B01-|F5-')==False)&(report_material['工单号码'].str.contains('602-'))&(report_material['项目号整理'].str.contains('LEW')==False)|(report_material['项目号整理'].str.contains('-')), '未税金额'] = 0
        report_material.loc[(report_material['料号'].str[:6].str.contains('F8-|B01-|F5-')) & (report_material['工单号码'].str.contains('602-')) & (report_material['项目号整理'].str.contains('LEW') == False) | (report_material['项目号整理'].str.contains('-')), '未税金额'] = 0
        '''
        report_material.loc[report_material['工单号码'].str.contains('602-'), '未税金额'] = 0
        report_material.loc[(report_material['工单号码'].str.contains('602-')) & (report_material['项目编号'].str.contains('LEW'))&(report_material['项目编号'].str.contains('-') == False), '未税金额'] = report_material['未税单价'] * report_material['数量']

        report_material.loc[(report_material['工单号码'].isin(mater_602['工单'])) & (report_material['料号'].str[:6].str.contains('F8-|B01-|F5-') == False), '未税金额'] = report_material['未税单价']*report_material['数量']
        report_material.loc[report_material['工单号码'].str.contains('604-|606-'),'未税金额']=0
        ###########6.27王小琴指定剔除新增料
        report_material.loc[(report_material['料号'].str.contains('JGZY23A001|JGZY23A002|JGZY23B001|JGZY23B002|JGZY23B003')), '未税金额'] =0
        ###########6.27王小琴指定剔除新增料

        ####拉出剔除项次
        report_material['剔除']=''
        report_material.loc[(report_material['品名'].str.contains('海目星')) & (report_material['品名'].str.contains('软件')), '剔除']='是'
        report_material.loc[(report_material['料号'].str[:5].str.contains('F5-')) & (report_material['料号'].str.len() <= 15) & (report_material['工单号码'].str.contains('601-') == False), '剔除']='是'
        report_material.loc[(report_material['料号'].str[:6].str.contains('B01-')) & (report_material['料号'].str.len() <= 11), '剔除']='是'
        report_material.loc[(report_material['料号'].str[:5].str.contains('F8-')), '剔除']='是'
        report_material.loc[report_material['工单号码'].str.contains('602-'), '剔除']='是'
        report_material.loc[(report_material['工单号码'].str.contains('602-')) & (report_material['项目编号'].str.contains('LEW')) & (report_material['项目编号'].str.contains('-') == False), '剔除']=''
        report_material.loc[(report_material['工单号码'].isin(mater_602['工单'])) & (report_material['料号'].str[:6].str.contains('F8-|B01-|F5-') == False), '剔除']=''
        report_material.loc[report_material['工单号码'].str.contains('604-|606-'), '剔除']='是'
        ###########6.27王小琴指定剔除新增料
        report_material.loc[(report_material['料号'].str.contains('JGZY23A001|JGZY23A002|JGZY23B001|JGZY23B002|JGZY23B003')), '剔除']='是'
        report_material_delete=report_material[report_material['剔除'].str.contains('是')].copy().reset_index(drop=True)
        ####拉出剔除项次

        ###输出料异常号
        report_material['核算项目号'] = report_material['核算项目号'].fillna('').astype(str)
        report_material_err = report_material[(report_material['核算项目号'].str.len() < 2)].drop_duplicates(subset=['项目编号']).reset_index(drop=True)

        screm.insert(INSERT, '\n******新增料存在未参与核算项目号：' + str(len(report_material_err)) + '个', '\n')
        window.update()

        time_mater=time.time()
        #print('处理四块料:%d秒' % (time_mat3 - time_mat))
        screm.insert(INSERT, '\n处理四块料:%d秒' % (time_mater - time_po), '\n')
        window.update()

#########################################################################################################################################################工
        screm.insert(INSERT, '\n1.4：工...', '\n')
        screm.insert(INSERT, '\n1.4.1：正在处理工...', '\n')
        window.update()
        path_work = r'数据源\工-数据源'
        need_work = ["项目号","姓名","工号","人员归属","成本归属","工种大类",'部门',"部门说明","科室","岗位","工作地点","提报人","项目类别",'报工单号',"工单号","完成日期","工时合计(小时)","备注","项目阶段","报工类别","工种",'交付阶段','大项目号','大项目名称']
        report_work=[]
        index=0
        file_work= os.listdir(path_work)
        for i in file_work:
            if '~$' in i:
                file_work.remove(i)
        if os.listdir(path_work):
            for name in file_work:
                if 'csv' in name and '~$' not in name:
                    #df = pd.read_excel(os.path.join(path_cost, name), header=2,thousands=',')
                    with open(os.path.join(path_work , name), encoding='gb18030', errors='ignore') as f:
                        df8 = pd.read_csv(f, sep=',',error_bad_lines=False, low_memory=False,thousands=',')
                    for col in df8.columns:
                        if col not in need_work:
                            del df8[col]
                    report_work.append(df8)
                    index += 1  # 为了查看合并到第几个表格了
            reportwork = pd.concat(report_work).reset_index(drop=True)
            reportwork['项目号'] = reportwork['项目号'].fillna('').astype(str)
            reportwork['项目号'] = reportwork['项目号'].replace(' ', '', regex=True)
            #####重命名字段
            reportwork = reportwork.rename(
                columns={'部门': '部门编码'})
            reportwork = reportwork.rename(columns={'工种大类': '工时种类','工单号':'报工工单号','部门说明':'部门','项目类别':'报工来源','工时合计(小时)':'工时','项目阶段':'阶段','报工类别':'报工组别','工种':'工种再分类'})
            reportwork['完成日期']=reportwork['完成日期'].fillna(default_date)
            reportwork['完成日期']= pd.to_datetime(reportwork['完成日期'],errors='coerce')
            reportwork['事业部']=reportwork['部门'].str.split('-', expand=True)[0]
            reportwork['工单标识']=''
            reportwork['工种'] = ''
            reportwork[['工时种类','交付阶段','部门','部门编码','报工单号']]=reportwork[['工时种类','交付阶段','部门','部门编码','报工单号']].fillna('')
            reportwork.loc[reportwork['工时种类'].str.contains('交付'),'工种'] = '交付工'
            reportwork.loc[reportwork['工时种类'].str.contains('生产'), '工种'] = '生产工'
            reportwork.loc[reportwork['工时种类'].str.contains('设计|项目|其他'), '工种'] = '设计工'
            reportwork.loc[(reportwork['工时种类'].str.contains('其他')) & (reportwork['部门'].str.contains('供应链')), '工种'] = '交付工'
            reportwork.loc[(reportwork['工时种类'].str.contains('其他'))&(reportwork['部门'].str.contains('供应链'))&(reportwork['交付阶段'].str.contains('其他')), '工种'] = '生产工'
            reportwork['科目类别']=''
            reportwork.loc[reportwork['工种'].str.contains('交付|生产'), '科目类别'] = '直接人工'
            reportwork.loc[reportwork['工种'].str.contains('设计'), '科目类别'] = '制造费用'
            reportwork['工时'] = reportwork['工时'].fillna(0)
            reportwork['工价']=0

##################################根据工价规则指定工价
            screm.insert(INSERT, '\n1.4.2：正在根据工价规则指定工价...', '\n')
            window.update()
            path_work_rule = r'数据源\工-数据源\工价规则'
            index = 0
            file_work_rule = os.listdir(path_work_rule)
            for i in file_work_rule:
                if '~$' in i:
                    file_work_rule.remove(i)
            if os.listdir(path_work_rule):
                for name in file_work_rule:
                    if 'xlsx' in name and '~$' not in name:
                        work_item = pd.read_excel(path_work_rule + '\\' + name, sheet_name='部门对应表')[['部门', '部门说明', '对应部门', '对应中心']]
                        work_rule = pd.read_excel(path_work_rule + '\\' + name, sheet_name='工价表')[['对应部门', '月工价', '期别']]
            ##########第一层按部门编码拉取
            reportwork['对应部门'] = ''
            reportwork.loc[:, '对应部门'] = vlookup(reportwork.loc[:, '部门编码'],
                                                work_item.loc[:, '部门'],
                                                work_item.loc[:, '对应部门'])
            ##########第二层按部门编码、部门做包含确定
            work_item['对应部门'] = work_item['对应部门'].fillna('')
            work_list = list(work_item['对应部门'].drop_duplicates())
            for i in work_list:
                reportwork.loc[(reportwork['对应部门'] == '') & (reportwork['部门编码'].str.contains(i)), '对应部门'] = i
                reportwork.loc[(reportwork['对应部门'] == '') & (reportwork['部门'].str.contains(i)), '对应部门'] = i
            ##########第三层按报工单号确定
            reportwork.loc[(reportwork['对应部门'] == '') & (reportwork['报工单号'].str.contains('LE')), '对应部门'] = '其他-设计部门'
            reportwork.loc[(reportwork['对应部门'] == '') & (reportwork['报工单号'].str.contains('LJ')), '对应部门'] = '华南供应链'
            reportwork.loc[(reportwork['对应部门'] == '') & (reportwork['报工单号'].str.contains('LS')), '对应部门'] = '华东供应链'
            ##期别
            reportwork['期别'] = reportwork['完成日期'].dt.strftime("%Y%m").astype(str)
            ###唯一字段
            reportwork['部门期别'] = reportwork['对应部门'] + reportwork['期别']
            work_rule['期别'] = work_rule['期别'].fillna('').astype(str)
            work_rule['对应部门'] = work_rule['对应部门'].fillna('')
            work_rule['月工价'] = work_rule['月工价'].fillna(0)
            work_rule['期别'] = work_rule['期别']
            work_rule['部门期别'] = work_rule['对应部门'] + work_rule['期别']
            reportwork.loc[:, '工价'] = vlookup(reportwork.loc[:, '部门期别'],work_rule.loc[:, '部门期别'],work_rule.loc[:, '月工价'])
            reportwork.loc[reportwork['工价'] == '', '工价'] = 0
            reportwork['工价'] = reportwork['工价'].astype(float)
            reportwork['工价']=reportwork['工价'].fillna(0)
            reportwork.loc[(reportwork['工种'].str.contains('设计'))&(reportwork['工价']==0), '工价'] = 75
            reportwork.loc[(reportwork['工种'].str.contains('生产|交付'))&(reportwork['工价']==0), '工价'] = 50
            reportwork['工时成本']=0
            reportwork['工价']=reportwork['工价'].astype(float)
            reportwork['工时成本'] = reportwork['工价']*reportwork['工时']
###############################
            screm.insert(INSERT, '\n1.4.3：正在整理工的字段信息...', '\n')
            window.update()
            del reportwork['部门编码']
            del reportwork['报工单号']
            del reportwork['对应部门']
            del reportwork['期别']
            del reportwork['部门期别']
            reportwork['项目号整理'] = ''
            if len(reportwork[reportwork['项目号'].str.contains('-')]) > 0:
                reportwork['项目整'] = reportwork['项目号'].str.split('-', expand=True)[0]
                reportwork['项目整1'] = reportwork['项目号'].str.split('-', expand=True)[1]
                reportwork['项目整1'] = reportwork['项目整1'].fillna('空值')
                reportwork['项目号整理'] = reportwork['项目整']
                reportwork.loc[(reportwork['项目整1'].str.isdigit())|(reportwork['项目整1'].str.contains('SH')), '项目号整理'] = reportwork['项目号整理'] + '-' +  reportwork['项目整1']
                del reportwork['项目整']
                del reportwork['项目整1']
            else:
                reportwork['项目号整理'] = reportwork['项目号']
            reportwork.loc[(reportwork['项目号整理'].str[0] == 'F') & (reportwork['项目号整理'].str[:3].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX')), '项目号整理'] = reportwork['项目号整理'].str[3:]
            reportwork.loc[(reportwork['项目号整理'].str[:2].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX', na=False)), '项目号整理'] = reportwork['项目号整理'].str[2:]
            #reportwork['项目财经']=''
            report_wo = pd.merge(reportwork, report_item, on='项目号整理', how='left')
            reportwork[['核算项目号', "客户简称", "产品线名称", '产品线编码',"设备名称",'自制/外包', '项目财经','项目财经再分类']] = report_wo[['核算项目号', '客户简称','产品线名称','产品线编码','设备名称','自制/外包','项目财经','项目财经再分类']]
            reportwork=reportwork.fillna('')
            reportwork = reportwork.reindex(columns=["核算项目号","项目号","姓名","工号","人员归属","成本归属","事业部","部门","科室","岗位","工作地点","提报人","报工来源","报工工单号","工单标识","完成日期",'工种',"工时","工价","工时成本","备注","阶段","报工组别","工种再分类","客户简称","产品线编码","产品线名称","大项目号","大项目名称","设备名称","产能","自制/外包","项目号整理","项目财经",'项目财经再分类'])
            reportwork.loc[(reportwork['报工工单号']=='0')|(reportwork['报工工单号']==0),'报工工单号']=''
        else:
            #print('工无文件')
            screm.insert(INSERT, '\n工无文件', '\n')
            window.update()
            reportwork = pd.DataFrame(columns=["核算项目号","项目号","姓名","工号","人员归属","成本归属","事业部","部门","科室","岗位","工作地点","提报人","报工来源","报工工单号","工单标识","完成日期",'工种',"工时","工价","工时成本","备注","阶段","报工组别","工种再分类","客户简称","产品线编码","产品线名称","大项目号","大项目名称","设备名称","产能","自制/外包",'工艺',"项目号整理","项目财经",'项目财经再分类'])
            reportwork=reportwork.fillna('')
        #检查工项目号
        '''
        if  len(reportwork[(reportwork['项目号'].str.contains(' |sz|Sz|Sz|j|s'))&(reportwork['项目号'].str.len()>8)])>1:
            #print("工-存在异常项目号如下")
            #print(reportwork[(reportwork['项目号'].str.contains(' |sz|Sz|Sz|j|s'))&(reportwork['项目号'].str.len()>8)].drop_duplicates())
            screm.insert(INSERT, '\n工-存在编辑异常项目号如下：', '\n')
            screm.insert(INSERT, '\n'+reportwork[(reportwork['项目号'].str.contains(' |sz|Sz|Sz|j|s'))&(reportwork['项目号'].str.len()>8)].drop_duplicates(), '\n')
            window.update()
        '''
        ###输出工项目号
        reportwork['核算项目号']=reportwork['核算项目号'].fillna('').astype(str)
        reportwork_err=reportwork[(reportwork['核算项目号'].str.len()<2)].drop_duplicates(subset=['项目号']).reset_index(drop=True)

        screm.insert(INSERT, '\n***********新增工-存在未参与核算项目号：'+str(len(reportwork_err))+'个', '\n')
        window.update()

        time_work = time.time()
        #print('处理工:%d秒' % (time_work - time_mat3))
        #print("正在处理费")
        screm.insert(INSERT, '\n处理工:%d秒' % (time_work - time_mater), '\n')

##########################################################################################费
        screm.insert(INSERT, '\n1.5：费...', '\n')
        screm.insert(INSERT, '\n1.5.1：正在处理费...', '\n')
        window.update()
        need_cost = ["月份","凭证编号","摘要","科目编码","科目名称","部门名称","项目号","金额","中心","汇总科目编码","费用类型","公司名称","事业部","事业部重分类","部门重分类","科室重分类","经管费用归属","经管科目一级","经管科目二级","金额(万元)","项目类型","无项目报工率","备注",'是否核算']
        #财务费
        screm.insert(INSERT, '\n正在读取财务：费用报销表第' + str(index + 1) + '份' + name, '\n')
        window.update()
        path_cost1 = r'数据源\费-数据源\财务'
        index = 0
        report_cost1 = []
        file_cost1 = os.listdir(path_cost1)
        for i in file_cost1:
            if '~$' in i:
                file_cost1.remove(i)
        if os.listdir(path_cost1):
            for name in file_cost1:
                if 'csv' in name and '~$' not in name:
                    # print('读取费用报销表第' + str(index + 1) + '份' + name)
                    screm.insert(INSERT, '\n正在读取财务-读取费用报销表第' + str(index + 1) + '份' + name, '\n')
                    window.update()
                    with open(os.path.join(path_cost1, name), encoding='gb18030', errors='ignore') as f:
                        df10 = pd.read_csv(f, sep=',', error_bad_lines=False, low_memory=False,thousands=',')
                        if '一级部门' in list(df10.columns):
                            df10 = df10.rename(columns={'一级部门': '事业部'})
                        if '二级部门' in list(df10.columns):
                            df10 = df10.rename(columns={'二级部门': '事业部重分类'})
                        if '三级部门' in list(df10.columns):
                            df10 = df10.rename(columns={'三级部门': '部门重分类'})
                        if '四级部门' in list(df10.columns):
                            df10 = df10.rename(columns={'四级部门': '科室重分类'})
                        if '部门属性' in list(df10.columns):
                            df10 = df10.rename(columns={'部门属性': '经管费用归属'})
                        for col in df10.columns:
                            if col not in need_cost:
                                del df10[col]
                    report_cost1.append(df10)
                    index += 1  # 为了查看合并到第几个表格了
            reportcost1 = pd.concat(report_cost1).reset_index(drop=True)
            screm.insert(INSERT, '\n1.5.2：正在整理费的字段数据...', '\n')
            window.update()
            reportcost1['是否核算']=reportcost1['是否核算'].fillna('')
            reportcost1[['金额','金额(万元)']]=reportcost1[['金额','金额(万元)']].fillna(0)
            reportcost1.loc[reportcost1['是否核算'].str.contains('否'),'金额(万元)']=0
            #reportcost1['金额(万元)']=reportcost1['金额']/10000
            reportcost1['金额万元'] = reportcost1['金额(万元)']*10000
            reportcost1['项目号整理'] = ''
            if len(reportcost1[reportcost1['项目号'].astype(str).str.contains('-')]) > 0:
                reportcost1['项目整'] = reportcost1['项目号'].str.split('-', expand=True)[0]
                reportcost1['项目整1'] = reportcost1['项目号'].str.split('-', expand=True)[1]
                reportcost1['项目整1'] = reportcost1['项目整1'].fillna('空值')
                reportcost1['项目号整理'] = reportcost1['项目整']
                reportcost1.loc[(reportcost1['项目整1'].str.isdigit())|(reportcost1['项目整1'].str.contains('SH')), '项目号整理'] = reportcost1['项目号整理'] + '-' + reportcost1['项目整1']
            else:
                reportcost1['项目号整理'] = reportcost1['项目号']

            reportcost1.loc[(reportcost1['项目号整理'].str[0] == 'F') & (
                reportcost1['项目号整理'].str[:3].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX')), '项目号整理'] = \
            reportcost1['项目号整理'].str[3:]
            reportcost1.loc[
                (reportcost1['项目号整理'].str[:2].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX', na=False)), '项目号整理'] = \
            reportcost1['项目号整理'].str[2:]
            report_co1 = pd.merge(reportcost1, report_item, on='项目号整理', how='left')
            screm.insert(INSERT, '\n1.5.3：费正在拉取对照表信息...', '\n')
            window.update()
            reportcost1[['核算项目号',"客户简称","产品线编码","产品线名称","大项目号","大项目名称","设备名称"]]= report_co1[['核算项目号','客户简称', '产品线编码', '产品线名称', '大项目号', '大项目名称','设备名称']]
            ##生成年份+月份
            reportcost1['月份']=reportcost1['月份'].fillna(0).astype(int)
            reportcost1['年份']=str(int(time.strftime("%Y", time.localtime(time.time()))))+'年'
            reportcost1.loc[reportcost1['月份']==12,'年度']=str(int(time.strftime("%Y", time.localtime(time.time())))-1)+'年'
            reportcost1 = reportcost1.reindex(columns=['核算项目号','年份',"月份","凭证编号","摘要","科目编码","科目名称","部门名称","项目号","金额","中心","汇总科目编码","费用类型","公司名称","事业部","事业部重分类","部门重分类","科室重分类","经管费用归属","经管科目一级","经管科目二级","金额(万元)","项目类型","无项目报工率","备注",'项目号整理',"有无项目号","费用大类","费用小类","是否核算","客户简称","产品线编码","产品线名称","大项目号","大项目名称","设备名称","产能",'工艺','金额万元'])
            reportcost1[['项目财经','项目财经再分类']]=pd.merge(reportcost1, report_item, left_on='项目号整理', right_on='项目号整理', how='left')[['项目财经','项目财经再分类']]
        else:
            screm.insert(INSERT, '\n财务-读取费用报销表无文件...', '\n')
            window.update()
            reportcost1 = pd.DataFrame(columns=['核算项目号','年份',"月份","凭证编号","摘要","科目编码","科目名称","部门名称","项目号","金额","中心","汇总科目编码","费用类型","公司名称","事业部","事业部重分类","部门重分类","科室重分类","经管费用归属","经管科目一级","经管科目二级","金额(万元)","项目类型","无项目报工率","备注",'项目号整理',"有无项目号","费用大类","费用小类","是否核算","客户简称","产品线编码","产品线名称","大项目号","大项目名称","设备名称","产能",'工艺',"项目财经","项目财经再分类"])
            reportcost1=reportcost1.fillna('')
        reportcost=reportcost1.copy()
        reportcost = reportcost.fillna('')
        #检查费项目号
        if  len(reportcost[(reportcost['项目号'].str.contains(' |sz|Sz|Sz|j|s'))&(reportcost['项目号'].str.len()>5)])>1:
            screm.insert(INSERT, '\n费-存在编辑异常项目号如下', '\n')
            screm.insert(INSERT, reportcost[(reportcost['项目号'].str.contains(' |sz|Sz|Sz|j|s'))&(reportcost['项目号'].str.len()>5)], '\n')
            window.update()
        ###输出费异常项目号号
        reportcost['核算项目号']=reportcost['核算项目号'].fillna('').astype(str)
        reportcost_err=reportcost[(reportcost['核算项目号'].str.len()<2)].drop_duplicates(subset=['项目号']).reset_index(drop=True)

        screm.insert(INSERT, '\n**********新增费-存在未参与核算项目号'+str(len(reportcost_err))+'个', '\n')
        window.update()

        time_cost = time.time()
        screm.insert(INSERT, '\n处理費:%d秒' % (time_cost - time_work), '\n')
        screm.insert(INSERT, '\n第一阶段执行时长:%d秒' % (time_cost - time_start), '\n')
        window.update()
        person_list = list(report_item['项目财经'].drop_duplicates())

#################################################################################################################################################################读取历史数据
        screm.insert(INSERT, '\n二、第二阶段--读取历史明细', '\n')
        window.update()
        screm.insert(INSERT, '\n2.1：正在读取历史采购PO', '\n')
        window.update()
        dir_po = r'数据源\历史核算明细\采购PO'
        index = 0
        report_po = []
        filename_po = os.listdir(dir_po)
        for i in filename_po:
            if '~$' in i:
                filename_po.remove(i)
        if os.listdir(dir_po):
            for name in filename_po:
                if 'csv' in name and '~$' not in name:
                    with open(os.path.join(dir_po, name), encoding='gb18030', errors='ignore') as f:
                        df3_1 = pd.read_csv(f, sep=',', error_bad_lines=False, low_memory=False, thousands=',')
                        if '产品线编号' in list(df3_1.columns):
                            df3_1 = df3_1.rename(columns={'产品线编号': '产品线编码'})
                        if '采购人员工号' in df3_1.columns:
                            del df3_1['采购人员工号']
                        if '审核员' in df3_1.columns:
                            del df3_1['审核员']
                        if '开单员' in df3_1.columns:
                            del df3_1['开单员']
                        if 'Unnamed: 55' in df3_1.columns:  # Unnamed: 56
                            del df3_1['Unnamed: 55']
                        if 'Unnamed: 56' in df3_1.columns:  # Unnamed: 56
                            del df3_1['Unnamed: 56']
                    report_po.append(df3_1)
                    index += 1  # 为了查看合并到第几个表格了
            old_report_po2 = pd.concat(report_po).reset_index(drop=True)
            old_report_po2['采购日期'] = old_report_po2['采购日期'].fillna(default_date)
            old_report_po2['采购日期'] = pd.to_datetime(old_report_po2['采购日期'], errors='coerce')
            #####剔除历史两个月数据
            old_report_po2 = old_report_po2[
                old_report_po2['采购日期'] < datetime.strptime(po_time.get(), "%Y-%m-%d")].reset_index(drop=True)
            old_report_po2['数据-来源'] = old_report_po2['数据-来源'].fillna('')
            old_report_po2 = old_report_po2.loc[old_report_po2['数据-来源'].str.contains('调账') == False].reset_index(
                drop=True)
            old_report_po2.loc[old_report_po2['数据-来源'].str.contains('数据'), '数据-来源'] = '历史数据'

        if os.listdir(dir_po) == False:
            old_report_po2 = pd.DataFrame(columns=['核算项目号', "采购单号", "采购日期", "采购单类型", "料件编号", "品名", "规格"
                , "采购数量", "未税单价", "采购金额-未税", "单价"
                , "含税金额", "税率", "已收货量", "未交量", "备注", "项目编号", "采购人员", "采购供应商"
                , "库存管理特征", "行状态", "状态", "作业编号", "模组标识", "模组名称", "品牌"
                , "零件类型", "产能", "标准/非标", "客户简称", "大项目名称", "大项目号", "设备名称", '产品线编码'
                , "产品线名称", '工艺', "项目财经", "项目财经再分类"
                , "验退量", "仓退换货量", "仓退量", "项目号整理", "留置原因说明"
                , "采购部门", "料件分类", "来源单号", '数据-来源'])

        ##############################################采购PO调账库
        screm.insert(INSERT, '\n2-1-1：正在读取采购PO调账库', '\n')
        window.update()
        dir_po1 = r'数据源\调账库\采购PO'
        index = 0
        report_po1 = []
        filename_po1 = os.listdir(dir_po1)
        for i in filename_po1:
            if '~$' in i:
                filename_po1.remove(i)
        if os.listdir(dir_po1):
            for name in filename_po1:
                if 'csv' in name and '~$' not in name:
                    with open(os.path.join(dir_po1, name), encoding='gb18030', errors='ignore') as f:
                        df3_2 = pd.read_csv(f, sep=',', error_bad_lines=False, low_memory=False, thousands=',')
                        if '产品线编号' in list(df3_1.columns):
                            df3_2 = df3_2.rename(columns={'产品线编号': '产品线编码'})
                    report_po1.append(df3_2)
                    index += 1  # 为了查看合并到第几个表格了
            old_report_po1 = pd.concat(report_po1).reset_index(drop=True)
            old_report_po1['采购日期'] = old_report_po1['采购日期'].fillna(default_date)
            old_report_po1['采购日期'] = pd.to_datetime(old_report_po1['采购日期'], errors='coerce')
            old_report_po1['数据-来源'] = '调账库'
        if os.listdir(dir_po1) == False:
            old_report_po1 = pd.DataFrame(columns=['核算项目号', "采购单号", "采购日期", "采购单类型", "料件编号", "品名", "规格"
                , "采购数量", "未税单价", "采购金额-未税", "单价"
                , "含税金额", "税率", "已收货量", "未交量", "备注", "项目编号", "采购人员", "采购供应商"
                , "库存管理特征", "行状态", "状态", "作业编号", "模组标识", "模组名称", "品牌"
                , "零件类型", "产能", "标准/非标", "客户简称", "大项目名称", "大项目号", "设备名称", '产品线编码'
                , "产品线名称", '工艺', "项目财经", "项目财经再分类"
                , "验退量", "仓退换货量", "仓退量", "项目号整理", "留置原因说明"
                , "采购部门", "料件分类", "来源单号", '数据-来源'])
        ######采购PO调账库
        ###合并调账库和历史
        if len(old_report_po2) > 0:
            old_report_po = pd.concat([old_report_po2, old_report_po1], ignore_index=True).reset_index(drop=True)
            ###输出剔除
            process_delete_out = pd.concat([process_delete, old_report_po1], ignore_index=True).reset_index(drop=True)
            ###输出剔除
        if len(old_report_po2) <= 0:
            old_report_po = old_report_po1.copy()
        ###合并调账库和历史
        process['数据分布'] = '新增'
        old_report_po['数据分布'] = '历史'
        if '工单开立时间' in old_report_po.columns:
            del old_report_po['工单开立时间']
        if '工单完工时间' in old_report_po.columns:
            del old_report_po['工单完工时间']
        if '实际验收时间' in old_report_po.columns:
            del old_report_po['实际验收时间']
        process = pd.concat([process, old_report_po], ignore_index=True).reset_index(drop=True)

        ###===================================================================================================================料
        screm.insert(INSERT, '\n2-2：正在读取历史的料', '\n')
        window.update()
        dir_mater= r'数据源\历史核算明细\料'
        index = 0
        report_mater = []
        filename_mater = os.listdir(dir_mater)
        for i in filename_mater:
            if '~$' in i:
                filename_mater.remove(i)
        if  os.listdir(dir_mater):
            for name in filename_mater:
                if 'csv' in name and '~$' not in name:
                    with open(os.path.join(dir_mater, name), encoding='utf-8', errors='ignore') as f:
                        df10_1 = pd.read_csv(f, sep=',', error_bad_lines=False, low_memory=False, thousands=',')
                        if '产品线编号' in list(df10_1.columns):
                            df10_1 = df10_1.rename(columns={'产品线编号': '产品线编码'})
                        if '单据日期' in list(df10_1.columns):
                            df10_1 = df10_1.rename(columns={'单据日期': '扣账日期'})
                        if '理由说明' in list(df10_1.columns):
                            df10_1 = df10_1.rename(columns={'理由说明': '单别'})
                        if '发料单号' in list(df10_1.columns):
                            df10_1 = df10_1.rename(columns={'发料单号': '发退料单号'})
                        if '作业' in list(df10_1.columns):
                            df10_1 = df10_1.rename(columns={'作业': '作业编号'})
                        if '库存管理特征' in list(df10_1.columns):
                            del df10_1['库存管理特征']
                        if 'Unnamed: 42' in list(df10_1.columns): #Unnamed: 42	Unnamed: 43	Unnamed: 31	Unnamed: 32	Unnamed: 33	Unnamed: 34
                            del df10_1['Unnamed: 42']
                        if 'Unnamed: 43' in list(df10_1.columns): #Unnamed: 42	Unnamed: 43	Unnamed: 31	Unnamed: 32	Unnamed: 33	Unnamed: 34
                            del df10_1['Unnamed: 43']
                        if 'Unnamed: 31' in list(df10_1.columns): #Unnamed: 42	Unnamed: 43	Unnamed: 31	Unnamed: 32	Unnamed: 33	Unnamed: 34
                            del df10_1['Unnamed: 31']
                        if 'Unnamed: 32' in list(df10_1.columns): #Unnamed: 42	Unnamed: 43	Unnamed: 31	Unnamed: 32	Unnamed: 33	Unnamed: 34
                            del df10_1['Unnamed: 32']
                        if 'Unnamed: 33' in list(df10_1.columns): #Unnamed: 42	Unnamed: 43	Unnamed: 31	Unnamed: 32	Unnamed: 33	Unnamed: 34
                            del df10_1['Unnamed: 33']
                        if 'Unnamed: 34' in list(df10_1.columns): #Unnamed: 42	Unnamed: 43	Unnamed: 31	Unnamed: 32	Unnamed: 33	Unnamed: 34
                            del df10_1['Unnamed: 34']
                    report_mater.append(df10_1)
                    index += 1  # 为了查看合并到第几个表格了
            old_report_mater2= pd.concat(report_mater).reset_index(drop=True)
            if '工单开立时间' in old_report_mater2.columns:
                del old_report_mater2['工单开立时间' ]
            if '工单完工时间' in old_report_mater2.columns:
                del old_report_mater2['工单完工时间']
            if '实际验收时间' in old_report_mater2.columns:
                del old_report_mater2['实际验收时间']

            default_date = '1990/01/01'
            old_report_mater2['扣账日期'] = old_report_mater2['扣账日期'].fillna(default_date)
            old_report_mater2['扣账日期'] = pd.to_datetime(old_report_mater2['扣账日期'], errors='coerce')
            old_report_mater2['领料类型'] = old_report_mater2['领料类型'].fillna('')
            old_report_mater2=old_report_mater2[old_report_mater2['扣账日期']<datetime.strptime(mater_time.get(),"%Y-%m-%d")].reset_index(drop=True)
            '''
            old_report_mater2['删在制'] = ''
            old_report_mater2.loc[(old_report_mater2['扣账日期'] >= pd.Timestamp(2023, 3, 1)) & (
                old_report_mater2['领料类型'].str.contains('在制')), '删在制'] = '删除'
            old_report_mater2 = old_report_mater2[old_report_mater2['删在制'].str.contains('删除') == False].reset_index(drop=True)
            del old_report_mater2['删在制']
            '''
            old_report_mater2['数据-来源'] = old_report_mater2['数据-来源'].fillna('')
            old_report_mater2 = old_report_mater2[old_report_mater2['数据-来源'].str.contains('调账') == False].reset_index(drop=True)
            old_report_mater2.loc[old_report_mater2['数据-来源'].str.contains('数据'), '数据-来源'] = '历史数据'

        if os.listdir(dir_mater)==False:
            old_report_mater2=pd.DataFrame(
                columns=["核算项目号", "项目编号", "扣账日期", "发退料单号", "料号", "品名"
                          , "规格", '仓库说明', '单位', "数量", "未税单价", "未税金额"
                          , "工单号码","备注",  '单别',"作业编号","模组标识", "模组名称"
                          , "品牌/供应商", "零件类型", "产能", "标件/非标件", "客户简称",'产品线编码', "产品线名称","大项目号",'大项目名称'
                          , "设备名称",'工艺',"项目号整理",'成品料号','是否有工单','项目财经',"领料类型",'项目财经再分类','数据-来源'])
        ######--------------------------------------------------------------------------合并历史+最新的料
        ####读取调账库
        screm.insert(INSERT, '\n2-2-1:正在读取料-调账库', '\n')
        window.update()
        dir_mater1 = r'数据源\调账库\料'
        index = 0
        report_mater1 = []
        filename_mater1 = os.listdir(dir_mater1)
        for i in filename_mater1:
            if '~$' in i:
                filename_mater1.remove(i)
        if os.listdir(dir_mater1):
            for name in filename_mater1:
                if 'csv' in name and '~$' not in name:
                    with open(os.path.join(dir_mater1, name), encoding='gb18030', errors='ignore') as f:
                        df10_2 = pd.read_csv(f, sep=',', error_bad_lines=False, low_memory=False, thousands=',')
                        if '产品线编号' in list(df10_2.columns):
                            df10_2 = df10_2.rename(columns={'产品线编号': '产品线编码'})
                    report_mater1.append(df10_2)
                    index += 1  # 为了查看合并到第几个表格了
            old_report_mater1 = pd.concat(report_mater1).reset_index(drop=True)
        old_report_mater1['数据-来源'] = '调账库'
        ####读取调账库
        old_report_mater = pd.concat([old_report_mater2, old_report_mater1], ignore_index=True).reset_index(drop=True)

        ###输出剔除
        report_material_delete_out=pd.concat([report_material_delete, old_report_mater1], ignore_index=True).reset_index(drop=True)
        ###输出剔除
        ####去掉在制三月数据4.29
        report_material['数据分布'] = '新增'
        old_report_mater['数据分布'] = '历史'
        report_material=pd.concat([report_material,old_report_mater], ignore_index=True).reset_index(drop=True)
        ###===================================================================================================================工
        screm.insert(INSERT, '\n2-3：正在读取历史的工', '\n')
        window.update()
        dir_job = r'数据源\历史核算明细\工'
        index = 0
        report_job = []
        filename_job = os.listdir(dir_job)
        for i in filename_job:
            if '~$' in i:
                filename_job.remove(i)
        if os.listdir(dir_job):
            for name in filename_job:
                if 'csv' in name and '~$' not in name:
                    with open(os.path.join(dir_job, name), encoding='gb18030', errors='ignore') as f:
                        df10_2= pd.read_csv(f, sep=',', error_bad_lines=False, low_memory=False, thousands=',')
                        if '产品线编号' in list(df10_2.columns):
                            df10_2= df10_2.rename(columns={'产品线编号': '产品线编码'})
                    report_job.append(df10_2)
                    index += 1  # 为了查看合并到第几个表格了
            old_report_work = pd.concat(report_job).reset_index(drop=True)
            if '工单开立时间' in old_report_work.columns:
                del old_report_work['工单开立时间' ]
            if '工单完工时间' in old_report_work.columns:
                del old_report_work['工单完工时间']
            if '实际验收时间' in old_report_work.columns:
                del old_report_work['实际验收时间']
        if os.listdir(dir_job)==False:
            old_report_work=pd.DataFrame(
                columns=["核算项目号","项目号","姓名","工号","人员归属","成本归属","事业部","部门","科室","岗位","工作地点","提报人","报工来源","报工工单号","工单标识","完成日期",'工种',"工时","工价","工时成本","备注","阶段","报工组别","工种再分类","客户简称","产品线编码","产品线名称","大项目号","大项目名称","设备名称","产能","自制/外包",'工艺',"项目号整理","项目财经",'项目财经再分类'])
        ####按需求剔除历史部分历史工
        old_report_work['完成日期'] = old_report_work['完成日期'].fillna(default_date)
        old_report_work['完成日期']= pd.to_datetime(old_report_work['完成日期'], errors='coerce')
        old_report_work = old_report_work[old_report_work['完成日期'] < datetime.strptime(work_time.get(), "%Y-%m-%d")].reset_index(drop=True)
        ######--------------------------------------------------------------------------合并历史+最新的料
        reportwork['数据分布'] = '新增'
        old_report_work['数据分布'] = '历史'
        reportwork= pd.concat([reportwork, old_report_work], ignore_index=True).reset_index(drop=True)
        ###===================================================================================================================费
        screm.insert(INSERT, '\n2-4：正在读取历史的费', '\n')
        window.update()
        dir_cost= r'数据源\历史核算明细\费'
        index = 0
        report_cost = []
        filename_cost = os.listdir(dir_cost)
        for i in filename_cost:
            if '~$' in i:
                filename_cost.remove(i)
        if os.listdir(dir_cost):
            for name in filename_cost:
                if 'csv' in name and '~$' not in name:
                    screm.insert(INSERT, '\n正在读取第' + str(index + 1) + '份历史的费:' + name, '\n')
                    window.update()
                    with open(os.path.join(dir_cost, name), encoding='gb18030', errors='ignore') as f:
                        df10_3 = pd.read_csv(f, sep=',', error_bad_lines=False, low_memory=False, thousands=',')
                        if '产品线编号' in list(df10_3.columns):
                            df10_3 = df10_3.rename(columns={'产品线编号': '产品线编码'})
                    report_cost.append(df10_3)
                    index += 1  # 为了查看合并到第几个表格了
            old_report_cost= pd.concat(report_cost).reset_index(drop=True)
            old_report_cost['是否核算']=old_report_cost['是否核算'].fillna('')
            if '工单开立时间' in old_report_cost.columns:
                del old_report_cost['工单开立时间' ]
            if '工单完工时间' in old_report_cost.columns:
                del old_report_cost['工单完工时间']
            if '实际验收时间' in old_report_cost.columns:
                del old_report_cost['实际验收时间']
            old_report_cost['金额(万元)'] = old_report_cost['金额(万元)'].fillna(0)
            old_report_cost.loc[old_report_cost['是否核算'].str.contains('否'),'金额(万元)']=0
            old_report_cost['金额万元'] = old_report_cost['金额(万元)'] * 10000 #金额(万元)
        if os.listdir(dir_cost)==False:
            old_report_cost= pd.DataFrame(columns=['核算项目号','年份',"月份","凭证编号","摘要","科目编码","科目名称","部门名称","项目号","金额","中心","汇总科目编码","费用类型","公司名称","事业部","事业部重分类","部门重分类","科室重分类","经管费用归属","经管科目一级","经管科目二级","金额(万元)","项目类型","无项目报工率","备注",'项目号整理',"有无项目号","费用大类","费用小类","是否核算","客户简称","产品线编码","产品线名称","大项目号","大项目名称","设备名称","产能",'工艺',"项目财经","项目财经再分类",'金额万元'])
        ####费没有独立日期字段需生成
        old_report_cost['月份1'] = old_report_cost['月份'].copy().astype(str)
        old_report_cost['月份1'] = old_report_cost['月份1'].fillna('01')
        old_report_cost.loc[old_report_cost['月份1'].str.len() == 1, '月份1'] = str('0') + old_report_cost['月份1']
        old_report_cost['申请时间'] = pd.to_datetime(
            old_report_cost['年份'].astype(str).str.replace('年', '', regex=True) + '-' + old_report_cost[
                '月份1'] + '-' + str('01'), errors='coerce')
        old_report_cost['申请时间'] = old_report_cost['申请时间'] + pd.DateOffset(months=1) - pd.DateOffset(days=1)
        del old_report_cost['月份1']

        screm.insert(INSERT, '\n2-5：按需求剔除四表部分历史数据', '\n')
        window.update()
        ####按需求剔除历史部分历史费
        old_report_cost['申请时间'] =old_report_cost['申请时间'].fillna(default_date)
        old_report_cost['申请时间'] = pd.to_datetime(old_report_cost['申请时间'], errors='coerce')
        old_report_cost = old_report_cost[old_report_cost['申请时间'] < datetime.strptime(cost_time.get(), "%Y-%m-%d")].reset_index(drop=True)
        ######--------------------------------------------------------------------------合并历史+最新的料
        reportcost['数据分布']='新增'
        old_report_cost['数据分布'] = '历史'
        reportcost = pd.concat([reportcost, old_report_cost], ignore_index=True).reset_index(drop=True)

        proj_name=report_item[['项目号整理','项目财经','项目财经再分类','核算项目号']]
        del process['项目财经']
        del process['项目财经再分类']

        del report_material['项目财经']
        del report_material['项目财经再分类']

        del reportwork['项目财经']
        del reportwork['项目财经再分类']

        del reportcost['项目财经']
        del reportcost['项目财经再分类']

        screm.insert(INSERT, '\n2-6：所有数据拉取对照表数据', '\n')
        window.update()
        process['项目号整理']=process['项目号整理'].fillna('')
        report_material['项目号整理'] = report_material['项目号整理'].fillna('')
        reportwork['项目号整理'] = reportwork['项目号整理'].fillna('')
        reportcost['项目号整理'] =reportcost['项目号整理'].fillna('')
        process[['项目财经','项目财经再分类','核算项目号']]=pd.merge(process,proj_name,on='项目号整理',how='left')[['项目财经','项目财经再分类','核算项目号_y']]
        report_material[['项目财经', '项目财经再分类','核算项目号']] = pd.merge(report_material, proj_name, on='项目号整理', how='left')[['项目财经', '项目财经再分类','核算项目号_y']]
        reportwork[['项目财经', '项目财经再分类','核算项目号']] = pd.merge(reportwork, proj_name, on='项目号整理', how='left')[['项目财经', '项目财经再分类','核算项目号_y']]
        reportcost[['项目财经', '项目财经再分类','核算项目号']] = pd.merge(reportcost, proj_name, on='项目号整理', how='left')[['项目财经','项目财经再分类','核算项目号_y']]
        ##############&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&四个日期处理

        process['采购日期']=process['采购日期'].fillna(default_date)
        report_material['扣账日期'] = report_material['扣账日期'].fillna(default_date)
        reportwork['完成日期'] = reportwork['完成日期'].fillna(default_date)
        process['采购日期']= pd.to_datetime(process['采购日期'],errors='coerce')
        report_material['扣账日期'] = pd.to_datetime(report_material['扣账日期'], errors='coerce')
        reportwork['完成日期']= pd.to_datetime(reportwork['完成日期'], errors='coerce')

        ####费没有独立日期字段需生成
        reportcost['月份1'] = reportcost['月份'].copy().astype(str)
        reportcost['月份1'] = reportcost['月份1'].fillna('01')
        reportcost.loc[reportcost['月份1'].str.len() == 1, '月份1'] = str('0') + reportcost['月份1']
        reportcost['申请时间'] = pd.to_datetime(reportcost['年份'].astype(str).str.replace('年', '', regex=True) + '-' + reportcost['月份1'] + '-' + str('01'),errors='coerce')
        reportcost['申请时间'] = reportcost['申请时间'] + pd.DateOffset(months=1) - pd.DateOffset(days=1)
        del reportcost['月份1']
        ##########################生成历史里找不到的
        process[['项目财经','项目财经再分类']]= process[['项目财经','项目财经再分类']].fillna('')
        report_material[['项目财经','项目财经再分类']] = report_material[['项目财经','项目财经再分类']].fillna('')
        reportwork[['项目财经','项目财经再分类']]  = reportwork[['项目财经','项目财经再分类']] .fillna('')
        reportcost[['项目财经', '项目财经再分类']] = reportcost[['项目财经', '项目财经再分类']].fillna('')
        '''
        old_po_err=process[(process['项目财经']=='')&(process['数据分布']=='历史')].reset_index(drop=True)[['项目号整理','项目编号','核算项目号']]
        old_mater_err = report_material[(report_material['项目财经'] == '') & (report_material['数据分布'] == '历史')].reset_index(
            drop=True)[['项目号整理', '项目编号', '核算项目号']]
        old_work_err = reportwork[(reportwork['项目财经'] == '') & (reportwork['数据分布'] == '历史')].reset_index(drop=True)[['项目号整理', '项目号', '核算项目号']]
        old_cost_err = reportcost[(reportcost['项目财经'] == '') & (reportcost['数据分布'] == '历史')].reset_index(drop=True)[['项目号整理', '项目号', '核算项目号']]
        '''
        time_old=time.time()
        screm.insert(INSERT, '\n第二阶段执行时长:%d秒' % (time_old - time_cost), '\n')

##################################################################################################################################################制作汇总表
        screm.insert(INSERT, '\n三、第三阶段--处理辅助表并汇入明细表', '\n')
        window.update()
#######################################
        screm.insert(INSERT, '\n3.1：正在处理台账出货、验收表', '\n')
        window.update()
        path_bill_time = r'数据源\汇总表所需数据源\台账'
        bill_time_file = os.listdir(path_bill_time)
        for i in bill_time_file:
            if '~$' in i:
                bill_time_file.remove(i)
        for name in bill_time_file:
            if 'xlsm' in name and '~$' not in name:
                report_bill_out = pd.read_excel(path_bill_time + '\\' + name, sheet_name='2.出货')[
                    ['大项目号', '项目号', '实际数量', '实际出货日期', '料号']]
                report_bill_rece = pd.read_excel(path_bill_time + '\\' + name, sheet_name='3.验收')[
                    ['大项目号', '项目号', '数量', '终验收时间', '料号']]

        report_bill_out[['大项目号', '项目号', '料号']]= report_bill_out[['大项目号', '项目号', '料号']].fillna('')
        report_bill_rece[['大项目号', '项目号','料号']]=report_bill_rece[['大项目号', '项目号','料号']].fillna('')
        if len(report_bill_out[report_bill_out['项目号'].str.contains('-')]) > 0:
            report_bill_out['项目号整理'] = ''
            report_bill_out['项目整'] = report_bill_out['项目号'].str.split('-', expand=True)[0]
            report_bill_out['项目整1'] = report_bill_out['项目号'].str.split('-', expand=True)[1]
            report_bill_out['项目整1'] = report_bill_out['项目整1'].fillna('空值')
            report_bill_out['项目号整理'] = report_bill_out['项目整']
            report_bill_out.loc[(report_bill_out['项目整1'].str.isdigit()) | (report_bill_out['项目整1'].str.contains('SH')), '项目号整理'] = \
            report_bill_out['项目号整理'] + '-' + report_bill_out['项目整1']
        if len(report_bill_out[report_bill_out['项目号'].str.contains('-')]) == 0:
            report_bill_out['项目号整理'] = report_bill_out['项目号']
        report_bill_out.loc[(report_bill_out['项目号整理'].str[0] == 'F') & (report_bill_out['项目号整理'].str[:3].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX')), '项目号整理'] = \
        report_bill_out['项目号整理'].str[3:]
        report_bill_out.loc[(report_bill_out['项目号整理'].str[:2].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX',na=False)), '项目号整理'] = report_bill_out['项目号整理'].str[2:]

        if len(report_bill_rece[report_bill_rece['项目号'].str.contains('-')]) > 0:
            report_bill_rece['项目号整理'] = ''
            report_bill_rece['项目整'] = report_bill_rece['项目号'].str.split('-', expand=True)[0]
            report_bill_rece['项目整1'] = report_bill_rece['项目号'].str.split('-', expand=True)[1]
            report_bill_rece['项目整1'] = report_bill_rece['项目整1'].fillna('空值')
            report_bill_rece['项目号整理'] = report_bill_rece['项目整']
            report_bill_rece.loc[(report_bill_rece['项目整1'].str.isdigit()) | (report_bill_rece['项目整1'].str.contains('SH')), '项目号整理'] = \
            report_bill_rece['项目号整理'] + '-' + report_bill_rece['项目整1']
        if len(report_bill_rece[report_bill_rece['项目号'].str.contains('-')]) == 0:
            report_bill_rece['项目号整理'] = report_bill_rece['项目号']
        report_bill_rece.loc[(report_bill_rece['项目号整理'].str[0] == 'F') & (report_bill_rece['项目号整理'].str[:3].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX')), '项目号整理'] = \
        report_bill_rece['项目号整理'].str[3:]
        report_bill_rece.loc[(report_bill_rece['项目号整理'].str[:2].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX',na=False)), '项目号整理'] = report_bill_rece['项目号整理'].str[2:]

        ##扣除内部关联交易数据
        report_bill_out['实际数量'] = report_bill_out['实际数量'].fillna(0)
        report_bill_out[['大项目号', '项目号', '料号','项目号整理']] = report_bill_out[['大项目号', '项目号', '料号','项目号整理']].fillna('')
        report_bill_out = report_bill_out[report_bill_out['料号'].str[:5].str.contains('311-') == False][['大项目号', '项目号整理', '实际数量', '实际出货日期']].reset_index(drop=True)
        default_date = pd.Timestamp(1990, 1, 1)
        report_bill_out['实际出货日期'] = report_bill_out['实际出货日期'].fillna(default_date)
        report_bill_out['实际出货日期'] = pd.to_datetime(report_bill_out['实际出货日期'], errors='coerce')
        report_bill_out = report_bill_out[(report_bill_out['实际出货日期'] != pd.Timestamp(1990, 1, 1)) & (report_bill_out['项目号整理'] != '')].reset_index(drop=True)

        report_bill_rece['数量'] = report_bill_rece['数量'].fillna(0)
        report_bill_rece[['大项目号', '项目号整理', '料号']] = report_bill_rece[['大项目号','项目号整理', '料号']].fillna('')
        report_bill_rece=report_bill_rece[report_bill_rece['料号'].str[:5].str.contains('311-') == False][['大项目号', '项目号整理', '数量', '终验收时间']].reset_index(drop=True)
        default_date = pd.Timestamp(1990, 1, 1)
        report_bill_rece['终验收时间'] = report_bill_rece['终验收时间'].fillna(default_date)
        report_bill_rece['终验收时间'] = pd.to_datetime(report_bill_rece['终验收时间'], errors='coerce')
        report_bill_rece = report_bill_rece.rename(columns={'终验收时间': '实际验收时间'})

        report_bill_rece = report_bill_rece[(report_bill_rece['实际验收时间'] != pd.Timestamp(1990, 1, 1)) & (report_bill_rece['项目号整理'] != '')].reset_index(drop=True)
        report_bill_rece=report_bill_rece.sort_values(by=['实际验收时间'],ascending=False).reset_index(drop=True)
        report_bill_rece2=report_bill_rece.drop_duplicates(subset=['项目号整理'],keep='first').reset_index(drop=True)
        report_bill_rece1= report_bill_rece.drop_duplicates(subset=['大项目号','项目号整理'], keep='first').reset_index(drop=True)
        report_bill_out = report_bill_out.rename(columns={'实际出货日期': '实际出货时间'})
        report_bill_out = report_bill_out.sort_values(by=['实际出货时间'], ascending=False).reset_index(drop=True)
        report_bill_out2 = report_bill_out.drop_duplicates(subset=['项目号整理'], keep='first').reset_index(drop=True)
        report_bill_out1 = report_bill_out.drop_duplicates(subset=['大项目号', '项目号整理'], keep='first').reset_index(drop=True)

############################################
        screm.insert(INSERT, '\n3.2：正在处理工单开立表', '\n')
        window.update()
        need_time = ['大项目号', '工单单号', '项目编号', '单据日期', '过账日期', '单号', '关联的一般工单', '生产料号', '生产数量', '状态码']
        path_time_start = r'数据源\汇总表所需数据源\工单开立时间'
        file_time_start = os.listdir(path_time_start)
        for i in file_time_start:
            if '~$' in i:
                file_time_start.remove(i)
        report_time_star = []
        if os.listdir(path_time_start):
            for name in file_time_start:
                if 'csv' in name and '~$' not in name:
                    with open(os.path.join(path_time_start, name), encoding='utf-8', errors='ignore') as f:
                        df12 = pd.read_csv(f, sep=',', error_bad_lines=False, low_memory=False, thousands=',')
                        # df12 = df12[:-4]
                        for col in df12.columns:
                            if col not in need_time:
                                del df12[col]
                        report_time_star.append(df12)
        report_time_start = pd.concat(report_time_star).reset_index(drop=True)
        report_time_start[['状态码', '大项目号', '工单单号', '项目编号', '关联的一般工单', '生产料号']] = report_time_start[
            ['状态码', '大项目号', '工单单号', '项目编号', '关联的一般工单', '生产料号']].fillna('')
        report_time_start = report_time_start[report_time_start['状态码'].str.contains('作废') == False].reset_index(drop=True)
        report_time_start['生产数量'] = report_time_start['生产数量'].fillna(0)

        ##########做项目号整理
        if len(report_time_start[report_time_start['项目编号'].str.contains('-')]) > 0:
            report_time_start['项目号整理'] = ''
            report_time_start['项目整'] = report_time_start['项目编号'].str.split('-', expand=True)[0]
            report_time_start['项目整1'] = report_time_start['项目编号'].str.split('-', expand=True)[1]
            report_time_start['项目整1'] = report_time_start['项目整1'].fillna('空值')
            report_time_start['项目号整理'] = report_time_start['项目整']
            report_time_start.loc[(report_time_start['项目整1'].str.isdigit()) | (report_time_start['项目整1'].str.contains('SH')), '项目号整理'] = \
            report_time_start['项目号整理'] + '-' + report_time_start['项目整1']
        if len(report_time_start[report_time_start['项目编号'].str.contains('-')]) == 0:
            report_time_start['项目号整理'] = report_time_start['项目编号']
        report_time_start.loc[(report_time_start['项目号整理'].str[0] == 'F') & (
            report_time_start['项目号整理'].str[:3].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX')), '项目号整理'] = \
        report_time_start['项目号整理'].str[3:]
        report_time_start.loc[(report_time_start['项目号整理'].str[:2].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX',na=False)), '项目号整理'] = report_time_start['项目号整理'].str[2:]
        #####去掉604-，26A的工单
        report_time_start = report_time_start[
            report_time_start['工单单号'].str.contains('604-|26A-|606-') == False].reset_index(drop=True)  ##602
        report_time_start = report_time_start.rename(columns={'单据日期': '工单开立时间'})
        # ------给到明细表
        screm.insert(INSERT, '\n3.3：正在处理工单完工表', '\n')
        window.update()
        need_time = ['大项目号', '工单单号', '项目编号', '单据日期','过账日期', '单号','数量']
        path_time_end = r'数据源\汇总表所需数据源\工单完工时间'
        file_time_end = os.listdir(path_time_end)
        for i in file_time_end:
            if '~$' in i:
                file_time_end.remove(i)
        #########后续数据源会修改
        report_time_en = []
        if os.listdir(path_time_end):
            for name in file_time_end:
                if 'csv' in name and '~$' not in name:
                    # print('读取费用报销表第' + str(index + 1) + '份' + name)
                    # df = pd.read_excel(os.path.join(path_cost, name), header=2,thousands=',')
                    with open(os.path.join(path_time_end, name), encoding='utf-8', errors='ignore') as f:
                        # 再解决部分报错行如 ParserError：Error tokenizing data.C error:Expected 2 fields in line 407,saw 3.
                        df13 = pd.read_csv(f, sep=',', error_bad_lines=False, low_memory=False, thousands=',')
                        #df13 = df13[:-2]
                        for col in df13.columns:
                            if col not in need_time:
                                del df13[col]
                    report_time_en.append(df13)
        report_time_end = pd.concat(report_time_en).reset_index(drop=True)
        report_time_end[['工单单号', '项目编号']] = report_time_end[['工单单号', '项目编号']].fillna('')
        if len(report_time_end[report_time_end['项目编号'].str.contains('-')]) > 0:
            report_time_end['项目号整理'] = ''
            report_time_end['项目整'] = report_time_end['项目编号'].str.split('-', expand=True)[0]
            report_time_end['项目整1'] = report_time_end['项目编号'].str.split('-', expand=True)[1]
            report_time_end['项目整1'] = report_time_end['项目整1'].fillna('空值')
            report_time_end['项目号整理'] = report_time_end['项目整']
            report_time_end.loc[
                (report_time_end['项目整1'].str.isdigit()) | (report_time_end['项目整1'].str.contains('SH')), '项目号整理'] = \
            report_time_end['项目号整理'] + '-' + report_time_end['项目整1']
        if len(report_time_end[report_time_end['项目编号'].str.contains('-')]) == 0:
            report_time_end['项目号整理'] = report_time_end['项目编号']

        report_time_end.loc[(report_time_end['项目号整理'].str[0] == 'F') & (report_time_end['项目号整理'].str[:3].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX')), '项目号整理'] = \
        report_time_end['项目号整理'].str[3:]
        report_time_end.loc[(report_time_end['项目号整理'].str[:2].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX',na=False)), '项目号整理'] = report_time_end['项目号整理'].str[2:]
        #####去掉602-、604-，26A的工单
        report_time_end = report_time_end[report_time_end['工单单号'].str.contains('603-|604-|26A-') == False].reset_index(drop=True)
        report_time_end = report_time_end.rename(columns={'过账日期': '工单完工时间'})
        default_date1 = pd.Timestamp(2090, 1, 1)
        report_time_end['工单完工时间'] = report_time_end['工单完工时间'].fillna(default_date1)
        report_time_end['工单完工时间']= pd.to_datetime(report_time_end['工单完工时间'], errors='coerce')
        report_time_end_copy=report_time_end.copy()
        report_time_end=report_time_end.sort_values(by=["工单完工时间"],ascending=False).reset_index(drop=True)

        report_time_end = report_time_end.drop_duplicates(subset=['工单单号'],keep='first').reset_index(drop=True)

        report_time = pd.merge(report_time_start, report_time_end, on=['工单单号', '项目号整理'], how='left')
#######################################
        screm.insert(INSERT, '\n3.4：正在将工单信息汇入四个明细表', '\n')
        window.update()
        ############工单开立-完结整理，一刀切，每个项目选择最晚开单日期，选出来如有多个完工日期，选择最晚，如果存在空的定为M4阶段
        default_date1 = pd.Timestamp(2090, 1, 1)
        report_time['工单完工时间']= report_time['工单完工时间'].fillna(default_date1)
        report_time['工单开立时间']= report_time['工单开立时间'].fillna(default_date)
        report_time['工单开立时间'] = pd.to_datetime(report_time['工单开立时间'], errors='coerce')
        report_time['工单完工时间'] = pd.to_datetime(report_time['工单完工时间'], errors='coerce')
        ###'工单开立时间','工单完工时间'按两层给到明细表，1、工单单号--同一维度选最晚时间2、核算项目号--同一维度选最晚时间
        report_drop_time2 = report_time.groupby(['项目号整理']).agg({'工单开立时间': "min", '工单完工时间': "min"}).add_suffix('').reset_index()
        report_drop_time1 = report_time.groupby(['工单单号']).agg({'工单开立时间': "min", '工单完工时间': "min"}).add_suffix('').reset_index()
        ##给到明细表工单两个时间
        # 第一次
        report_material[['工单开立时间1', '工单完工时间1']] = pd.merge(report_material, report_drop_time1, left_on='工单号码', right_on='工单单号', how='left')[['工单开立时间', '工单完工时间']]
        reportwork[['工单开立时间1', '工单完工时间1']] =pd.merge(reportwork, report_drop_time1, left_on='报工工单号', right_on='工单单号', how='left')[['工单开立时间', '工单完工时间']]
        # 第二层
        process[['工单开立时间', '工单完工时间']] = pd.merge(process, report_drop_time2, left_on='项目号整理', right_on='项目号整理', how='left')[['工单开立时间', '工单完工时间']]
        report_material[['工单开立时间', '工单完工时间']] =pd.merge(report_material, report_drop_time2, left_on='项目号整理', right_on='项目号整理', how='left')[['工单开立时间', '工单完工时间']]
        reportwork[['工单开立时间', '工单完工时间']] = pd.merge(reportwork, report_drop_time2, left_on='项目号整理', right_on='项目号整理', how='left')[['工单开立时间', '工单完工时间']]
        reportcost[['工单开立时间', '工单完工时间']] =pd.merge(reportcost, report_drop_time2, left_on='项目号整理', right_on='项目号整理', how='left')[['工单开立时间', '工单完工时间']]
        # 一二层合并
        default_date1 = pd.Timestamp(1990, 1, 1)
        reportwork[['工单开立时间1', '工单完工时间1']] = reportwork[['工单开立时间1', '工单完工时间1']].fillna(default_date1)
        report_material[['工单开立时间1', '工单完工时间1']] = report_material[['工单开立时间1', '工单完工时间1']].fillna(default_date1)
        report_material.loc[report_material['工单开立时间1'] != pd.Timestamp(1990, 1, 1), '工单开立时间'] = report_material['工单开立时间1']
        report_material.loc[report_material['工单完工时间1'] != pd.Timestamp(1990, 1, 1), '工单完工时间'] = report_material['工单完工时间1']
        reportwork.loc[reportwork['工单开立时间1'] != pd.Timestamp(1990, 1, 1), '工单开立时间'] = reportwork['工单开立时间1']
        reportwork.loc[reportwork['工单完工时间1'] != pd.Timestamp(1990, 1, 1), '工单完工时间'] = reportwork['工单完工时间1']
        del reportwork['工单开立时间1']
        del reportwork['工单完工时间1']
        del report_material['工单开立时间1']
        del report_material['工单完工时间1']
        ####给到明细表验收日期
#########################################
        screm.insert(INSERT, '\n3.5：正在将验收时间汇入四个明细表', '\n')
        window.update()
        process['验收日期1'] =pd.merge(process, report_bill_rece1, left_on=['大项目号', '项目号整理'], right_on=['大项目号', '项目号整理'], how='left')['实际验收时间']
        report_material['验收日期1'] =pd.merge(report_material, report_bill_rece1, left_on=['大项目号', '项目号整理'], right_on=['大项目号', '项目号整理'], how='left')['实际验收时间']
        reportwork['验收日期1'] =pd.merge(reportwork, report_bill_rece1, left_on=['大项目号', '项目号整理'], right_on=['大项目号', '项目号整理'], how='left')['实际验收时间']
        ##第二层
        process['验收日期'] = pd.merge(process,report_bill_rece2, left_on='项目号整理', right_on='项目号整理', how='left')['实际验收时间']
        report_material['验收日期'] =pd.merge(report_material,report_bill_rece2, left_on='项目号整理', right_on='项目号整理', how='left')['实际验收时间']
        reportwork['验收日期'] = pd.merge(reportwork,report_bill_rece2, left_on='项目号整理', right_on='项目号整理', how='left')[ '实际验收时间']
        reportcost['验收日期'] = pd.merge(reportcost,report_bill_rece2, left_on='项目号整理', right_on='项目号整理', how='left')['实际验收时间']
        reportcost['验收日期'] = reportcost['验收日期'].fillna(pd.Timestamp(1990, 1, 1))
        ##一二层合并
        process['验收日期1'] = process['验收日期1'].fillna(pd.Timestamp(1990, 1, 1))
        report_material['验收日期1'] = report_material['验收日期1'].fillna(pd.Timestamp(1990, 1, 1))
        reportwork['验收日期1'] = reportwork['验收日期1'].fillna(pd.Timestamp(1990, 1, 1))
      #  reportcost['验收日期1'] = reportcost['验收日期1'].fillna(pd.Timestamp(1990, 1, 1))
        process.loc[process['验收日期1'] != pd.Timestamp(1990, 1, 1), '验收日期'] = process['验收日期1']
        process['验收日期']=process['验收日期'].fillna(pd.Timestamp(1990, 1, 1))
        report_material.loc[report_material['验收日期1'] != pd.Timestamp(1990, 1, 1), '验收日期'] = report_material['验收日期1']
        report_material['验收日期'] = report_material['验收日期'].fillna(pd.Timestamp(1990, 1, 1))
        reportwork.loc[reportwork['验收日期1'] != pd.Timestamp(1990, 1, 1), '验收日期'] = reportwork['验收日期1']
        reportwork['验收日期'] = reportwork['验收日期'].fillna(pd.Timestamp(1990, 1, 1))
       # reportcost.loc[reportcost['验收日期1'] != pd.Timestamp(1990, 1, 1), '验收日期'] = reportcost['验收日期1']
        del process['验收日期1']
        del report_material['验收日期1']
        del reportwork['验收日期1']
        #######给到明细表实际出货日期
#########################################
        screm.insert(INSERT, '\n3.6：正在将出货时间汇入四个明细表', '\n')
        window.update()
        process['出货日期1'] =pd.merge(process, report_bill_out1, left_on=['大项目号', '项目号整理'], right_on=['大项目号', '项目号整理'], how='left')['实际出货时间']
        report_material['出货日期1'] = \
        pd.merge(report_material, report_bill_out1, left_on=['大项目号', '项目号整理'], right_on=['大项目号', '项目号整理'], how='left')[
            '实际出货时间']
        reportwork['出货日期1'] = \
        pd.merge(reportwork, report_bill_out1, left_on=['大项目号', '项目号整理'], right_on=['大项目号', '项目号整理'], how='left')[
            '实际出货时间']
       # reportcost['出货日期1'] =  pd.merge(reportcost, report_bill_out1, left_on=['大项目号', '项目号整理'], right_on=['大项目号', '项目号整理'], how='left')['实际出货时间']
        ##第二层
        process['出货日期'] = pd.merge(process, report_bill_out2, left_on='项目号整理', right_on='项目号整理', how='left')['实际出货时间']
        report_material['出货日期'] = pd.merge(report_material, report_bill_out2, left_on='项目号整理', right_on='项目号整理', how='left')['实际出货时间']
        reportwork['出货日期'] = pd.merge(reportwork, report_bill_out2, left_on='项目号整理', right_on='项目号整理', how='left')[
            '实际出货时间']
        reportcost['出货日期'] = pd.merge(reportcost, report_bill_out2, left_on='项目号整理', right_on='项目号整理', how='left')['实际出货时间']
        reportcost['出货日期']=reportcost['出货日期'].fillna(pd.Timestamp(1990, 1, 1))
        ##一二层合并
        process['出货日期1'] = process['出货日期1'].fillna(pd.Timestamp(1990, 1, 1))
        report_material['出货日期1'] = report_material['出货日期1'].fillna(pd.Timestamp(1990, 1, 1))
        reportwork['出货日期1'] = reportwork['出货日期1'].fillna(pd.Timestamp(1990, 1, 1))
       # reportcost['出货日期1'] = reportcost['出货日期1'].fillna(pd.Timestamp(1990, 1, 1))
        process.loc[process['出货日期1'] != pd.Timestamp(1990, 1, 1), '出货日期'] = process['出货日期1']
        process['出货日期']=process['出货日期'].fillna(pd.Timestamp(1990, 1, 1))
        report_material.loc[report_material['出货日期1'] != pd.Timestamp(1990, 1, 1), '出货日期'] = report_material['出货日期1']
        report_material['出货日期']=report_material['出货日期'].fillna(pd.Timestamp(1990, 1, 1))
        reportwork.loc[reportwork['出货日期1'] != pd.Timestamp(1990, 1, 1), '出货日期'] = reportwork['出货日期1']
        reportwork['出货日期']=reportwork['出货日期'].fillna(pd.Timestamp(1990, 1, 1))
      #  reportcost.loc[reportcost['出货日期1'] != pd.Timestamp(1990, 1, 1), '出货日期'] = reportcost['出货日期1']
        del process['出货日期1']
        del report_material['出货日期1']
        del reportwork['出货日期1']
        ####当工单完工时间为空时等于实际出货时间
        process.loc[(process['工单完工时间'] == pd.Timestamp(1990, 1, 1))&(process['出货日期']!=pd.Timestamp(1990, 1, 1)), '工单完工时间'] = process['出货日期']
        report_material.loc[(report_material['工单完工时间'] == pd.Timestamp(1990, 1, 1))&(report_material['出货日期'] != pd.Timestamp(1990, 1, 1)), '工单完工时间'] = report_material['出货日期']
        reportwork.loc[(reportwork['工单完工时间'] == pd.Timestamp(1990, 1, 1))&(reportwork['出货日期'] != pd.Timestamp(1990, 1, 1)), '工单完工时间'] = reportwork['出货日期']
        reportcost.loc[(reportcost['出货日期'] != pd.Timestamp(1990, 1, 1))&(reportcost['工单完工时间'] == pd.Timestamp(1990, 1, 1)), '工单完工时间'] = reportcost['出货日期']
        del process['出货日期']
        del report_material['出货日期']
        del reportwork['出货日期']
        del reportcost['出货日期']

        time_assitant = time.time()
        screm.insert(INSERT, '\n第三阶段执行时长:%d秒' % (time_assitant - time_old), '\n')
#################################################################################################################################################################################
        screm.insert(INSERT, '\n四、第四阶段--四个明细表判断阶段', '\n')
        window.update()
#######################################
        screm.insert(INSERT, '\n4.1：采购po阶段', '\n')
        window.update()
        process.loc[process['采购日期'] == '', '采购日期'] = pd.Timestamp(1990, 1, 1)
        process['采购日期'] = pd.to_datetime(process['采购日期'], errors='coerce')
        process[['工单开立时间', '工单完工时间', '验收日期']] = process[['工单开立时间', '工单完工时间', '验收日期']].fillna(pd.Timestamp(1990, 1, 1))
        process['工单开立时间'] = pd.to_datetime(process['工单开立时间'], errors='coerce')
        process['工单完工时间'] = pd.to_datetime(process['工单完工时间'], errors='coerce')
        process['验收日期'] = pd.to_datetime(process['验收日期'], errors='coerce')
        process['核算阶段']=''
        process['核算项目号']=process['核算项目号'].fillna('')
        process.loc[(process['采购日期'] < process['工单开立时间']) | (process['工单开立时间'] == pd.Timestamp(1990, 1, 1)), '核算阶段'] = 'M2-M3阶段'
        process.loc[(process['采购日期'] >= process['工单开立时间']) & (process['采购日期'] <= process['工单完工时间']) & (
                process['工单完工时间'] != pd.Timestamp(1990, 1, 1)), '核算阶段'] = 'M4阶段'
        process.loc[(process['采购日期'] > process['工单完工时间']) & (process['工单完工时间'] != pd.Timestamp(1990, 1, 1)) & (process['采购日期'] <= process['验收日期']), '核算阶段'] = 'M5阶段'
        process.loc[(process['采购日期'] > process['工单完工时间']) & (process['验收日期'] == pd.Timestamp(1990, 1, 1))&(process['工单完工时间'] != pd.Timestamp(1990, 1, 1))& (process['工单完工时间'] != pd.Timestamp(2090, 1, 1)), '核算阶段'] = 'M5阶段'
        ########采购PO采购日期为空给M4阶段---出自4.6王珊
        process.loc[(process['采购日期'] == pd.Timestamp(1990, 1, 1)), '核算阶段'] = 'M4阶段'
        ###先判断M6阶段，之后的空给到M2-M3
        process.loc[(process['核算项目号'].str.contains('LEW')), '核算阶段'] = 'M5阶段'
        process.loc[(process['采购日期'] > process['验收日期'])&(process['验收日期']!=pd.Timestamp(1990, 1, 1)), '核算阶段'] = 'M6阶段'
        process.loc[(process['核算项目号'].str.contains('LEW')), '核算阶段'] = 'M5阶段'
        process['阶段补充']=''
        process.loc[(process['核算阶段'] ==''), '阶段补充'] = '无法判断'
        process.loc[(process['阶段补充'] == '无法判断'), '核算阶段'] = 'M2-M3阶段'
        process['核算阶段'] = process['核算阶段'].fillna('')
####################################
        screm.insert(INSERT, '\n4.2：料阶段', '\n')
        window.update()
        report_material['核算阶段'] = ''
        report_material.loc[report_material['扣账日期'] == '', '扣账日期'] = pd.Timestamp(1990, 1, 1)
        report_material['核算阶段'] = 'M4阶段'
        report_material['扣账日期'] = pd.to_datetime(report_material['扣账日期'], errors='coerce')
        report_material[['工单开立时间', '工单完工时间', '验收日期']] = report_material[['工单开立时间', '工单完工时间', '验收日期']].fillna(pd.Timestamp(1990, 1, 1))
        report_material['工单开立时间'] = pd.to_datetime(report_material['工单开立时间'], errors='coerce')
        report_material['工单完工时间'] = pd.to_datetime(report_material['工单完工时间'], errors='coerce')
        report_material['验收日期'] = pd.to_datetime(report_material['验收日期'], errors='coerce')
        report_material.loc[(report_material['扣账日期'] > report_material['工单完工时间']) &(report_material['扣账日期'] <= report_material['验收日期']), '核算阶段'] = 'M5阶段'
        report_material.loc[(report_material['扣账日期'] > report_material['工单完工时间']) & (report_material['验收日期'] == pd.Timestamp(1990, 1, 1))&(report_material['工单完工时间'] != pd.Timestamp(1990, 1, 1))& (report_material['工单完工时间'] != pd.Timestamp(2090, 1, 1)), '核算阶段'] = 'M5阶段'
        ########料单据日期为空给M4阶段---出自4.6王珊
        report_material.loc[(report_material['扣账日期'] == pd.Timestamp(1990, 1, 1)), '核算阶段'] = 'M4阶段'
        report_material['阶段补充'] = ''
        report_material['核算项目号'] = report_material['核算项目号'].fillna('')
        report_material.loc[(report_material['核算阶段'] == ''), '阶段补充'] = '无法判断'
        report_material.loc[(report_material['阶段补充'] == '无法判断'), '核算阶段'] = 'M4阶段'
        report_material['核算阶段'] = report_material['核算阶段'].fillna('')
        report_material.loc[(report_material['核算阶段'].str.contains('M4'))&(report_material['工单号码'].str.contains('603-')),'核算阶段']='M5阶段'
        report_material.loc[(report_material['核算项目号'].str.contains('LEW')), '核算阶段'] = 'M5阶段'
        report_material.loc[(report_material['扣账日期'] > report_material['验收日期']) & (report_material['验收日期'] != pd.Timestamp(1990, 1, 1)), '核算阶段'] = 'M6阶段'
#########################################
        screm.insert(INSERT, '\n4.3：工阶段', '\n')
        window.update()
        reportwork['核算阶段']=''
        reportwork['报工工单号'] = reportwork['报工工单号'].fillna('')
        reportwork['项目号整理'] = reportwork['项目号整理'].fillna('')
        reportwork['工种再分类'] = reportwork['工种再分类'].fillna('')
        reportwork['工种']=reportwork['工种'].fillna('')
        reportwork.loc[reportwork['完成日期'] == '', '完成日期'] = pd.Timestamp(1990, 1, 1)
        reportwork['完成日期'] = pd.to_datetime(reportwork['完成日期'], errors='coerce')
       # reportwork.loc[(reportwork['报工工单号'] == '') & (reportwork['工种'].str.contains('设计') == False), '核算阶段'] = 'M2-M3阶段'
        reportwork[['工单开立时间', '工单完工时间', '验收日期']] = reportwork[['工单开立时间', '工单完工时间', '验收日期']].fillna(pd.Timestamp(1990, 1, 1))
        reportwork['工单开立时间'] = pd.to_datetime(reportwork['工单开立时间'], errors='coerce')
        reportwork['工单完工时间'] = pd.to_datetime(reportwork['工单完工时间'], errors='coerce')
        reportwork['验收日期'] = pd.to_datetime(reportwork['验收日期'], errors='coerce')
        reportwork.loc[(reportwork['完成日期'] < reportwork['工单开立时间']) | (reportwork['工单开立时间'] == pd.Timestamp(1990, 1, 1)), '核算阶段'] = 'M2-M3阶段'
        reportwork.loc[(reportwork['核算阶段'].str.contains('M2'))&(reportwork['工种'].str.contains('生产')), '核算阶段'] = 'M4阶段'
        reportwork.loc[(reportwork['核算阶段'].str.contains('M2')) & (reportwork['工种'].str.contains('交付')), '核算阶段'] = 'M5阶段'
        reportwork.loc[(reportwork['完成日期'] >= reportwork['工单开立时间']) & (reportwork['完成日期'] <= reportwork['工单完工时间']) & (reportwork['工单完工时间'] != pd.Timestamp(1990, 1, 1)), '核算阶段'] = 'M4阶段'
        reportwork.loc[(reportwork['完成日期'] > reportwork['工单完工时间']) & (reportwork['完成日期'] <= reportwork['验收日期']), '核算阶段'] = 'M5阶段'
        reportwork.loc[(reportwork['完成日期'] > reportwork['工单完工时间']) & (reportwork['验收日期'] == pd.Timestamp(1990, 1, 1))&(reportwork['工单完工时间'] != pd.Timestamp(1990, 1, 1))&(reportwork['工单完工时间']!= pd.Timestamp(2090, 1, 1)), '核算阶段'] = 'M5阶段'
        reportwork.loc[(reportwork['完成日期'] > reportwork['验收日期']) & (reportwork['验收日期'] != pd.Timestamp(1990, 1, 1)), '核算阶段'] = 'M5阶段'
        reportwork.loc[(reportwork['完成日期'] > reportwork['验收日期']) & (reportwork['核算阶段'].str.contains('M5'))&((reportwork['项目号整理'].str.contains('-SH'))|(reportwork['报工工单号'].str.contains('609-'))), '核算阶段'] = 'M6阶段'
        reportwork.loc[(reportwork['项目号整理'].str.contains('LEW')),'核算阶段'] = 'M5阶段'
        reportwork.loc[(reportwork['项目号整理'].str.contains('LEK')), '核算阶段'] = 'M1阶段'
       # reportwork.loc[(reportwork['项目号整理'].str.contains('-')), '核算阶段'] = 'M5阶段'
        reportwork.loc[(reportwork['项目号整理'].str.contains('-SH')), '核算阶段'] = 'M6阶段'
        reportwork.loc[(reportwork['报工工单号'].str.contains('609-')), '核算阶段'] = 'M6阶段'
        reportwork.loc[(reportwork['报工工单号'].str.contains('603-|607-')), '核算阶段'] = 'M5阶段'
        ###工种修正
        reportwork.loc[(reportwork['核算阶段'].str.contains('M4')) , '工种'] = reportwork['工种'].str.replace('交付', '生产', regex=True)
        reportwork.loc[(reportwork['核算阶段'].str.contains('M4')) & (reportwork['报工工单号'].str.contains('601-')), '工种再分类'] = reportwork['工种再分类'].str.replace('交付', '生产', regex=True)
        #reportwork.loc[(reportwork['核算阶段'].str.contains('M5')), '工种'] = reportwork['工种'].str.replace('生产', '交付', regex=True).astype(str)
        #reportwork.loc[(reportwork['核算阶段'].str.contains('M5')), '工种再分类'] = reportwork['工种再分类'].str.replace('生产', '交付',regex=True).astype(str)
        reportwork['阶段补充'] = ''
        reportwork.loc[(reportwork['核算阶段'] == ''), '阶段补充'] = '无法判断'
        reportwork.loc[(reportwork['阶段补充'] == '无法判断'), '核算阶段'] = 'M4阶段'
        reportwork['核算阶段'] = reportwork['核算阶段'].fillna('')
#############################################
        screm.insert(INSERT, '\n4.4：费阶段', '\n')
        window.update()
        reportcost['核算阶段']=''
        # reportcost.loc[reportcost['申请时间'] == '', '申请时间'] = pd.Timestamp(1990, 1, 1)
        reportcost['申请时间'] = pd.to_datetime(reportcost['申请时间'], errors='coerce')
        reportcost[['工单开立时间', '工单完工时间', '验收日期']] = reportcost[['工单开立时间', '工单完工时间', '验收日期']].fillna(pd.Timestamp(1990, 1, 1))
        reportcost['工单开立时间'] = pd.to_datetime(reportcost['工单开立时间'], errors='coerce')
        reportcost['工单完工时间'] = pd.to_datetime(reportcost['工单完工时间'], errors='coerce')
        reportcost['验收日期'] = pd.to_datetime(reportcost['验收日期'], errors='coerce')
        reportcost.loc[(reportcost['申请时间'] < reportcost['工单开立时间']) | (reportcost['工单开立时间'] == pd.Timestamp(1990, 1, 1)), '核算阶段'] = 'M2-M3阶段'
        reportcost.loc[(reportcost['申请时间'] >= reportcost['工单开立时间']) & (reportcost['申请时间'] <= reportcost['工单完工时间']) & (reportcost['工单完工时间'] != pd.Timestamp(1990, 1, 1)), '核算阶段'] = 'M4阶段'
        reportcost.loc[ (reportcost['申请时间'] > reportcost['工单完工时间']) & (reportcost['申请时间'] <= reportcost['验收日期']), '核算阶段'] = 'M5阶段'
        reportcost.loc[(reportcost['申请时间'] > reportcost['工单完工时间']) & (reportcost['验收日期'] == pd.Timestamp(1990, 1, 1))&(reportcost['工单完工时间'] != pd.Timestamp(1990, 1, 1))&(reportcost['工单完工时间']!= pd.Timestamp(2090, 1, 1)), '核算阶段'] = 'M5阶段'
        reportcost['核算项目号']=reportcost['核算项目号'].fillna('')
        reportcost.loc[(reportcost['核算项目号'].str.contains('LEW')), '核算阶段'] = 'M5阶段'
        reportcost.loc[(reportcost['申请时间'] > reportcost['验收日期']) & (reportcost['验收日期'] != pd.Timestamp(1990, 1, 1)), '核算阶段'] = 'M6阶段'
        reportcost['阶段补充'] = ''
        reportcost.loc[(reportcost['核算阶段'] == ''), '阶段补充'] = '无法判断'
        reportcost.loc[(reportcost['阶段补充'] == '无法判断'), '核算阶段'] = 'M4阶段'
        reportcost['核算阶段'] = reportcost['核算阶段'].fillna('')
        ####工在M5阶段要加个出差补贴
        reportwork['出差补贴']=0
        reportwork.loc[reportwork['核算阶段'].str.contains('M5'),'出差补贴']=10*reportwork['工时']
        reportwork.loc[reportwork['核算阶段'].str.contains('M5'), '出差补贴'] = 10 * reportwork['工时']
        ###追加设定，海目星、生产工
        reportwork['工作地点']=reportwork['工作地点'].fillna('')
        reportwork.loc[reportwork['工作地点'].str.contains('海目星'), '出差补贴'] = 0
        reportwork.loc[reportwork['工种'].str.contains('生产'), '出差补贴'] = 0
        reportwork.loc[reportwork['完成日期'] >= pd.Timestamp(2023,5,1) , '出差补贴'] = 0

        time_stage = time.time()
        screm.insert(INSERT, '\n第四阶段执行时长:%d秒' % (time_stage-time_assitant ), '\n')

#################################################################################################################################################################################
        screm.insert(INSERT, '\n五、第五阶段-明细表金额数据汇入底表', '\n')
        window.update()
#######################################
        screm.insert(INSERT, '\n5.1：四个明细表金额全部汇入底表', '\n')
        window.update()
        report_material['领料类型']=report_material['领料类型'].fillna('')
        reportwork[['核算项目号','报工工单号']]=reportwork[['核算项目号','报工工单号']].fillna('')
        report_material[['核算项目号','工单号码']] = report_material[['核算项目号','工单号码']].fillna('')
        mater_M2_3 = report_material[report_material['核算阶段'].str.contains('M2')].reset_index(drop=True)
        order_mater_M2_3 = report_material[(report_material['核算阶段'].str.contains('M2'))&(report_material['领料类型'].str.contains('杂') == False)].reset_index(drop=True)
        changer_mater_M2_3 = report_material[(report_material['核算阶段'].str.contains('M2'))&(report_material['领料类型'].str.contains('杂'))].reset_index(drop=True)
        work_M2_3 = reportwork[reportwork['核算阶段'].str.contains('M2')].reset_index(drop=True)
        cost_M2_3 = reportcost[reportcost['核算阶段'].str.contains('M2')].reset_index(drop=True)
        po_M2_3 = process[process['核算阶段'].str.contains('M2')].reset_index(drop=True)
        prod_work_M2_3 = reportwork[(reportwork['核算阶段'].str.contains('M2')) & (reportwork['工种'].str.contains('生产'))].reset_index(drop=True)
        design_work_M2_3 = reportwork[(reportwork['核算阶段'].str.contains('M2')) & (reportwork['工种'].str.contains('设计'))].reset_index(drop=True)
        deli_work_M2_3 = reportwork[(reportwork['核算阶段'].str.contains('M2')) & (reportwork['工种'].str.contains('交付'))].reset_index(drop=True)

        mater_M4 = report_material[report_material['核算阶段'].str.contains('M4')].reset_index(drop=True)
        order_mater_M4= report_material[(report_material['核算阶段'].str.contains('M4'))&(report_material['领料类型'].str.contains('杂')== False)].reset_index(drop=True)
        changer_mater_M4 = report_material[(report_material['核算阶段'].str.contains('M4'))&(report_material['领料类型'].str.contains('杂'))].reset_index(drop=True)
        work_M4 = reportwork[reportwork['核算阶段'].str.contains('M4')].reset_index(drop=True)
        cost_M4 = reportcost[reportcost['核算阶段'].str.contains('M4')].reset_index(drop=True)
        po_M4 = process[process['核算阶段'].str.contains('M4')].reset_index(drop=True)
        prod_work_M4 = reportwork[ (reportwork['核算阶段'].str.contains('M4')) & (reportwork['工种'].str.contains('生产'))].reset_index(drop=True)
        design_work_M4 = reportwork[(reportwork['核算阶段'].str.contains('M4')) & (reportwork['工种'].str.contains('设计'))].reset_index(drop=True)
        deli_work_M4 = reportwork[(reportwork['核算阶段'].str.contains('M4')) & (reportwork['工种'].str.contains('交付'))].reset_index(drop=True)

        mater_M5 = report_material[report_material['核算阶段'].str.contains('M5')].reset_index(drop=True)
        order_mater_M5 = report_material[(report_material['核算阶段'].str.contains('M5')) & (report_material['领料类型'].str.contains('杂') == False)].reset_index(drop=True)
        changer_mater_M5 = report_material[(report_material['核算阶段'].str.contains('M5')) & (report_material['领料类型'].str.contains('杂'))].reset_index(drop=True)
        work_M5 = reportwork[reportwork['核算阶段'].str.contains('M5')].reset_index(drop=True)
        cost_M5 = reportcost[reportcost['核算阶段'].str.contains('M5')].reset_index(drop=True)
        po_M5 = process[process['核算阶段'].str.contains('M5')].reset_index(drop=True)
        prod_work_M5 = reportwork[(reportwork['核算阶段'].str.contains('M5')) & (reportwork['工种'].str.contains('生产'))].reset_index(drop=True)
        design_work_M5 = reportwork[(reportwork['核算阶段'].str.contains('M5')) & (reportwork['工种'].str.contains('设计'))].reset_index(drop=True)
        deli_work_M5 = reportwork[(reportwork['核算阶段'].str.contains('M5')) & (reportwork['工种'].str.contains('交付'))].reset_index(drop=True)
        ##空值置0
        mater_M2_3['核算项目号'] = mater_M2_3['核算项目号'].fillna('')
        mater_M2_3['未税金额'] = mater_M2_3['未税金额'].fillna(0)
        mater_M2_3.loc[mater_M2_3['未税金额'] == '', '未税金额'] = 0
        mater_M4['核算项目号'] = mater_M4['核算项目号'].fillna('')
        mater_M4['未税金额'] = mater_M4['未税金额'].fillna(0)
        mater_M4.loc[mater_M4['未税金额'] == '', '未税金额'] = 0
        mater_M5['核算项目号'] = mater_M5['核算项目号'].fillna('')
        mater_M5['未税金额'] = mater_M5['未税金额'].fillna(0)
        mater_M5.loc[mater_M5['未税金额'] == '', '未税金额'] = 0

        order_mater_M2_3['核算项目号'] = order_mater_M2_3['核算项目号'].fillna('')
        order_mater_M2_3['未税金额'] = order_mater_M2_3['未税金额'].fillna(0)
        order_mater_M2_3.loc[order_mater_M2_3['未税金额'] == '', '未税金额'] = 0
        order_mater_M4['核算项目号'] = order_mater_M4['核算项目号'].fillna('')
        order_mater_M4['未税金额'] = order_mater_M4['未税金额'].fillna(0)
        order_mater_M4.loc[order_mater_M4['未税金额'] == '', '未税金额'] = 0
        order_mater_M5['核算项目号'] = order_mater_M5['核算项目号'].fillna('')
        order_mater_M5['未税金额'] = order_mater_M5['未税金额'].fillna(0)
        order_mater_M5.loc[order_mater_M5['未税金额'] == '', '未税金额'] = 0

        changer_mater_M2_3['核算项目号'] = changer_mater_M2_3['核算项目号'].fillna('')
        changer_mater_M2_3['未税金额'] = changer_mater_M2_3['未税金额'].fillna(0)
        changer_mater_M2_3.loc[changer_mater_M2_3['未税金额'] == '', '未税金额'] = 0
        changer_mater_M4['核算项目号'] = changer_mater_M4['核算项目号'].fillna('')
        changer_mater_M4['未税金额'] = changer_mater_M4['未税金额'].fillna(0)
        changer_mater_M4.loc[changer_mater_M4['未税金额'] == '', '未税金额'] = 0
        changer_mater_M5['核算项目号'] = changer_mater_M5['核算项目号'].fillna('')
        changer_mater_M5['未税金额'] = changer_mater_M5['未税金额'].fillna(0)
        changer_mater_M5.loc[changer_mater_M5['未税金额'] == '', '未税金额'] = 0

        cost_M2_3['核算项目号'] = cost_M2_3['核算项目号'].fillna('')
        cost_M2_3['金额万元'] = cost_M2_3['金额万元'].fillna(0)
        cost_M2_3.loc[cost_M2_3['金额万元'] == '', '金额万元'] = 0
        cost_M4['核算项目号'] = cost_M4['核算项目号'].fillna('')
        cost_M4['金额万元'] = cost_M4['金额万元'].fillna(0)
        cost_M4.loc[cost_M4['金额万元'] == '', '金额万元'] = 0
        cost_M5['核算项目号'] = cost_M5['核算项目号'].fillna('')
        cost_M5['金额万元'] = cost_M5['金额万元'].fillna(0)
        cost_M5.loc[cost_M5['金额万元'] == '', '金额万元'] = 0

        po_M2_3['核算项目号'] = po_M2_3['核算项目号'].fillna('')
        po_M2_3['采购金额-未税'] = po_M2_3['采购金额-未税'].fillna(0)
        po_M2_3.loc[po_M2_3['采购金额-未税'] == '', '采购金额-未税'] = 0
        po_M4['核算项目号'] = po_M4['核算项目号'].fillna('')
        po_M4['采购金额-未税'] = po_M4['采购金额-未税'].fillna(0)
        po_M4.loc[po_M4['采购金额-未税'] == '', '采购金额-未税'] = 0
        po_M5['核算项目号'] = po_M5['核算项目号'].fillna('')
        po_M5['采购金额-未税'] = po_M5['采购金额-未税'].fillna(0)
        po_M5.loc[po_M5['采购金额-未税'] == '', '采购金额-未税'] = 0

        prod_work_M2_3['核算项目号'] = prod_work_M2_3['核算项目号'].fillna('')
        prod_work_M2_3['工时成本'] = prod_work_M2_3['工时成本'].fillna(0)
        prod_work_M2_3.loc[prod_work_M2_3['工时成本'] == '', '工时成本'] = 0
        prod_work_M4['核算项目号'] = prod_work_M4['核算项目号'].fillna('')
        prod_work_M4['工时成本'] = prod_work_M4['工时成本'].fillna(0)
        prod_work_M4.loc[prod_work_M4['工时成本'] == '', '工时成本'] = 0
        prod_work_M5['核算项目号'] = prod_work_M5['核算项目号'].fillna('')
        prod_work_M5['工时成本'] = prod_work_M5['工时成本'].fillna(0)
        prod_work_M5.loc[prod_work_M5['工时成本'] == '', '工时成本'] = 0

        design_work_M2_3['核算项目号'] = design_work_M2_3['核算项目号'].fillna('')
        design_work_M2_3['工时成本'] = design_work_M2_3['工时成本'].fillna(0)
        design_work_M2_3.loc[design_work_M2_3['工时成本'] == '', '工时成本'] = 0
        design_work_M4['核算项目号'] = design_work_M4['核算项目号'].fillna('')
        design_work_M4['工时成本'] = design_work_M4['工时成本'].fillna(0)
        design_work_M4.loc[design_work_M4['工时成本'] == '', '工时成本'] = 0
        design_work_M5['核算项目号'] = design_work_M5['核算项目号'].fillna('')
        design_work_M5['工时成本'] = design_work_M5['工时成本'].fillna(0)
        design_work_M5.loc[design_work_M5['工时成本'] == '', '工时成本'] = 0

        deli_work_M2_3['核算项目号'] = deli_work_M2_3['核算项目号'].fillna('')
        deli_work_M2_3['工时成本'] = deli_work_M2_3['工时成本'].fillna(0)
        deli_work_M2_3.loc[deli_work_M2_3['工时成本'] == '', '工时成本'] = 0
        deli_work_M4['核算项目号'] = deli_work_M4['核算项目号'].fillna('')
        deli_work_M4['工时成本'] = deli_work_M4['工时成本'].fillna(0)
        deli_work_M4.loc[deli_work_M4['工时成本'] == '', '工时成本'] = 0
        deli_work_M5['核算项目号'] = deli_work_M5['核算项目号'].fillna('')
        deli_work_M5['工时成本'] = deli_work_M5['工时成本'].fillna(0)
        deli_work_M5.loc[deli_work_M5['工时成本'] == '', '工时成本'] = 0
        report_item_cal_out = report_item_cal_out.reindex(columns=["序列号", "区域", "行业中心", "设备类型", "客户简称", "大项目名称", "大项目号", "产品线名称", "核算项目号", "设备名称", "项目数量"
            , "已出货数量", "在产数量", "生产状态", "集团收入", "软件收入", "硬件收入"
            , 'M2-3成本合计', 'M2-3料','M2-3工单料','M2-3设变料', 'M2-3采购PO', 'M2-3工', 'M2-3生产工','M2-3交付工', 'M2-3费', 'M2-3设计工', 'M2-3其他费'
            , "一般工单号601/608", "工单开立时间", "工单完工时间", 'M4成本合计', 'M4料','M4工单料','M4设变料', 'M4采购PO', 'M4工','M4生产工', 'M4交付工', 'M4费', 'M4设计工', 'M4其他费', "系统出货时间", "实际出货时间"
            , "返工工单号603", 'M5成本合计', 'M5料','M5工单料','M5设变料', 'M5采购PO', 'M5工','M5生产工', 'M5交付工', 'M5费', 'M5设计工', 'M5其他费', "系统验收时间", "实际验收时间"
            , "项目号整理", "成品料号", "是否预验收", "全面预算有无",'OA状态', "项目财经", "项目财经再分类"])
########################按工单号计算汇总表金额
        screm.insert(INSERT, '\n5.2：按工单号及数量均摊金额', '\n')
        window.update()
        mater_M2_3_value = mater_M2_3.groupby(['核算项目号'])[['未税金额']].sum().add_suffix('-之和').reset_index()
        ##work_M2_3_value=work_M2_3.groupby(['核算项目号'])[['工时成本']].sum().add_suffix('-之和').reset_index()
        order_mater_M2_3_value = order_mater_M2_3.groupby(['核算项目号'])[['未税金额']].sum().add_suffix('-之和').reset_index()
        changer_mater_M2_3_value =changer_mater_M2_3.groupby(['核算项目号'])[['未税金额']].sum().add_suffix('-之和').reset_index()
        cost_M2_3_value = cost_M2_3.groupby(['核算项目号'])[['金额万元']].sum().add_suffix('-之和').reset_index()
        po_M2_3_value = po_M2_3.groupby(['核算项目号'])[['采购金额-未税']].sum().add_suffix('-之和').reset_index()
        prod_work_M2_3_value = prod_work_M2_3.groupby(['核算项目号'])[['工时成本']].sum().add_suffix('-之和').reset_index()
        design_work_M2_3_value = design_work_M2_3.groupby(['核算项目号'])[['工时成本']].sum().add_suffix('-之和').reset_index()
        deli_work_M2_3_value = deli_work_M2_3.groupby(['核算项目号'])[['工时成本']].sum().add_suffix('-之和').reset_index()
        ###M4
        ##料带汇总表工单
        mater_M4['工单号码']=mater_M4['工单号码'].fillna('').astype(str)
        mater_M4['项目号_工单']='无工单'
        mater_M4.loc[(mater_M4['工单号码'].isin(report_item_cal_out['一般工单号601/608']))&(mater_M4['工单号码']!=''),'项目号_工单']=mater_M4['核算项目号']+'_'+mater_M4['工单号码']
        mater_M4_value = mater_M4.groupby(['项目号_工单'])[['未税金额']].sum().add_suffix('-之和').reset_index()
        mater_M4_value1 = mater_M4.groupby(['核算项目号'])[['未税金额']].sum().add_suffix('-之和').reset_index()

        order_mater_M4['工单号码'] = order_mater_M4['工单号码'].fillna('').astype(str)
        order_mater_M4['项目号_工单'] = '无工单'
        order_mater_M4.loc[(order_mater_M4['工单号码'].isin(report_item_cal_out['一般工单号601/608'])) & (order_mater_M4['工单号码'] != ''), '项目号_工单'] = order_mater_M4['核算项目号'] + '_' + order_mater_M4['工单号码']
        order_mater_M4_value = order_mater_M4.groupby(['项目号_工单'])[['未税金额']].sum().add_suffix('-之和').reset_index()
        order_mater_M4_value1 = order_mater_M4.groupby(['核算项目号'])[['未税金额']].sum().add_suffix('-之和').reset_index()

        changer_mater_M4['工单号码'] = changer_mater_M4['工单号码'].fillna('').astype(str)
        changer_mater_M4['项目号_工单'] = '无工单'
        changer_mater_M4.loc[(changer_mater_M4['工单号码'].isin(report_item_cal_out['一般工单号601/608'])) & ( changer_mater_M4['工单号码'] != ''), '项目号_工单'] = changer_mater_M4['核算项目号'] + '_' + changer_mater_M4['工单号码']
        changer_mater_M4_value = changer_mater_M4.groupby(['项目号_工单'])[['未税金额']].sum().add_suffix('-之和').reset_index()
        changer_mater_M4_value1 = changer_mater_M4.groupby(['核算项目号'])[['未税金额']].sum().add_suffix('-之和').reset_index()

        ##work_M4_value=work_M4.groupby(['核算项目号'])[['工时成本']].sum().add_suffix('-之和').reset_index()
        cost_M4_value = cost_M4.groupby(['核算项目号'])[['金额万元']].sum().add_suffix('-之和').reset_index()
        po_M4_value = po_M4.groupby(['核算项目号'])[['采购金额-未税']].sum().add_suffix('-之和').reset_index()

        prod_work_M4['报工工单号'] = prod_work_M4['报工工单号'].fillna('')
        prod_work_M4['项目号_工单'] = '无工单'
        prod_work_M4.loc[(prod_work_M4['报工工单号'].isin(report_item_cal_out['一般工单号601/608'])) & (prod_work_M4['报工工单号'] != ''), '项目号_工单'] = prod_work_M4['核算项目号'] + '_' + prod_work_M4['报工工单号']
        prod_work_M4_value = prod_work_M4.groupby(['项目号_工单'])[['工时成本']].sum().add_suffix('-之和').reset_index()
        prod_work_M4_value1 = prod_work_M4.groupby(['核算项目号'])[['工时成本']].sum().add_suffix('-之和').reset_index()

        design_work_M4['报工工单号'] = design_work_M4['报工工单号'].fillna('')
        design_work_M4['项目号_工单'] = '无工单'
        design_work_M4.loc[(design_work_M4['报工工单号'].isin(report_item_cal_out['一般工单号601/608'])) & (design_work_M4['报工工单号'] != ''), '项目号_工单'] = design_work_M4['核算项目号'] + '_' + design_work_M4['报工工单号']
        design_work_M4_value = design_work_M4.groupby(['项目号_工单'])[['工时成本']].sum().add_suffix('-之和').reset_index()
        design_work_M4_value1 = design_work_M4.groupby(['核算项目号'])[['工时成本']].sum().add_suffix('-之和').reset_index()

        deli_work_M4['报工工单号'] = deli_work_M4['报工工单号'].fillna('')
        deli_work_M4['项目号_工单'] = '无工单'
        deli_work_M4.loc[(deli_work_M4['报工工单号'].isin(report_item_cal_out['一般工单号601/608'])) & (deli_work_M4['报工工单号'] != ''), '项目号_工单'] = deli_work_M4['核算项目号'] + '_' + deli_work_M4['报工工单号']
        deli_work_M4_value = deli_work_M4.groupby(['项目号_工单'])[['工时成本']].sum().add_suffix('-之和').reset_index()
        deli_work_M4_value1 = deli_work_M4.groupby(['核算项目号'])[['工时成本']].sum().add_suffix('-之和').reset_index()

        ###M5
        mater_M5['工单号码'] = mater_M5['工单号码'].fillna('')
        mater_M5['项目号_工单'] = '无工单'
        mater_M5.loc[(mater_M5['工单号码'].isin(report_item_cal_out['返工工单号603'])) & (mater_M5['工单号码'] != ''), '项目号_工单'] = mater_M5['核算项目号'] + '_' + mater_M5['工单号码']
        mater_M5_value = mater_M5.groupby(['项目号_工单'])[['未税金额']].sum().add_suffix('-之和').reset_index()
        mater_M5_value1 = mater_M5.groupby(['核算项目号'])[['未税金额']].sum().add_suffix('-之和').reset_index()
        ##work_M5_value=work_M5.groupby(['核算项目号'])[['工时成本']].sum().add_suffix('-之和').reset_index()

        order_mater_M5['工单号码'] = order_mater_M5['工单号码'].fillna('')
        order_mater_M5['项目号_工单'] = '无工单'
        order_mater_M5.loc[(order_mater_M5['工单号码'].isin(report_item_cal_out['返工工单号603'])) & (order_mater_M5['工单号码'] != ''), '项目号_工单'] =order_mater_M5['核算项目号'] + '_' + order_mater_M5['工单号码']
        order_mater_M5_value = order_mater_M5.groupby(['项目号_工单'])[['未税金额']].sum().add_suffix('-之和').reset_index()
        order_mater_M5_value1 = order_mater_M5.groupby(['核算项目号'])[['未税金额']].sum().add_suffix('-之和').reset_index()

        changer_mater_M5['工单号码'] = changer_mater_M5['工单号码'].fillna('')
        changer_mater_M5['项目号_工单'] = '无工单'
        changer_mater_M5.loc[(changer_mater_M5['工单号码'].isin(report_item_cal_out['返工工单号603'])) & (changer_mater_M5['工单号码'] != ''), '项目号_工单'] = changer_mater_M5['核算项目号'] + '_' + changer_mater_M5['工单号码']
        changer_mater_M5_value = changer_mater_M5.groupby(['项目号_工单'])[['未税金额']].sum().add_suffix('-之和').reset_index()
        changer_mater_M5_value1 = changer_mater_M5.groupby(['核算项目号'])[['未税金额']].sum().add_suffix('-之和').reset_index()

        cost_M5_value = cost_M5.groupby(['核算项目号'])[['金额万元']].sum().add_suffix('-之和').reset_index()
        po_M5_value = po_M5.groupby(['核算项目号'])[['采购金额-未税']].sum().add_suffix('-之和').reset_index()

        prod_work_M5['报工工单号'] = prod_work_M5['报工工单号'].fillna('')
        prod_work_M5['项目号_工单'] = '无工单'
        prod_work_M5.loc[(prod_work_M5['报工工单号'].isin(report_item_cal_out['返工工单号603'])) & (prod_work_M5['报工工单号'] != ''), '项目号_工单'] = prod_work_M5['核算项目号'] + '_' + prod_work_M5['报工工单号']
        prod_work_M5_value = prod_work_M5.groupby(['项目号_工单'])[['工时成本']].sum().add_suffix('-之和').reset_index()
        prod_work_M5_value1 = prod_work_M5.groupby(['核算项目号'])[['工时成本']].sum().add_suffix('-之和').reset_index()

        design_work_M5['报工工单号'] = design_work_M5['报工工单号'].fillna('')
        design_work_M5['项目号_工单'] = '无工单'
        design_work_M5.loc[(design_work_M5['报工工单号'].isin(report_item_cal_out['返工工单号603'])) & (
                design_work_M5['报工工单号'] != ''), '项目号_工单'] = design_work_M5['核算项目号'] + '_' + design_work_M5['报工工单号']
        design_work_M5_value = design_work_M5.groupby(['项目号_工单'])[['工时成本']].sum().add_suffix('-之和').reset_index()
        design_work_M5_value1 = design_work_M5.groupby(['核算项目号'])[['工时成本']].sum().add_suffix('-之和').reset_index()

        deli_work_M5['报工工单号'] = deli_work_M5['报工工单号'].fillna('')
        deli_work_M5['项目号_工单'] = '无工单'
        deli_work_M5.loc[(deli_work_M5['报工工单号'].isin(report_item_cal_out['返工工单号603'])) & (deli_work_M5['报工工单号'] != ''), '项目号_工单'] = deli_work_M5['核算项目号'] + '_' + deli_work_M5['报工工单号']
        deli_work_M5_value = deli_work_M5.groupby(['项目号_工单'])[['工时成本']].sum().add_suffix('-之和').reset_index()
        deli_work_M5_value1 = deli_work_M5.groupby(['核算项目号'])[['工时成本']].sum().add_suffix('-之和').reset_index()
        reportwork['核算项目号'] = reportwork['核算项目号'].fillna('')
        travel_money = reportwork.groupby(['核算项目号'])[['出差补贴']].sum().add_suffix('').reset_index()
        ##M2-M3
        report_item_cal_out['M2-3料'] = 0
        report_item_cal_out['M2-3工单料'] = 0
        report_item_cal_out['M2-3设变料'] = 0
        report_item_cal_out['M2-3采购PO'] = 0
        report_item_cal_out['M2-3工'] = 0
        report_item_cal_out['M2-3生产工'] = 0
        report_item_cal_out['M2-3交付工'] = 0
        report_item_cal_out['M2-3费'] = 0
        report_item_cal_out['M2-3设计工'] = 0
        report_item_cal_out['M2-3其他费'] = 0
        ##M4
        report_item_cal_out['M4料'] = 0
        report_item_cal_out['M4工单料'] = 0
        report_item_cal_out['M4设变料'] = 0
        report_item_cal_out['M4采购PO'] = 0
        report_item_cal_out['M4工'] = 0
        report_item_cal_out['M4生产工'] = 0
        report_item_cal_out['M4交付工'] = 0
        report_item_cal_out['M4费'] = 0
        report_item_cal_out['M4设计工'] = 0
        report_item_cal_out['M4其他费'] = 0
        ##m5
        report_item_cal_out['M5料'] = 0
        report_item_cal_out['M5工单料'] = 0
        report_item_cal_out['M5设变料'] = 0
        report_item_cal_out['M5采购PO'] = 0
        report_item_cal_out['M5工'] = 0
        report_item_cal_out['M5生产工'] = 0
        report_item_cal_out['M5交付工'] = 0
        report_item_cal_out['M5费'] = 0
        report_item_cal_out['M5设计工'] = 0
        report_item_cal_out['M5其他费'] = 0

        report_item_cal_out['核算项目号'] = report_item_cal_out['核算项目号'].fillna('')
        if len(mater_M2_3) == 0:
            report_item_cal_out['M2-3料'] = 0
        else:
            report_item_cal_out['M2-3料'] = pd.merge(report_item_cal_out, mater_M2_3_value, on='核算项目号', how='left')['未税金额-之和']

        if len(order_mater_M2_3) == 0:
            report_item_cal_out['M2-3工单料'] = 0
        else:
            report_item_cal_out['M2-3工单料'] = pd.merge(report_item_cal_out, order_mater_M2_3_value, on='核算项目号', how='left')['未税金额-之和']

        if len(changer_mater_M2_3) == 0:
            report_item_cal_out['M2-3设变料'] = 0
        else:
            report_item_cal_out['M2-3设变料'] = pd.merge(report_item_cal_out, changer_mater_M2_3_value, on='核算项目号', how='left')['未税金额-之和']
        if len(po_M2_3) == 0:
            report_item_cal_out['M2-3采购PO'] = 0
        else:
            report_item_cal_out['M2-3采购PO'] = pd.merge(report_item_cal_out, po_M2_3_value, on='核算项目号', how='left')['采购金额-未税-之和']
        if len(prod_work_M2_3) == 0:
            report_item_cal_out['M2-3生产工'] = 0
        else:
            report_item_cal_out['M2-3生产工'] = \
            pd.merge(report_item_cal_out, prod_work_M2_3_value, on='核算项目号', how='left')['工时成本-之和']
        if len(deli_work_M2_3) == 0:
            report_item_cal_out['M2-3交付工'] = 0
        else:
            report_item_cal_out['M2-3交付工'] = \
            pd.merge(report_item_cal_out, deli_work_M2_3_value, on='核算项目号', how='left')['工时成本-之和']
        if len(design_work_M2_3) == 0:
            report_item_cal_out['M2-3设计工'] = 0
        else:
            report_item_cal_out['M2-3设计工'] = \
                pd.merge(report_item_cal_out, design_work_M2_3_value, on='核算项目号', how='left')['工时成本-之和']
        if len(cost_M2_3) == 0:
            report_item_cal_out['M2-3其他费'] = 0
        else:
            report_item_cal_out['M2-3其他费'] = pd.merge(report_item_cal_out, cost_M2_3_value, on='核算项目号', how='left')[
                '金额万元-之和']
        report_item_cal_out['项目号_工单1']='找不到工单'
        report_item_cal_out.loc[report_item_cal_out['一般工单号601/608']!="",'项目号_工单1']=report_item_cal_out['核算项目号']+'_' + report_item_cal_out['一般工单号601/608']
        if len(mater_M4) == 0:
            report_item_cal_out['M4料'] = 0
            report_item_cal_out['M4料2'] = 0
        else:
            report_item_cal_out['M4料'] = pd.merge(report_item_cal_out, mater_M4_value,left_on='项目号_工单1',right_on='项目号_工单', how='left')['未税金额-之和']
            report_item_cal_out['M4料1'] = pd.merge(report_item_cal_out, mater_M4_value1, on='核算项目号', how='left')['未税金额-之和']
        ###处理两个表
            total_mater_M4_1=report_item_cal_out.drop_duplicates(subset=['核算项目号','一般工单号601/608']).reset_index(drop=True)
            total_mater_M4_1=total_mater_M4_1.groupby(['核算项目号'])[['M4料']].sum().add_suffix('').reset_index()
            total_mater_M4_2 = report_item_cal_out.drop_duplicates(subset=['核算项目号']).reset_index(drop=True)[['核算项目号','M4料1']]
            total_mater_M4_2['M4料'] = pd.merge(total_mater_M4_2,total_mater_M4_1,on='核算项目号',how='left')['M4料']
            total_mater_M4_2['M4料2']=total_mater_M4_2['M4料1']-total_mater_M4_2['M4料']
            report_item_cal_out['M4料2']=pd.merge(report_item_cal_out, total_mater_M4_2, on='核算项目号', how='left')['M4料2']

        if len(order_mater_M4) == 0:
            report_item_cal_out['M4工单料'] = 0
            report_item_cal_out['M4工单料2'] = 0
        else:
            report_item_cal_out['M4工单料'] = pd.merge(report_item_cal_out, order_mater_M4_value, left_on='项目号_工单1', right_on='项目号_工单', how='left')['未税金额-之和']
            report_item_cal_out['M4工单料1'] = pd.merge(report_item_cal_out, order_mater_M4_value1, on='核算项目号', how='left')['未税金额-之和']
            ###处理两个表
            total_order_mater_M4_1 = report_item_cal_out.drop_duplicates(subset=['核算项目号', '一般工单号601/608']).reset_index(drop=True)
            total_order_mater_M4_1 = total_order_mater_M4_1.groupby(['核算项目号'])[['M4工单料']].sum().add_suffix('').reset_index()
            total_order_mater_M4_2 = report_item_cal_out.drop_duplicates(subset=['核算项目号']).reset_index(drop=True)[['核算项目号', 'M4工单料1']]
            total_order_mater_M4_2['M4工单料'] = pd.merge(total_order_mater_M4_2, total_order_mater_M4_1, on='核算项目号', how='left')['M4工单料']
            total_order_mater_M4_2['M4工单料2'] = total_order_mater_M4_2['M4工单料1'] - total_order_mater_M4_2['M4工单料']
            report_item_cal_out['M4工单料2'] =pd.merge(report_item_cal_out, total_order_mater_M4_2, on='核算项目号', how='left')['M4工单料2']

        if len(changer_mater_M4) == 0:
            report_item_cal_out['M4设变料'] = 0
            report_item_cal_out['M4设变料2'] = 0
        else:
            report_item_cal_out['M4设变料'] = pd.merge(report_item_cal_out, changer_mater_M4_value, left_on='项目号_工单1', right_on='项目号_工单', how='left')['未税金额-之和']
            report_item_cal_out['M4设变料1'] = pd.merge(report_item_cal_out, changer_mater_M4_value1, on='核算项目号', how='left')['未税金额-之和']
            ###处理两个表
            total_changer_mater_M4_1 = report_item_cal_out.drop_duplicates(subset=['核算项目号', '一般工单号601/608']).reset_index(drop=True)
            total_changer_mater_M4_1 = total_changer_mater_M4_1.groupby(['核算项目号'])[['M4设变料']].sum().add_suffix('').reset_index()
            total_changer_mater_M4_2 = report_item_cal_out.drop_duplicates(subset=['核算项目号']).reset_index(drop=True)[['核算项目号', 'M4设变料1']]
            total_changer_mater_M4_2['M4设变料'] = pd.merge(total_changer_mater_M4_2, total_changer_mater_M4_1, on='核算项目号',how='left')['M4设变料']
            total_changer_mater_M4_2['M4设变料2'] = total_changer_mater_M4_2['M4设变料1'] - total_changer_mater_M4_2['M4设变料']
            report_item_cal_out['M4设变料2'] =pd.merge(report_item_cal_out, total_changer_mater_M4_2, on='核算项目号', how='left')['M4设变料2']

        if len(po_M4) == 0:
            report_item_cal_out['M4采购PO'] = 0
        else:
            report_item_cal_out['M4采购PO'] = pd.merge(report_item_cal_out, po_M4_value, on='核算项目号', how='left')['采购金额-未税-之和']

        if len(prod_work_M4) == 0:
            report_item_cal_out['M4生产工'] = 0
            report_item_cal_out['M4生产工2'] = 0
        else:
            report_item_cal_out['M4生产工'] = pd.merge(report_item_cal_out, prod_work_M4_value, left_on='项目号_工单1',right_on='项目号_工单', how='left')['工时成本-之和']
            report_item_cal_out['M4生产工1'] = pd.merge(report_item_cal_out, prod_work_M4_value1, on='核算项目号',how='left')['工时成本-之和']
            total_prod_work_M4_1 = report_item_cal_out.drop_duplicates(subset=['核算项目号', '一般工单号601/608']).reset_index(drop=True)
            total_prod_work_M4_1 = total_prod_work_M4_1.groupby(['核算项目号'])[['M4生产工']].sum().add_suffix('').reset_index()
            total_prod_work_M4_2 = report_item_cal_out.drop_duplicates(subset=['核算项目号']).reset_index(drop=True)[['核算项目号', 'M4生产工1']]
            total_prod_work_M4_2['M4生产工'] = pd.merge(total_prod_work_M4_2, total_prod_work_M4_1, on='核算项目号', how='left')['M4生产工']
            total_prod_work_M4_2['M4生产工2'] = total_prod_work_M4_2['M4生产工1'] - total_prod_work_M4_2['M4生产工']
            report_item_cal_out['M4生产工2'] = pd.merge(report_item_cal_out, total_prod_work_M4_2, on='核算项目号', how='left')['M4生产工2']
        if len(deli_work_M4) == 0:
            report_item_cal_out['M4交付工'] = 0
            report_item_cal_out['M4交付工2'] = 0
        else:
            report_item_cal_out['M4交付工'] = pd.merge(report_item_cal_out, deli_work_M4_value,left_on='项目号_工单1',right_on='项目号_工单', how='left')['工时成本-之和']
            report_item_cal_out['M4交付工1'] =pd.merge(report_item_cal_out, deli_work_M4_value1, on='核算项目号', how='left')['工时成本-之和']
            total_deli_work_M4_1 = report_item_cal_out.drop_duplicates(subset=['核算项目号', '一般工单号601/608']).reset_index(rop=True)
            total_deli_work_M4_1 = total_deli_work_M4_1.groupby(['核算项目号'])[['M4交付工']].sum().add_suffix('').reset_index()
            total_deli_work_M4_2 = report_item_cal_out.drop_duplicates(subset=['核算项目号']).reset_index(drop=True)[['核算项目号', 'M4交付工1']]
            total_deli_work_M4_2['M4交付工'] =pd.merge(total_deli_work_M4_2, total_deli_work_M4_1, on='核算项目号', how='left')['M4交付工']
            total_deli_work_M4_2['M4交付工2'] = total_deli_work_M4_2['M4生产工1'] - total_deli_work_M4_2['M4交付工']
            report_item_cal_out['M4交付工2'] = pd.merge(report_item_cal_out, total_deli_work_M4_2, on='核算项目号', how='left')['M4交付工2']
        if len(design_work_M4) == 0:
            report_item_cal_out['M4设计工'] = 0
            report_item_cal_out['M4设计工2'] = 0
        else:
            report_item_cal_out['M4设计工'] = pd.merge(report_item_cal_out, design_work_M4_value, left_on='项目号_工单1',right_on='项目号_工单', how='left')['工时成本-之和']
            report_item_cal_out['M4设计工1'] = pd.merge(report_item_cal_out, design_work_M4_value1, on='核算项目号', how='left')['工时成本-之和']
            total_design_work_M4_1 = report_item_cal_out.drop_duplicates(subset=['核算项目号', '一般工单号601/608']).reset_index(drop=True)
            total_design_work_M4_1 = total_design_work_M4_1.groupby(['核算项目号'])[['M4设计工']].sum().add_suffix('').reset_index()
            total_design_work_M4_2 = report_item_cal_out.drop_duplicates(subset=['核算项目号']).reset_index(drop=True)[['核算项目号', 'M4设计工1']]
            total_design_work_M4_2['M4设计工'] =pd.merge(total_design_work_M4_2, total_design_work_M4_1, on='核算项目号', how='left')['M4设计工']
            total_design_work_M4_2['M4设计工2'] = total_design_work_M4_2['M4设计工1'] - total_design_work_M4_2['M4设计工']
            report_item_cal_out['M4设计工2'] = pd.merge(report_item_cal_out, total_design_work_M4_2, on='核算项目号', how='left')['M4设计工2']
        if len(cost_M4) == 0:
            report_item_cal_out['M4其他费'] = 0
        else:
            report_item_cal_out['M4其他费'] = pd.merge(report_item_cal_out, cost_M4_value, on='核算项目号', how='left')['金额万元-之和']

        report_item_cal_out['项目号_工单2'] = '找不到工单'
        report_item_cal_out.loc[report_item_cal_out['返工工单号603'] != "", '项目号_工单2'] = report_item_cal_out['核算项目号'] + '_' + report_item_cal_out['返工工单号603']
        if len(mater_M5) == 0:
            report_item_cal_out['M5料'] = 0
            report_item_cal_out['M5料2'] = 0
        else:
            report_item_cal_out['M5料'] = pd.merge(report_item_cal_out, mater_M5_value, left_on='项目号_工单2',right_on='项目号_工单', how='left')['未税金额-之和']
            report_item_cal_out['M5料1'] =pd.merge(report_item_cal_out, mater_M5_value1, on='核算项目号', how='left')['未税金额-之和']
            total_mater_M5_1 = report_item_cal_out.drop_duplicates(subset=['核算项目号', '返工工单号603']).reset_index(drop=True)
            total_mater_M5_1 = total_mater_M5_1.groupby(['核算项目号'])[['M5料']].sum().add_suffix('').reset_index()
            total_mater_M5_2 = report_item_cal_out.drop_duplicates(subset=['核算项目号']).reset_index(drop=True)[['核算项目号', 'M5料1']]
            total_mater_M5_2['M5料'] = pd.merge(total_mater_M5_2, total_mater_M5_1, on='核算项目号', how='left')['M5料']
            total_mater_M5_2['M5料2'] = total_mater_M5_2['M5料1'] - total_mater_M5_2['M5料']
            report_item_cal_out['M5料2'] = pd.merge(report_item_cal_out, total_mater_M5_2, on='核算项目号', how='left')['M5料2']

        if len(order_mater_M5) == 0:
            report_item_cal_out['M5工单料'] = 0
            report_item_cal_out['M5工单料2'] = 0
        else:
            report_item_cal_out['M5工单料'] = pd.merge(report_item_cal_out, order_mater_M5_value, left_on='项目号_工单2', right_on='项目号_工单', how='left')['未税金额-之和']
            report_item_cal_out['M5工单料1'] =pd.merge(report_item_cal_out, order_mater_M5_value1, on='核算项目号', how='left')['未税金额-之和']
            total_order_mater_M5_1 = report_item_cal_out.drop_duplicates(subset=['核算项目号', '返工工单号603']).reset_index(drop=True)
            total_order_mater_M5_1 = total_order_mater_M5_1.groupby(['核算项目号'])[['M5工单料']].sum().add_suffix('').reset_index()
            total_order_mater_M5_2 = report_item_cal_out.drop_duplicates(subset=['核算项目号']).reset_index(drop=True)[['核算项目号', 'M5工单料1']]
            total_order_mater_M5_2['M5工单料'] = pd.merge(total_order_mater_M5_2, total_order_mater_M5_1, on='核算项目号', how='left')['M5工单料']
            total_order_mater_M5_2['M5工单料2'] = total_order_mater_M5_2['M5工单料1'] - total_order_mater_M5_2['M5工单料']
            report_item_cal_out['M5工单料2'] = pd.merge(report_item_cal_out,total_order_mater_M5_2,on='核算项目号',how='left')['M5工单料2']

        if len(changer_mater_M5) == 0:
            report_item_cal_out['M5设变料'] = 0
            report_item_cal_out['M5设变料2'] = 0
        else:
            report_item_cal_out['M5设变料'] = pd.merge(report_item_cal_out, changer_mater_M5_value, left_on='项目号_工单2', right_on='项目号_工单', how='left')['未税金额-之和']
            report_item_cal_out['M5设变料1'] =pd.merge(report_item_cal_out, changer_mater_M5_value1, on='核算项目号', how='left')['未税金额-之和']
            total_changer_mater_M5_1 = report_item_cal_out.drop_duplicates(subset=['核算项目号', '返工工单号603']).reset_index(drop=True)
            total_changer_mater_M5_1 = total_changer_mater_M5_1.groupby(['核算项目号'])[['M5设变料']].sum().add_suffix('').reset_index()
            total_changer_mater_M5_2 = report_item_cal_out.drop_duplicates(subset=['核算项目号']).reset_index(drop=True)[['核算项目号', 'M5设变料1']]
            total_changer_mater_M5_2['M5设变料'] = pd.merge(total_changer_mater_M5_2, total_changer_mater_M5_1, on='核算项目号', how='left')['M5设变料']
            total_changer_mater_M5_2['M5设变料2'] = total_changer_mater_M5_2['M5设变料1'] - total_changer_mater_M5_2['M5设变料']
            report_item_cal_out['M5设变料2'] = pd.merge(report_item_cal_out, total_changer_mater_M5_2, on='核算项目号', how='left')['M5设变料2']
        if len(po_M5) == 0:
            report_item_cal_out['M5采购PO'] = 0
        else:
            report_item_cal_out['M5采购PO'] = pd.merge(report_item_cal_out, po_M5_value, on='核算项目号', how='left')['采购金额-未税-之和']
        if len(prod_work_M5) == 0:
            report_item_cal_out['M5生产工'] = 0
            report_item_cal_out['M5生产工2'] = 0
        else:
            report_item_cal_out['M5生产工'] = pd.merge(report_item_cal_out, prod_work_M5_value, left_on='项目号_工单2',right_on='项目号_工单', how='left')['工时成本-之和']
            report_item_cal_out['M5生产工1'] =pd.merge(report_item_cal_out, prod_work_M5_value1, on='核算项目号', how='left')['工时成本-之和']
            total_prod_work_M5_1 = report_item_cal_out.drop_duplicates(subset=['核算项目号', '返工工单号603']).reset_index(drop=True)
            total_prod_work_M5_1 = total_prod_work_M5_1.groupby(['核算项目号'])[['M5生产工']].sum().add_suffix('').reset_index()
            total_prod_work_M5_2 = report_item_cal_out.drop_duplicates(subset=['核算项目号']).reset_index(drop=True)[['核算项目号', 'M5生产工1']]
            total_prod_work_M5_2['M5生产工'] = pd.merge(total_prod_work_M5_2, total_prod_work_M5_1, on='核算项目号', how='left')['M5生产工']
            total_prod_work_M5_2['M5生产工2'] = total_prod_work_M5_2['M5生产工1'] - total_prod_work_M5_2['M5生产工']
            report_item_cal_out['M5生产工2'] = pd.merge(report_item_cal_out, total_prod_work_M5_2, on='核算项目号', how='left')['M5生产工2']
        if len(deli_work_M5) == 0:
            report_item_cal_out['M5交付工'] = 0
            report_item_cal_out['M5交付工2'] = 0
        else:
            report_item_cal_out['M5交付工'] = pd.merge(report_item_cal_out, deli_work_M5_value, left_on='项目号_工单2',right_on='项目号_工单', how='left')['工时成本-之和']
            report_item_cal_out['M5交付工1'] = pd.merge(report_item_cal_out, deli_work_M5_value1, on='核算项目号', how='left')['工时成本-之和']
            total_deli_work_M5_1 = report_item_cal_out.drop_duplicates(subset=['核算项目号', '返工工单号603']).reset_index(drop=True)
            total_deli_work_M5_1 = total_deli_work_M5_1.groupby(['核算项目号'])[['M5交付工']].sum().add_suffix('').reset_index()
            total_deli_work_M5_2 = report_item_cal_out.drop_duplicates(subset=['核算项目号']).reset_index(drop=True)[['核算项目号', 'M5交付工1']]
            total_deli_work_M5_2['M5交付工'] = \
            pd.merge(total_deli_work_M5_2, total_deli_work_M5_1, on='核算项目号', how='left')['M5交付工']
            total_deli_work_M5_2['M5交付工2'] = total_deli_work_M5_2['M5交付工1'] - total_deli_work_M5_2['M5交付工']
            report_item_cal_out['M5交付工2'] = pd.merge(report_item_cal_out, total_deli_work_M5_2, on='核算项目号', how='left')['M5交付工2']
        if len(design_work_M5) == 0:
            report_item_cal_out['M5设计工'] = 0
            report_item_cal_out['M5设计工2'] = 0
        else:
            report_item_cal_out['M5设计工'] = pd.merge(report_item_cal_out, design_work_M5_value, left_on='项目号_工单2',right_on='项目号_工单', how='left')[
                '工时成本-之和']
            report_item_cal_out['M5设计工1'] = pd.merge(report_item_cal_out, design_work_M5_value1, on='核算项目号', how='left')[
                '工时成本-之和']
            total_design_work_M5_1 = report_item_cal_out.drop_duplicates(subset=['核算项目号', '返工工单号603']).reset_index(
                drop=True)
            total_design_work_M5_1 = total_design_work_M5_1.groupby(['核算项目号'])[['M5设计工']].sum().add_suffix('').reset_index()
            total_design_work_M5_2 = report_item_cal_out.drop_duplicates(subset=['核算项目号']).reset_index(drop=True)[['核算项目号', 'M5设计工1']]
            total_design_work_M5_2['M5设计工'] =pd.merge(total_design_work_M5_2, total_design_work_M5_1, on='核算项目号', how='left')['M5设计工']
            total_design_work_M5_2['M5设计工2'] =total_design_work_M5_2['M5设计工1'] - total_design_work_M5_2['M5设计工']
            report_item_cal_out['M5设计工2'] =pd.merge(report_item_cal_out, total_design_work_M5_2, on='核算项目号', how='left')['M5设计工2']
        if len(cost_M5) == 0:
            report_item_cal_out['M5其他费'] = 0
        if len(cost_M5) == 0:
            report_item_cal_out['M5其他费'] = 0
        else:
            report_item_cal_out['M5其他费'] = pd.merge(report_item_cal_out, cost_M5_value, on='核算项目号', how='left')['金额万元-之和']

        ############出差补贴
        report_item_cal_out['出差补贴'] = pd.merge(report_item_cal_out, travel_money, on='核算项目号', how='left')['出差补贴']
        report_item_cal_out[['M2-3料','M2-3工单料','M2-3设变料', 'M2-3采购PO', 'M2-3生产工',
             'M2-3交付工', 'M2-3设计工',
             'M2-3其他费', 'M4料','M4料2','M4工单料','M4工单料2','M4设变料','M4设变料2', 'M4采购PO',
             'M4生产工', 'M4交付工', 'M4设计工','M4生产工2', 'M4交付工2', 'M4设计工2',
             'M4其他费', 'M5料','M5料2','M5工单料','M5工单料2','M5设变料','M5设变料2', 'M5采购PO',
             'M5生产工', 'M5交付工', 'M5设计工','M5生产工2', 'M5交付工2', 'M5设计工2',
             'M5其他费', '出差补贴']] = report_item_cal_out[
            ['M2-3料', 'M2-3工单料', 'M2-3设变料', 'M2-3采购PO', 'M2-3生产工',
             'M2-3交付工', 'M2-3设计工',
             'M2-3其他费', 'M4料', 'M4料2', 'M4工单料', 'M4工单料2', 'M4设变料', 'M4设变料2', 'M4采购PO',
             'M4生产工', 'M4交付工', 'M4设计工', 'M4生产工2', 'M4交付工2', 'M4设计工2',
             'M4其他费', 'M5料', 'M5料2', 'M5工单料', 'M5工单料2', 'M5设变料', 'M5设变料2', 'M5采购PO',
             'M5生产工', 'M5交付工', 'M5设计工', 'M5生产工2', 'M5交付工2', 'M5设计工2',
             'M5其他费', '出差补贴']].fillna(0)
        ####计算汇总表工单数量
        pro_num = pd.DataFrame(report_item_cal_out.groupby(['核算项目号'])['项目财经再分类'].count()).add_suffix('数量').reset_index()
        report_item_cal_out['工单数量'] = pd.merge(report_item_cal_out, pro_num, on='核算项目号', how='left')['项目财经再分类数量']
        pro_num1 = pd.DataFrame(report_item_cal_out.groupby(['项目号_工单1'])['项目财经再分类'].count()).add_suffix('数量').reset_index()
        report_item_cal_out['工单数量1'] = pd.merge(report_item_cal_out, pro_num1, on='项目号_工单1', how='left')['项目财经再分类数量']
        pro_num2 = pd.DataFrame(report_item_cal_out.groupby(['项目号_工单2'])['项目财经再分类'].count()).add_suffix('数量').reset_index()
        report_item_cal_out['工单数量2'] = pd.merge(report_item_cal_out, pro_num2, on='项目号_工单2', how='left')['项目财经再分类数量']
        ####按工单数量重新计算各种费用
        report_item_cal_out['M2-3料'] = report_item_cal_out['M2-3料'] / report_item_cal_out['工单数量']
        report_item_cal_out['M2-3工单料'] = report_item_cal_out['M2-3工单料'] / report_item_cal_out['工单数量']
        report_item_cal_out['M2-3设变料'] = report_item_cal_out['M2-3设变料'] / report_item_cal_out['工单数量']
        report_item_cal_out['M2-3采购PO'] = report_item_cal_out['M2-3采购PO'] / report_item_cal_out['工单数量']
        report_item_cal_out['M2-3生产工'] = report_item_cal_out['M2-3生产工'] / report_item_cal_out['工单数量']
        report_item_cal_out['M2-3交付工'] = report_item_cal_out['M2-3交付工'] / report_item_cal_out['工单数量']
        report_item_cal_out['M2-3设计工'] = report_item_cal_out['M2-3设计工'] / report_item_cal_out['工单数量']
        report_item_cal_out['M2-3其他费'] = report_item_cal_out['M2-3其他费'] / report_item_cal_out['工单数量']
        report_item_cal_out['出差补贴'] = report_item_cal_out['出差补贴'] / report_item_cal_out['工单数量']
        if len(mater_M4) == 0:
            report_item_cal_out['M4料'] = 0
        else:
            report_item_cal_out['M4料'] = report_item_cal_out['M4料'] / report_item_cal_out['工单数量1']
            report_item_cal_out['M4料2'] = report_item_cal_out['M4料2'] / report_item_cal_out['工单数量']
            report_item_cal_out['M4料']=report_item_cal_out['M4料'] +report_item_cal_out['M4料2']

        if len(order_mater_M4) == 0:
            report_item_cal_out['M4工单料'] = 0
        else:
            report_item_cal_out['M4工单料'] = report_item_cal_out['M4工单料'] / report_item_cal_out['工单数量1']
            report_item_cal_out['M4工单料2'] = report_item_cal_out['M4工单料2'] / report_item_cal_out['工单数量']
            report_item_cal_out['M4工单料'] = report_item_cal_out['M4工单料'] + report_item_cal_out['M4工单料2']

        if len(changer_mater_M4) == 0:
            report_item_cal_out['M4设变料'] = 0
        else:
            report_item_cal_out['M4设变料'] = report_item_cal_out['M4设变料'] / report_item_cal_out['工单数量1']
            report_item_cal_out['M4设变料2'] = report_item_cal_out['M4设变料2'] / report_item_cal_out['工单数量']
            report_item_cal_out['M4设变料'] = report_item_cal_out['M4设变料'] + report_item_cal_out['M4设变料2']

        report_item_cal_out['M4采购PO'] = report_item_cal_out['M4采购PO'] / report_item_cal_out['工单数量']
        if len(prod_work_M4) == 0:
            report_item_cal_out['M4生产工'] = 0
        else:
            report_item_cal_out['M4生产工'] = report_item_cal_out['M4生产工'] / report_item_cal_out['工单数量1']
            report_item_cal_out['M4生产工2'] = report_item_cal_out['M4生产工2'] / report_item_cal_out['工单数量']
            report_item_cal_out['M4生产工'] = report_item_cal_out['M4生产工'] + report_item_cal_out['M4生产工2']
        if len(deli_work_M4) == 0:
            report_item_cal_out['M4交付工'] = 0
        else:
            report_item_cal_out['M4交付工'] = report_item_cal_out['M4交付工'] / report_item_cal_out['工单数量1']
            report_item_cal_out['M4交付工2'] = report_item_cal_out['M4交付工2'] / report_item_cal_out['工单数量']
            report_item_cal_out['M4交付工'] = report_item_cal_out['M4交付工'] + report_item_cal_out['M4交付工2']
        if len(design_work_M4) == 0:
            report_item_cal_out['M4交付工'] = 0
        else:
            report_item_cal_out['M4设计工'] = report_item_cal_out['M4设计工'] / report_item_cal_out['工单数量1']
            report_item_cal_out['M4设计工2'] = report_item_cal_out['M4设计工2'] / report_item_cal_out['工单数量']
            report_item_cal_out['M4设计工'] = report_item_cal_out['M4设计工'] + report_item_cal_out['M4设计工2']
            report_item_cal_out['M4其他费'] = report_item_cal_out['M4其他费'] / report_item_cal_out['工单数量']
        if len(mater_M5) == 0:
            report_item_cal_out['M5料'] = 0
        else:
            report_item_cal_out['M5料'] = report_item_cal_out['M5料'] / report_item_cal_out['工单数量2']
            report_item_cal_out['M5料2'] = report_item_cal_out['M5料2'] / report_item_cal_out['工单数量']
            report_item_cal_out['M5料'] = report_item_cal_out['M5料'] + report_item_cal_out['M5料2']

        if len(order_mater_M5) == 0:
            report_item_cal_out['M5工单料'] = 0
        else:
            report_item_cal_out['M5工单料'] = report_item_cal_out['M5工单料'] / report_item_cal_out['工单数量2']
            report_item_cal_out['M5工单料2'] = report_item_cal_out['M5工单料2'] / report_item_cal_out['工单数量']
            report_item_cal_out['M5工单料'] = report_item_cal_out['M5工单料'] + report_item_cal_out['M5工单料2']

        if len(changer_mater_M5) == 0:
            report_item_cal_out['M5设变料'] = 0
        else:
            report_item_cal_out['M5设变料'] = report_item_cal_out['M5设变料'] / report_item_cal_out['工单数量2']
            report_item_cal_out['M5设变料2'] = report_item_cal_out['M5设变料2'] / report_item_cal_out['工单数量']
            report_item_cal_out['M5设变料'] = report_item_cal_out['M5设变料'] + report_item_cal_out['M5设变料2']

        report_item_cal_out['M5采购PO'] = report_item_cal_out['M5采购PO'] / report_item_cal_out['工单数量']
        if len(prod_work_M5) == 0:
            report_item_cal_out['M5生产工'] = 0
        else:
            report_item_cal_out['M5生产工'] = report_item_cal_out['M5生产工'] / report_item_cal_out['工单数量2']
            report_item_cal_out['M5生产工2'] = report_item_cal_out['M5生产工2'] / report_item_cal_out['工单数量']
            report_item_cal_out['M5生产工'] = report_item_cal_out['M5生产工'] + report_item_cal_out['M5生产工2']
        if len(deli_work_M5) == 0:
            report_item_cal_out['M5交付工'] = 0
        else:
            report_item_cal_out['M5交付工'] = report_item_cal_out['M5交付工'] / report_item_cal_out['工单数量2']
            report_item_cal_out['M5交付工2'] = report_item_cal_out['M5交付工2'] / report_item_cal_out['工单数量']
            report_item_cal_out['M5交付工'] = report_item_cal_out['M5交付工'] + report_item_cal_out['M5交付工2']
        if len(design_work_M5) == 0:
            report_item_cal_out['M5设计工'] = 0
        else:
            report_item_cal_out['M5设计工'] = report_item_cal_out['M5设计工'] / report_item_cal_out['工单数量2']
            report_item_cal_out['M5设计工2'] = report_item_cal_out['M5设计工2'] / report_item_cal_out['工单数量']
            report_item_cal_out['M5设计工'] = report_item_cal_out['M5设计工'] + report_item_cal_out['M5设计工2']
        report_item_cal_out['M5其他费'] = report_item_cal_out['M5其他费'] / report_item_cal_out['工单数量']
        ###计算汇总类费
        report_item_cal_out['M5其他费'] = report_item_cal_out['M5其他费'] + report_item_cal_out['出差补贴']
        del report_item_cal_out['出差补贴']
        report_item_cal_out['M2-3工'] = report_item_cal_out['M2-3交付工'] + report_item_cal_out['M2-3生产工']
        report_item_cal_out['M2-3费'] = report_item_cal_out['M2-3设计工'] + report_item_cal_out['M2-3其他费']
        report_item_cal_out['M2-3成本合计'] = report_item_cal_out['M2-3料'] + report_item_cal_out['M2-3费'] + report_item_cal_out['M2-3工']

        report_item_cal_out['M4工'] = report_item_cal_out['M4交付工'] + report_item_cal_out['M4生产工']
        report_item_cal_out['M4费'] = report_item_cal_out['M4设计工'] + report_item_cal_out['M4其他费']
        report_item_cal_out['M4成本合计'] = report_item_cal_out['M4料'] + report_item_cal_out['M4费'] + report_item_cal_out['M4工']

        report_item_cal_out['M5工'] = report_item_cal_out['M5交付工'] + report_item_cal_out['M5生产工']
        report_item_cal_out['M5费'] = report_item_cal_out['M5设计工'] + report_item_cal_out['M5其他费']
        report_item_cal_out['M5成本合计'] = report_item_cal_out['M5料'] + report_item_cal_out['M5费'] + report_item_cal_out['M5工']
        #
#####################################
        screm.insert(INSERT, '\n5.3：底表计算汇总类金额', '\n')
        window.update()
        report_item_cal_out = report_item_cal_out.fillna('')
        report_item_cal_out['成本'] = report_item_cal_out['M2-3成本合计']+report_item_cal_out['M4成本合计']+report_item_cal_out['M5成本合计']
        report_item_cal_out['毛利'] = 0
        report_item_cal_out['毛利率'] = 0
        report_item_cal_out['采购PO'] = report_item_cal_out['M2-3采购PO']+report_item_cal_out['M4采购PO']+report_item_cal_out['M5采购PO']
        report_item_cal_out['料'] = report_item_cal_out['M2-3料'] + report_item_cal_out['M4料'] + report_item_cal_out['M5料']
        report_item_cal_out['工单料'] = report_item_cal_out['M2-3工单料'] + report_item_cal_out['M4工单料'] + report_item_cal_out['M5工单料']
        report_item_cal_out['设变料'] = report_item_cal_out['M2-3设变料'] + report_item_cal_out['M4设变料'] + report_item_cal_out['M5设变料']
        report_item_cal_out['工'] = report_item_cal_out['M2-3工'] + report_item_cal_out['M4工'] + report_item_cal_out['M5工']
        report_item_cal_out['生产工'] = report_item_cal_out['M2-3生产工'] + report_item_cal_out['M4生产工'] + report_item_cal_out['M5生产工']
        report_item_cal_out['交付工'] = report_item_cal_out['M2-3交付工'] + report_item_cal_out['M4交付工'] +report_item_cal_out['M5交付工']
        report_item_cal_out['设计工'] = report_item_cal_out['M2-3设计工'] + report_item_cal_out['M4设计工'] +  report_item_cal_out['M5设计工']
        report_item_cal_out['费'] = report_item_cal_out['M2-3费'] + report_item_cal_out['M4费'] + report_item_cal_out['M5费']
        report_item_cal_out['其他费'] = report_item_cal_out['M2-3其他费'] + report_item_cal_out['M4其他费'] +  report_item_cal_out['M5其他费']
        ######金额取万元
        report_item_cal_out[['成本'
                ,'料','工单料','设变料','采购PO', '工', '生产工','交付工', '费', '设计工', '其他费'
                ,'M2-3成本合计','M2-3料', 'M2-3工单料','M2-3设变料','M2-3采购PO', 'M2-3工', 'M2-3生产工','M2-3交付工', 'M2-3费', 'M2-3设计工', 'M2-3其他费'
                ,'M4成本合计','M4料','M4工单料','M4设变料','M4采购PO', 'M4工','M4生产工', 'M4交付工', 'M4费', 'M4设计工', 'M4其他费'
                ,'M5成本合计','M5料','M5工单料','M5设变料','M5采购PO', 'M5工','M5生产工', 'M5交付工', 'M5费', 'M5设计工', 'M5其他费']]=report_item_cal_out[['成本'
                ,'料','工单料','设变料','采购PO', '工', '生产工','交付工', '费', '设计工', '其他费'
                ,'M2-3成本合计','M2-3料', 'M2-3工单料','M2-3设变料','M2-3采购PO', 'M2-3工', 'M2-3生产工','M2-3交付工', 'M2-3费', 'M2-3设计工', 'M2-3其他费'
                ,'M4成本合计','M4料','M4工单料','M4设变料','M4采购PO', 'M4工','M4生产工', 'M4交付工', 'M4费', 'M4设计工', 'M4其他费'
                ,'M5成本合计','M5料','M5工单料','M5设变料','M5采购PO', 'M5工','M5生产工', 'M5交付工', 'M5费', 'M5设计工', 'M5其他费']]/10000

        ##总额
        report_item_cal_out['毛利'] = report_item_cal_out['集团收入'] - report_item_cal_out['成本']
        report_item_cal_out.loc[report_item_cal_out['集团收入'] != 0, '毛利率'] = report_item_cal_out['毛利'] / report_item_cal_out['集团收入']
        report_item_cal_out.loc[report_item_cal_out['集团收入'] == 0, '毛利率'] = ''
        report_item_cal_out=report_item_cal_out.rename(columns={'工单数量': '设备总数量'})
        report_item_cal_out = report_item_cal_out.reindex(
            columns=['序列号','区域','行业中心','设备类型',"客户简称", "大项目名称", "大项目号", "产品线名称", "核算项目号",'设备名称', "项目财经","项目财经再分类",'项目数量','已出货数量','在产数量','生产状态','集团收入','软件收入','硬件收入','成本','毛利','毛利率'
                , '料','工单料','设变料', '采购PO', '工', '生产工','交付工', '费', '设计工', '其他费'
                ,'M2-3成本合计', 'M2-3料','M2-3工单料','M2-3设变料', 'M2-3采购PO', 'M2-3工', 'M2-3生产工','M2-3交付工', 'M2-3费', 'M2-3设计工', 'M2-3其他费'
                , '一般工单号601/608', '工单开立时间', '工单完工时间', 'M4成本合计', 'M4料','M4工单料','M4设变料','M4采购PO', 'M4工','M4生产工', 'M4交付工', 'M4费', 'M4设计工', 'M4其他费','系统出货时间', '实际出货时间'
                , '返工工单号603', 'M5成本合计', 'M5料','M5工单料','M5设变料','M5采购PO', 'M5工','M5生产工', 'M5交付工', 'M5费', 'M5设计工', 'M5其他费', '系统验收时间', '实际验收时间'
                ,'项目号整理', '成品料号', '设备总数量','是否预验收','全面预算有无','OA状态'])

        del report_item_cal_out['设备总数量']
        ###汇总表时间
#############################
        screm.insert(INSERT, '\n5.3：底表处理最终时间格式', '\n')
        window.update()

        ####处理四个明细表日期
        reportcost['申请时间'] = pd.to_datetime(reportcost['申请时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        reportcost['申请时间'] = ['' if i == '1990-01-01' else i for i in reportcost['申请时间']]
        reportwork['完成日期'] = pd.to_datetime(reportwork['完成日期'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        reportwork['完成日期'] = ['' if i == '1990-01-01' else i for i in reportwork['完成日期']]
        report_material['扣账日期'] = pd.to_datetime(report_material['扣账日期'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        report_material['扣账日期'] = ['' if i == '1990-01-01' else i for i in report_material['扣账日期']]
        process['采购日期'] = pd.to_datetime(process['采购日期'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        process['采购日期'] = ['' if i == '1990-01-01' else i for i in process['采购日期']]
        process['工单开立时间'] = pd.to_datetime(process['工单开立时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        process['工单开立时间'] = ['' if i == '1990-01-01' else i for i in process['工单开立时间']]

        process['工单完工时间'] = pd.to_datetime(process['工单完工时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        process['工单完工时间'] = ['' if i == '2090-01-01' else i for i in process['工单完工时间']]
        process['工单完工时间'] = ['' if i == '1990-01-01' else i for i in process['工单完工时间']]
        process = process.rename(columns={'验收日期': '实际验收时间'})
        process['实际验收时间'] = pd.to_datetime(process['实际验收时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        process['实际验收时间'] = ['' if i == '1990-01-01' else i for i in process['实际验收时间']]
        report_material['工单开立时间'] = pd.to_datetime(report_material['工单开立时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        report_material['工单开立时间'] = ['' if i == '1990-01-01' else i for i in report_material['工单开立时间']]
        report_material['工单完工时间'] = pd.to_datetime(report_material['工单完工时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        report_material['工单完工时间'] = ['' if i == '2090-01-01' else i for i in report_material['工单完工时间']]
        report_material['工单完工时间'] = ['' if i == '1990-01-01' else i for i in report_material['工单完工时间']]
        report_material = report_material.rename(columns={'验收日期': '实际验收时间'})
        report_material['实际验收时间'] = pd.to_datetime(report_material['实际验收时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        report_material['实际验收时间'] = ['' if i == '1990-01-01' else i for i in report_material['实际验收时间']]
        reportwork['工单开立时间'] = pd.to_datetime(reportwork['工单开立时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        reportwork['工单开立时间'] = ['' if i == '1990-01-01' else i for i in reportwork['工单开立时间']]
        reportwork['工单完工时间'] = pd.to_datetime(reportwork['工单完工时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        reportwork['工单完工时间'] = ['' if i == '2090-01-01' else i for i in reportwork['工单完工时间']]
        reportwork['工单完工时间'] = ['' if i == '1990-01-01' else i for i in reportwork['工单完工时间']]
        reportwork = reportwork.rename(columns={'验收日期': '实际验收时间'})
        reportwork['实际验收时间'] = pd.to_datetime(reportwork['实际验收时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        reportwork['实际验收时间'] = ['' if i == '1990-01-01' else i for i in reportwork['实际验收时间']]
        reportcost['工单开立时间'] = pd.to_datetime(reportcost['工单开立时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        reportcost['工单开立时间'] = ['' if i == '1990-01-01' else i for i in reportcost['工单开立时间']]
        reportcost['工单完工时间'] = pd.to_datetime(reportcost['工单完工时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        reportcost['工单完工时间'] = ['' if i == '2090-01-01' else i for i in reportcost['工单完工时间']]
        reportcost['工单完工时间'] = ['' if i == '1990-01-01' else i for i in reportcost['工单完工时间']]
        reportcost = reportcost.rename(columns={'验收日期': '实际验收时间'})
        reportcost['实际验收时间'] = pd.to_datetime(reportcost['实际验收时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        reportcost['实际验收时间'] = ['' if i == '1990-01-01' else i for i in reportcost['实际验收时间']]
        time_data_cal = time.time()
        screm.insert(INSERT, '\n第五阶段执行时长:%d秒' % (time_data_cal - time_stage), '\n')
#################################################################################################################################################################################
        screm.insert(INSERT, '\n六、第六阶段-数据输出', '\n')
        window.update()
        del reportcost['金额万元']
        ########################################输出文件
        l=0
        def writer_contents(sheet, array, start_row, start_col, format=None, percent_format=None, percentlist=[]):
            start_col = 0
            for col in array:
                if percentlist and (start_col in percentlist):
                    sheet.write_column(start_row, start_col, col, percent_format)
                else:
                    sheet.write_column(start_row, start_col, col, format)
                start_col += 1
        process=process.fillna('')
        report_material= report_material.fillna('')
        reportwork=reportwork.fillna('')
        reportcost = reportcost.fillna('')
####################################
        screm.insert(INSERT, '\n6.1：输出整体数据........', '\n')
        window.update()
        for  names in person_list:
            workbook = xlsxwriter.Workbook('输出文件\\'+names+'.xlsx', {'nan_inf_to_errors': True})
            worksheet0= workbook.add_worksheet('成本汇总')
            worksheet1 = workbook.add_worksheet('采购PO')
            worksheet2 = workbook.add_worksheet('料')
            worksheet4 = workbook.add_worksheet('工')
            worksheet3 = workbook.add_worksheet('费')
            title_format = workbook.add_format({'font_name': 'Arial',
                                                'font_size': 10,
                                                'font_color':'white',
                                                'bg_color':'#1F4E78',
                                                'bold': True,
                                                'bold': True,
                                                'align':'center',
                                                'valign':'vcenter',
                                                'border':1,
                                                'border_color':'white'
                                                })
            title_format.set_align('vcenter')
            title_format1 = workbook.add_format({'font_name': 'Arial',
                                                'font_size': 10,
                                                'font_color': 'white',
                                                'bg_color': '#006666',
                                                'bold': True,
                                                'bold': True,
                                                'align': 'center',
                                                'valign': 'vcenter',
                                                'border': 1,
                                                'border_color': 'white'
                                                })
            title_format1.set_align('vcenter')
            title_format2 = workbook.add_format({'font_name': 'Arial',
                                                 'font_size': 10,
                                                 'font_color': 'white',
                                                 'bg_color': '#008000',
                                                 'bold': True,
                                                 'bold': True,
                                                 'align': 'center',
                                                 'valign': 'vcenter',
                                                 'border': 1,
                                                 'border_color': 'white'
                                                 })
            title_format2.set_align('vcenter')
            title_format3= workbook.add_format({'font_name': 'Arial',
                                                 'font_size': 10,
                                                 'font_color': 'white',
                                                 'bg_color': '#7030A0',
                                                 'bold': True,
                                                 'bold': True,
                                                 'align': 'center',
                                                 'valign': 'vcenter',
                                                 'border': 1,
                                                 'border_color': 'white'
                                                 })
            title_format3.set_align('vcenter')

            title_format4 = workbook.add_format({'font_name': 'Arial',
                                                 'font_size': 10,
                                                 'font_color': 'white',
                                                 'bg_color': '#2F75B5',
                                                 'bold': True,
                                                 'bold': True,
                                                 'align': 'center',
                                                 'valign': 'vcenter',
                                                 'border': 1,
                                                 'border_color': 'white'
                                                 })
            title_format4.set_align('vcenter')
            title_format5 = workbook.add_format({'font_name': 'Arial',
                                                 'font_size': 10,
                                                 'font_color': 'white',
                                                 'bg_color': '#305496',
                                                 'bold': True,
                                                 'bold': True,
                                                 'align': 'center',
                                                 'valign': 'vcenter',
                                                 'border': 1,
                                                 'border_color': 'white'
                                                 })
            title_format5.set_align('vcenter')
            title_format6 = workbook.add_format({'font_name': 'Arial',
                                                 'font_size': 10,
                                                 'font_color': 'white',
                                                 'bg_color': '#333F4F',
                                                 'bold': True,
                                                 'bold': True,
                                                 'align': 'center',
                                                 'valign': 'vcenter',
                                                 'border': 1,
                                                 'border_color': 'white'
                                                 })
            title_format6.set_align('vcenter')
            col_format = workbook.add_format({'font_name': 'Arial',
                                                'font_size': 8,
                                                'font_color':'white',
                                                'bg_color':'#1F4E78',
                                                'text_wrap':True,
                                                'border':1,
                                                'border_color':'white',
                                                'align':'center',
                                                'valign':'vcenter'
                                                })

            data_format = workbook.add_format({'font_name': 'Arial',
                                                'font_size': 10,
                                                'align':'left',
                                                'valign':'vcenter'
                                                })
            data_format2 = workbook.add_format({'font_name': 'Arial',
                                                'font_size': 10,
                                                'align': 'center',
                                                'valign': 'vcenter'
                                                })
            data_format2.set_num_format('0.00')

            data_format1 = workbook.add_format({'font_name': 'Arial',
                                                'font_size': 10,
                                                'align':'center',
                                                'valign':'vcenter'
                                                })
            data_format_percent = workbook.add_format({'font_name': 'Arial',
                                                  'font_size': 10,
                                                  'align': 'center',
                                                  'valign': 'vcenter'
                                                  })
            data_format_percent.set_num_format('0.00%')
            num_percent_data_format = workbook.add_format({'font_name':'Arial',
                                                            'font_size': 10,
                                                            'align':'center',
                                                            'valign':'vcenter',
                                                            'num_format':'0.00%'
                                                            })
            statis_format2 = workbook.add_format({'font_name':'Arial',   #系列总计
                                                            'font_size': 9,
                                                            'align':'center',
                                                            'valign':'vcenter',
                                                            'bg_color':'#92CDDC'
                                                            })
            report_item_cal_out_cut=report_item_cal_out[report_item_cal_out['项目财经'] == names]
            report_po = process[process['项目财经'] == names]
            report_po['备注'] = report_po['备注'].replace('http', '链接：http', regex=True).astype(str)
            report_po['备注'] = report_po['备注'].replace('www', 'sss', regex=True)
            report_m = report_material[report_material['项目财经'] == names]
            report_c= reportcost[reportcost['项目财经'] == names]
            report_w = reportwork[reportwork['项目财经'] == names]

            ###处理采购PO再分类
            report_item_cal_out_cut['项目财经']= report_item_cal_out_cut['项目财经再分类']
            report_po['项目财经']=report_po['项目财经再分类']
            report_m['项目财经']=report_m['项目财经再分类']
            report_c['项目财经']=report_c['项目财经再分类']
            report_w['项目财经']=report_w['项目财经再分类']
            del report_po['项目财经再分类']
            del report_m['项目财经再分类']
            del report_c['项目财经再分类']
            del report_w['项目财经再分类']
            del report_item_cal_out_cut['项目财经再分类']
            worksheet0.write_row("A3", report_item_cal_out_cut.columns, title_format)

            worksheet0.write_row("AF3:AP3", ['M2-3成本合计', 'M2-3料','M2-3工单料','M2-3设变料', 'M2-3采购PO', 'M2-3工', 'M2-3生产工','M2-3交付工', 'M2-3费', 'M2-3设计工', 'M2-3其他费'],title_format4)
            worksheet0.write_row("AQ3:BF3",['一般工单号601/608', '工单开立时间', '工单完工时间', 'M4成本合计', 'M4料','M4工单料','M4设变料','M4采购PO', 'M4工','M4生产工', 'M4交付工', 'M4费', 'M4设计工', 'M4其他费','系统出货时间', '实际出货时间'], title_format5)
            worksheet0.write_row("BG3:BT3",['返工工单号603', 'M5成本合计', 'M5料','M5工单料','M5设变料','M5采购PO', 'M5工','M5生产工', 'M5交付工', 'M5费', 'M5设计工', 'M5其他费', '系统验收时间', '实际验收时间'],title_format6)

            writer_contents(sheet=worksheet0, array=report_item_cal_out_cut.T.values, start_row=3,start_col=0)
            worksheet0.merge_range("A1:BY1",
                                   "成本汇总——————————————————————————成本汇总——————————————————————————成本汇总——————————————————————————成本汇总——————————————————————————成本汇总———————————————————————————成本汇总——————————————————————————成本汇总——————————————————————————成本汇总——————————————————————————成本汇总——————————————————————————成本汇总——————————————————————————成本汇总",
                                   title_format)

            worksheet0.merge_range("A2:O2", "项目基本信息", title_format)
            worksheet0.merge_range("P2:AE2", "项目核算毛利", title_format)
           # worksheet0.merge_range("Q2:X2", "M2-5阶段明细汇总", title_format)
            worksheet0.merge_range("AF2:AP2", "M2-3阶段信息", title_format4)
            worksheet0.merge_range("AQ2:BF2", "M4阶段信息", title_format5)
            worksheet0.merge_range("BG2:BT2", "M5阶段信息", title_format6)
            worksheet0.merge_range("BU2:BY2", "辅助信息", title_format)

            worksheet0.set_row(0, 25)
            worksheet0.set_row(1, 22)
            worksheet0.set_column('A:D', 8, data_format)
            worksheet0.set_column('E:J', 12, data_format)
            worksheet0.set_column('K:M', 10, data_format2)
            worksheet0.set_column('U:U', 8, data_format_percent)
            worksheet0.set_column('P:T', 8, data_format2)
            worksheet0.set_column('V:AP', 8, data_format2)
            worksheet0.set_column('AQ:AS', 11, data_format)
            worksheet0.set_column('AT:BD', 8, data_format2)
            worksheet0.set_column('BE:BG', 11, data_format)
            worksheet0.set_column('BH:BR', 8, data_format2)
            worksheet0.set_column('BS:BV', 11, data_format)
            worksheet0.set_column('BW:BY', 8, data_format)
            worksheet1.write_row("A1", report_po.columns, title_format)
            writer_contents(sheet=worksheet1, array=report_po.T.values, start_row=1,
                            start_col=0)
            worksheet1.set_row(0, 24)

            worksheet2.write_row("A1", report_m.columns, title_format)
            writer_contents(sheet=worksheet2, array=report_m.T.values, start_row=1,start_col=0)
            worksheet2.set_row(0, 24)
            worksheet4.write_row("A1", report_w.columns, title_format)
            writer_contents(sheet=worksheet4, array=report_w.T.values, start_row=1,start_col=0)
            worksheet4.set_row(0, 24)
            worksheet3.write_row("A1", report_c.columns, title_format)
            writer_contents(sheet=worksheet3, array=report_c.T.values, start_row=1,start_col=0)
            worksheet3.set_row(0, 24)
            workbook.close()
            l=l+1
            screm.insert(INSERT, '\n第'+str(l)+'份明细文件已生成： ' +names, '\n')
            window.update()
####################################
        screm.insert(INSERT, '\n6.2：输出新增+历史全部数据........', '\n')
        window.update()
        outfile_po = open('输出历史总文件\采购PO\截止当前所有采购PO.csv', 'wb')
        process.to_csv(outfile_po, index=False, encoding='gb18030')
        outfile_mater = open('输出历史总文件\料\截止当前所有料.csv', 'wb')
        report_material.to_csv(outfile_mater, index=False, encoding='gb18030')
        outfile_work = open('输出历史总文件\工\截止当前所有工.csv', 'wb')
        reportwork.to_csv(outfile_work, index=False, encoding='gb18030')
        outfile_cost = open('输出历史总文件\费\截止当前所有费.csv', 'wb')
        reportcost.to_csv(outfile_cost, index=False, encoding='gb18030')

        screm.insert(INSERT, '\n程序执行完成100%,当前执行无异常', '\n')
        window.update()
        image_file = ImageTk.PhotoImage(file=r'软件附带文件\背景.jpg')
        frame.create_image(300, 0, anchor='n', image=image_file)
        frame.create_image(300, 450, anchor='n', image=image_file)
        def pick1(event):
            global a, flag
            while 1:
                im = Image.open(r'软件附带文件\绿.gif')
                iter = ImageSequence.Iterator(im)
                for jp in iter:
                    pic = ImageTk.PhotoImage(jp)
                    pic_act = frame.create_image((470, 130), image=pic)
                    time.sleep(0.1)
                    window.update_idletasks()  # 刷新
                    window.update()
        frame.bind("<Enter>", pick1)
    except Exception as f:
        # print('异常信息为:', e)  # 异常信息为: division by zero
        #print('——---#@*&程序报错，异常信息为:' + traceback.format_exc())
        screm.insert(INSERT, '\n——---#@*&程序报错，异常信息为:' + traceback.format_exc(), '\n')
        window.update()
        image_file = ImageTk.PhotoImage(file=r'软件附带文件\背景.jpg')
        frame.create_image(300, 0, anchor='n', image=image_file)
        frame.create_image(300, 450, anchor='n', image=image_file)
        def pick2(event):
            global a, flag
            while 1:
                im = Image.open(r'软件附带文件\红.gif')
                # GIF图片流的迭代器
                iter = ImageSequence.Iterator(im)
                # frame就是gif的每一帧，转换一下格式就能显示了
                for jp in iter:
                    pic = ImageTk.PhotoImage(jp)
                    pic_act = frame.create_image((470, 130), image=pic)
                    time.sleep(0.1)
                    window.update_idletasks()  # 刷新
                    window.update()
        frame.bind("<Enter>", pick2)


    ####################################
    screm.insert(INSERT, '\n6.3：输出异常清单........', '\n')
    window.update()
    writer = pd.ExcelWriter("异常清单.xlsx")
    process_err = process[['项目编号', '项目号整理', '核算项目号', '数据-来源']]
    process_err=process_err.drop_duplicates(subset=['项目编号']).reset_index(drop=True)
    process_err=process_err[process_err['核算项目号']==''].reset_index(drop=True)
    process_err.to_excel(writer, index=False,sheet_name='采购PO未核算项目')

    mater_err=report_material[['项目编号', '项目号整理','工单号码', '核算项目号', '领料类型','数据-来源']].drop_duplicates(subset=['项目编号']).reset_index(drop=True)
    mater_err=mater_err[mater_err['核算项目号']==''].reset_index(drop=True)
    mater_err.to_excel(writer, index=False, sheet_name='料未核算项目')

    work_err=reportwork[['项目号', '项目号整理', '核算项目号','数据分布']].drop_duplicates(subset=['项目号']).reset_index(drop=True)
    work_err=work_err[work_err['核算项目号']==''].reset_index(drop=True)
    work_err.to_excel(writer, index=False, sheet_name='工未核算项目')

    cost_err=reportcost[['项目号', '项目号整理', '核算项目号','数据分布']].drop_duplicates(subset=['项目号']).reset_index(drop=True)
    cost_err=cost_err[cost_err['核算项目号']==''].reset_index(drop=True)
    cost_err.to_excel(writer,index=False,sheet_name='费未核算项目')

    ###输出剔除
    process_delete_out.to_excel(writer,index=False,sheet_name='采购PO-剔除+调账库')
    report_material_delete_out.to_excel(writer, index=False, sheet_name='料-剔除+调账库')
    writer.save()
    time_end = time.time()
    screm.insert(INSERT, '\n总执行时长:%d秒' % (time_end - time_start), '\n')
    ####################################
def t2():
    button_execute['command']=execute
thread1 = threading.Thread(name='t1', target=pick)
thread2 = threading.Thread(name='t2', target=t2)
thread1.start()  # 启动线程1
thread2.start()
window.mainloop()
