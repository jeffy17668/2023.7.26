import pandas as pd
#coding:utf-8
import zipfile
import threading
import pandas as pd
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
warnings.filterwarnings('ignore')
time_start=time.time()
try:
    ###############################################################################################################################################一
    time_start = time.time()
    print("一、数据读取..")

    print("1.1：正在读取核算进度底表..")
    #########################读取核算进度底表
    filePath_base = r'01：核算-进度底表'
    file_name_base = os.listdir(filePath_base)
    for i in range(len(file_name_base)):
        if str(file_name_base[i]).count('~$') == 0:
            # order_new = pd.read_excel(filePathT_1 + '/' + str(file_name1[i]), header=2)
            base = pd.read_excel(filePath_base + '/' + str(file_name_base[i]))

    print("1.2：正在读取概预算表..")
    #########################读取概预算表
    path_line = r'02：概预算'
    index = 0
    line = []
    line_file = os.listdir(path_line)
    for i in line_file:
        if '~$' in i:
            line_file.remove(i)
    for name in line_file:
        if 'xlsx' in name and '~$' not in name:
            budget_estimate = pd.read_excel(path_line + '\\' + name, sheet_name='概预算', header=1)
            year_budget = pd.read_excel(path_line + '\\' + name, sheet_name='年初预算表', header=2)

    print("1.3：正在读取核算汇总表..")
    #########################读取核算汇总表
    path_calcu = r'03：核算汇总表'
    calcu_file = os.listdir(path_calcu)
    for i in calcu_file:
        if '~$' in i:
            calcu_file.remove(i)
    for name in calcu_file:
        if 'xlsx' in name and '~$' not in name:
            calcu = pd.read_excel(path_calcu + '\\' + name, header=2)

    print("1.4：正在读取财务成本表..")
    #########################读取财务成本表
    path_finance = r'04：财务成本'
    finance_file = os.listdir(path_finance)
    for i in finance_file:
        if '~$' in i:
            finance_file.remove(i)
    for name in finance_file:
        if 'xlsx' in name and '~$' not in name:
            finance = pd.read_excel(path_finance + '\\' + name, header=2, sheet_name='收入成本表')

    print("1.5：正在读取周进度表..")
    #########################读取周进度表
    path_advance = r'05：周进度表'
    advance_file = os.listdir(path_advance)
    for i in advance_file:
        if '~$' in i:
            advance_file.remove(i)
    for name in advance_file:
        if 'xlsx' in name and '~$' not in name:
            advance = pd.read_excel(path_advance + '\\' + name, header=2, sheet_name='明细表')

    print("1.6：正在读取还需表..")
    #########################读取周进度表
    path_need = r'06：还需'
    need_file = os.listdir(path_need)
    for i in need_file:
        if '~$' in i:
            need_file.remove(i)
    for name in need_file:
        if 'xlsx' in name and '~$' not in name:
            need = pd.read_excel(path_need + '\\' + name)

    print("1.7：正在读取存货表..")
    #########################读取存货表
    path_inventory = r'07：存货'
    inventory_file = os.listdir(path_inventory)
    for i in inventory_file:
        if '~$' in i:
            inventory_file.remove(i)
    for name in inventory_file:
        if 'xlsm' in name and '~$' not in name:
            inventory = pd.read_excel(path_inventory + '\\' + name, header=3)

    time_data_read = time.time()
    print('第一阶段【数据读取】执行时长:%d秒' % (time_data_read - time_start))

    ######################################################################################################################################二
    print("二、数据格式处理..")

    base = base[['序列号', '区域', '行业中心', '设备类型', '客户简称', '大项目名称', '大项目号', '产品线名称', '产品线编码',
                 '核算项目号', '设备名称', '项目数量', '已出货数量', '在产数量', '生产状态', '集团收入', '软件收入',
                 '硬件收入', '一般工单号601/608', '工单开立时间', '工单完工时间', '系统出货时间', '实际出货时间',
                 '返工工单号603', '系统验收时间', '实际验收时间', '项目号整理', '成品料号', '是否预验收', '全面预算有无',
                 'OA状态', '自制/外包', '项目财经','是否预验未终验','子项目状态']]

    print("2.1：核算进度底表..")
    base_str = ['序列号', '区域', '行业中心', '设备类型', '客户简称', '大项目名称', '大项目号', '产品线名称', '产品线编码',
                '核算项目号', '设备名称', '生产状态', '一般工单号601/608', '返工工单号603', '项目号整理', '成品料号', '是否预验收', '全面预算有无',
                'OA状态', '自制/外包', '项目财经','是否预验未终验','子项目状态']
    base[base_str] = base[base_str].fillna('')
    base['大项目名称'] = base['大项目名称'].str.strip()
    base['大项目名称'] = base['大项目名称'].replace(' ', '', regex=True).astype(str)


    base_num = ['项目数量', '已出货数量', '在产数量', '集团收入', '软件收入', '硬件收入']
    base[base_num] = base[base_num].fillna(0)

    base_date = ['工单开立时间', '工单完工时间', '系统出货时间', '实际出货时间', '系统验收时间', '实际验收时间']
    default_date = pd.Timestamp(1990, 1, 1)
    for dat in base_date:
        base[dat] = base[dat].fillna(default_date)
        base[dat] = pd.to_datetime(base[dat], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        base[dat] = ['' if i == '1990-01-01' else i for i in base[dat]]


    print("2.2：概预算表..")
    budget_estimate_str = ['大项目号', '项目号', '设备名称', '生产料号', '类型']
    budget_estimate[budget_estimate_str] = budget_estimate[budget_estimate_str].fillna('')
    budget_estimate['大项目号'] = budget_estimate['大项目号'] .str.strip()
    budget_estimate['大项目号'] = budget_estimate['大项目号'].replace(' ', '', regex=True).astype(str)
    budget_estimate['项目号'] = budget_estimate['项目号'].str.strip()
    budget_estimate['项目号'] = budget_estimate['项目号'].replace(' ', '', regex=True).astype(str)

    budget_estimate_num = ['设备数量', '成本金额', '料', '生产工', '交付工', '设计工', '项目工', '其他', '制费']
    budget_estimate[budget_estimate_num] = budget_estimate[budget_estimate_num].fillna(0)
    estimate = budget_estimate[budget_estimate['类型'].str.contains('概算')].reset_index(drop=True)
    budget = budget_estimate[budget_estimate['类型'].str.contains('预算')].reset_index(drop=True)

    year_budget_str = ['项目号整理', '归属', '客户', '线体', '大项目', '产品线编码', '产品线', '核算项目号', '设备名称-整理',
                       '产能', '自制/外包', '生产主体', '销售主体', '业务', '项目经理', '产品经理', '产品\n类型', '全面预算\n有无']
    year_budget[year_budget_str] = year_budget[year_budget_str].fillna('')
    year_budget['线体'] = year_budget['线体'].str.strip()
    year_budget['线体'] = year_budget['线体'].replace(' ', '', regex=True).astype(str)
    year_budget['核算项目号'] = year_budget['核算项目号'].str.strip()
    year_budget['核算项目号'] = year_budget['核算项目号'].replace(' ', '', regex=True).astype(str)
    year_budget_num = ['数量', '成本合计', '料', '工', '生产工', '交付工', '费', '设计工', '其他费']
    year_budget[year_budget_num] = year_budget[year_budget_num].fillna(0)

    print("2.3：核算汇总表..")
    calcu = calcu[['序列号', '区域', '行业中心', '设备类型', '客户简称', '大项目名称', '大项目号', '产品线名称', '核算项目号',
                   '设备名称', '项目财经', '项目数量', '已出货数量', '在产数量', '生产状态', '集团收入', '软件收入', '硬件收入',
                   '成本', '毛利', '毛利率', '料','工单料','设变料', '采购PO', '工', '生产工', '交付工', '费', '设计工', '其他费',
                   '一般工单号601/608', '工单开立时间', '工单完工时间', '系统出货时间', '实际出货时间', '返工工单号603'
        , '系统验收时间', '实际验收时间', '项目号整理', '成品料号', '是否预验收', '全面预算有无', 'OA状态']]

    calcu_str = ['序列号', '区域', '行业中心', '设备类型', '客户简称', '大项目名称', '大项目号', '产品线名称'
        , '核算项目号', '设备名称', '项目财经', '生产状态', '一般工单号601/608', '返工工单号603', '项目号整理'
        , '成品料号', '是否预验收', '全面预算有无', 'OA状态']
    calcu[calcu_str] = calcu[calcu_str].fillna('')

    calcu_num = ['项目数量', '已出货数量', '在产数量', '集团收入', '软件收入', '硬件收入',
                 '成本', '毛利', '毛利率', '料','工单料','设变料', '采购PO', '工', '生产工', '交付工', '费', '设计工', '其他费']
    calcu[calcu_num] = calcu[calcu_num].fillna(0)

    calcu_date = ['工单开立时间', '工单完工时间', '系统出货时间', '实际出货时间', '系统验收时间', '实际验收时间']
    default_date = pd.Timestamp(1990, 1, 1)
    for dat in calcu_date:
        calcu[dat] = calcu[dat].fillna(default_date)
        calcu[dat] = pd.to_datetime(calcu[dat], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        calcu[dat] = ['' if i == '1990-01-01' else i for i in calcu[dat]]

    print("2.4：财务成本表..")
    finance = finance[['收入日期', '公司代码', '公司 \n简称', '中心', '内部关联交易', '收入类别', '销售类型', '行业类别',
                       '区域', '内销/外销', '是否报关（是/否）', '报关单号', 'PO（订单号）', '项目号', '客户名称', '业务员',
                       '产品编码', '产品名称', '规格型号', '数量', '合并收入', '合并料', '合并工', '合并费', '合并成本合计']]
    finance_str = ['收入日期', '公司代码', '公司 \n简称', '中心', '内部关联交易', '收入类别', '销售类型', '行业类别',
                   '区域', '内销/外销', '是否报关（是/否）', '报关单号', 'PO（订单号）', '项目号', '客户名称', '业务员','产品编码', '产品名称', '规格型号']
    finance[finance_str] = finance[finance_str].fillna('')
    finance=finance[finance['项目号']!=''].reset_index(drop=True)
    finance['项目号'] = finance['项目号'].str.strip()
    finance['项目号']= finance['项目号'].replace(' ', '', regex=True).astype(str)

    finance_num = ['数量', '合并收入', '合并料', '合并工', '合并费', '合并成本合计']
    finance[finance_num] = finance[finance_num].fillna(0)
    finance[['合并收入', '合并料', '合并工', '合并费', '合并成本合计']]=finance[['合并收入', '合并料', '合并工', '合并费', '合并成本合计']]/10000
    finance.loc[finance['项目号'].astype(str).str.contains('-R'),'数量']=0
    finance.loc[finance['产品编码'].astype(str).str.contains('311-'), '数量'] = 0

    print("2.5：周进度表..")
    advance = advance[['序列号', '设备类型', '客户简称', '大项目名称', '大项目号', '产品线名称', '核算项目号', '设备名称'
        , '项目财经', '项目经理', '项目数量', '已出货数量', '在产数量', '生产状态', '集团收入', '实际出货时间'
        , '实际验收时间', '看板实际验收时间', '项目号整理', '成品料号', 'OA状态', '区域', '项目阶段', '姓名'
        , '23年预算出货时间', '计划出货时间', '23年预算验收时间', '产品线计划验收时间', '关键问题或风险点'
        , '一览表进度', '是否有风险', '原因分类', '原因大类', '原因小类', 'PC备注', '生产实际进度', '风险等级'
        , '风险分类', '验收实际进度']]
    advance_str = ['序列号', '设备类型', '客户简称', '大项目名称', '大项目号', '产品线名称', '核算项目号', '设备名称'
        , '项目财经', '项目经理', '生产状态', '项目号整理', '成品料号', 'OA状态', '区域', '项目阶段', '姓名'
        , '关键问题或风险点', '一览表进度', '是否有风险', '原因分类', '原因大类', '原因小类', 'PC备注', '生产实际进度'
        , '风险等级', '风险分类', '验收实际进度']
    advance[advance_str] = advance[advance_str].fillna('')

    advance_num = ['项目数量', '已出货数量', '在产数量', '集团收入']
    advance[advance_num] = advance[advance_num].fillna(0)

    advance_date = ['实际出货时间', '实际验收时间', '看板实际验收时间', '23年预算出货时间'
        , '计划出货时间', '23年预算验收时间', '产品线计划验收时间']
    for dat in advance_date:
        advance[dat] = advance[dat].fillna(default_date)
        advance[dat] = pd.to_datetime(advance[dat], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
        advance[dat] = ['' if i == '1990-01-01' else i for i in advance[dat]]

    print("2.6：还需..")
    need = need[['客户简称', '大项目名称', '大项目号', '产品线名称', '核算项目号', '已出货未验收数量', '成本', '料', '工', '生产工', '交付工', '费','设计工','其他费']]
    need_str = ['客户简称', '大项目名称', '大项目号', '产品线名称', '核算项目号']
    need[need_str] = need[need_str].fillna('')
    need['核算项目号'] = need['核算项目号'].str.strip()
    need['核算项目号']  = need['核算项目号'].replace(' ', '', regex=True).astype(str)
    need['大项目号'] = need['大项目号'].str.strip()
    need['大项目号'] = need['大项目号'].replace(' ', '', regex=True).astype(str)
    need['大项目名称'] = need['大项目名称'].str.strip()
    need['大项目名称'] = need['大项目名称'].replace(' ', '', regex=True).astype(str)

    need_num = ['已出货未验收数量', '成本', '料', '工', '生产工', '交付工', '费', '设计工', '其他费']
    need[need_num] = need[need_num].fillna(0)

    print("2.7：存货..")
    inventory = inventory[['主体', '中心', '存货大类（重分类前）', '存货大类（重分类后）', '产品类别（原材料除外）', '项目号/批号',
                           '存货编码', '存货名称', '规格型号', '结存数量', '结存金额', '料', '工', '费', '合计']]
    inventory_str = ['主体','中心','存货大类（重分类前）','存货大类（重分类后）','产品类别（原材料除外）', '项目号/批号',
                     '存货编码', '存货名称', '规格型号']
    inventory[inventory_str] = inventory[inventory_str].fillna('')
    inventory['项目号/批号'] = inventory['项目号/批号'].str.strip()
    inventory['项目号/批号'] = inventory['项目号/批号'].replace(' ', '', regex=True).astype(str)
    inventory['存货编码'] = inventory['存货编码'].str.strip()
    inventory['存货编码'] = inventory['存货编码'].replace(' ', '', regex=True).astype(str)

    inventory_num = ['结存数量', '结存金额', '料', '工', '费', '合计']
    inventory[inventory_num] = inventory[inventory_num].fillna(0)
    inventory[['结存金额', '料', '工', '费', '合计']]=inventory[['结存金额', '料', '工', '费', '合计']]/10000
    inventory.loc[inventory['存货编码'].astype(str).str.contains('311-'), '结存数量'] = 0
    inventory.loc[inventory['项目号/批号'].astype(str).str.contains('-R'), '结存数量'] = 0
    inventory = inventory[(inventory['项目号/批号']!="")].reset_index(drop=True)
    time_data_format = time.time()
    print('第二阶段【数据格式处理】执行时长:%d秒' % (time_data_format - time_data_read))

    ######################################################################################################################################三
    print("三、生成概预核决汇总表..")

    print("3.1：生成概预核决汇总表基础数据(根据底表)..")
    total_out = base.copy()
    total_out = total_out[['序列号', '区域', '行业中心', '客户简称', '大项目名称', '大项目号', '产品线名称'
        , '核算项目号', '设备名称', '项目数量', '已出货数量', '在产数量', '生产状态', '全面预算有无', '设备类型', '集团收入'
        , '软件收入', '硬件收入', '成品料号', '一般工单号601/608', '返工工单号603', '实际出货时间'
        , '实际验收时间', '系统验收时间', '项目号整理', '是否预验收','是否预验未终验','子项目状态']]

    print("3.2：拉取概算数据..")
    total_out['大小项目'] = total_out['大项目号'] + total_out['核算项目号']
    total_out_group = total_out.groupby(['大小项目']).agg({'核算项目号': "count"}).add_suffix('数量').reset_index()
    total_out['大小项目数量'] = pd.merge(total_out, total_out_group, on='大小项目', how='left')['核算项目号数量']

    estimate['大小项目'] = estimate['大项目号'] + estimate['项目号']
    estimate_group = estimate.groupby(['大小项目']).agg(
        {'成本金额': "sum", '料': "sum", '生产工': "sum", '交付工': "sum", '设计工': "sum", '制费': "sum"}).add_suffix('').reset_index()
    total_out[['成本合计-概算', '料-概算', '生产工-概算', '交付工-概算', '设计工-概算',  '制费-概算']] = \
    pd.merge(total_out, estimate_group, on='大小项目', how='left')[['成本金额', '料', '生产工', '交付工', '设计工',  '制费']]

    estimate_num = ['成本合计-概算', '料-概算', '生产工-概算', '交付工-概算', '设计工-概算', '制费-概算']
    for num in estimate_num:
        total_out[num] = total_out[num] / total_out['大小项目数量']

    print("3.3：拉取预算数据..")
    budget['大小项目'] = budget['大项目号'] + budget['项目号']
    budget_group = budget.groupby(['大小项目']).agg(
        {'成本金额': "sum", '料': "sum", '生产工': "sum", '交付工': "sum", '设计工': "sum", '制费': "sum"}).add_suffix('').reset_index()

    total_out[['成本合计-预算', '料-预算', '生产工-预算', '交付工-预算', '设计工-预算',  '制费-预算']] =pd.merge(total_out, budget_group, on='大小项目', how='left')[['成本金额', '料', '生产工', '交付工', '设计工','制费']]

    budget_num = ['成本合计-预算', '料-预算', '生产工-预算', '交付工-预算', '设计工-预算',  '制费-预算']
    for num in budget_num:
        total_out[num] = total_out[num]/total_out['大小项目数量']

    print("3.4：拉取财务成本表..")
    finance_use = finance[['项目号', '数量', '合并料', '合并工', '合并费', '合并成本合计']]
    finance_use['项目号'] = finance_use['项目号'].fillna('').astype(str)
    if len(finance_use[finance_use['项目号'].str.contains('-')]) > 0:
        finance_use['项目号整理'] = ''
        finance_use['项目整'] = finance_use['项目号'].str.split('-', expand=True)[0]
        finance_use['项目整1'] = finance_use['项目号'].str.split('-', expand=True)[1]
        finance_use['项目整1'] = finance_use['项目整1'].fillna('空值')
        finance_use['项目号整理'] = finance_use['项目整']
        finance_use.loc[(finance_use['项目整1'].str.isdigit()) | (finance_use['项目整1'].str.contains('SH')), '项目号整理'] = \
        finance_use['项目号整理'] + '-' + finance_use['项目整1']
    if len(finance_use[finance_use['项目号'].str.contains('-')]) == 0:
        finance_use['项目号整理'] = finance_use['项目号']
    finance_use.loc[(finance_use['项目号整理'].str[0] == 'F') & (finance_use['项目号整理'].str[:3].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX')), '项目号整理'] = finance_use['项目号整理'].str[3:]
    finance_use.loc[(finance_use['项目号整理'].str[:2].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX', na=False)), '项目号整理'] =finance_use['项目号整理'].str[2:]

    finance_use_group = finance_use.groupby(['项目号整理']).agg({'数量': 'sum', '合并成本合计': "sum", '合并料': "sum", '合并工': "sum", '合并费': "sum"}).add_suffix('').reset_index()
    finance_num = ['合并料', '合并工', '合并费', '合并成本合计']
    for num in finance_num:
        finance_use_group.loc[finance_use_group['数量']!=0,num] = finance_use_group[num] / finance_use_group['数量']

    total_out[['成本合计-财务', '料-财务', '工-财务', '费-财务']] = pd.merge(total_out, finance_use_group, on=['项目号整理'], how='left')[['合并成本合计', '合并料', '合并工', '合并费']]

    print("3.5：拉取核算汇总表..")
    calcu_use = calcu[['序列号', '采购PO','成本', '料','工单料','设变料', '工', '生产工', '交付工', '费', '设计工', '其他费', '毛利', '毛利率']]
    total_out[['采购PO','成本合计-核算', '料-核算','工单料','设变料', '工-核算', '生产工-核算', '交付工-核算', '费-核算', '设计工-核算', '其他费-核算', '毛利-核算', '毛利率-核算']]=pd.merge(total_out, calcu_use, on='序列号', how='left')[['采购PO','成本', '料','工单料','设变料', '工', '生产工', '交付工', '费', '设计工', '其他费', '毛利', '毛利率']]

    print("3.6：拉取进度表..")
    advance_use = advance[['序列号', '项目经理', '项目阶段', '是否有风险', '风险分类', '生产实际进度', '验收实际进度','计划出货时间', '产品线计划验收时间']]
    total_out[['项目经理', '项目阶段', '是否有风险', '风险分类', '生产实际进度', '验收实际进度','计划出货时间', '产品线计划验收时间']] = pd.merge(total_out, advance_use, on='序列号', how='left')[['项目经理', '项目阶段', '是否有风险', '风险分类', '生产实际进度', '验收实际进度','计划出货时间', '产品线计划验收时间']]
    total_out[['项目经理', '项目阶段', '是否有风险', '风险分类', '生产实际进度', '验收实际进度']]=total_out[['项目经理', '项目阶段', '是否有风险', '风险分类', '生产实际进度', '验收实际进度']].fillna('')
    print("3.7：拉取年初预算表..")
    year_budget['大小项目'] = year_budget['线体'] + year_budget['核算项目号']
    year_budget_group = year_budget.groupby(['大小项目']).agg(
        {'成本合计': "sum", '料': "sum", '工': "sum", '生产工': "sum", '交付工': "sum", '费': "sum", '设计工': "sum",'其他费': "sum"}).add_suffix('').reset_index()
    total_out[['成本合计-年初', '料-年初', '工-年初', '生产工-年初', '交付工-年初', '费-年初', '设计工-年初', '其他费-年初']] = \
    pd.merge(total_out, year_budget_group, on='大小项目', how='left')[['成本合计', '料', '工', '生产工', '交付工', '费', '设计工', '其他费']]
    year_budget_num = ['成本合计-年初', '料-年初', '工-年初', '生产工-年初', '交付工-年初', '费-年初', '设计工-年初', '其他费-年初']
    for num in year_budget_num:
        total_out[num] = total_out[num] / total_out['大小项目数量']

    print("3.8：拉取还需表..")
    need_use = need[['大项目名称', '核算项目号', '成本', '料', '工', '生产工', '交付工', '费','设计工', '其他费']]
    need_east=need_use[need_use['核算项目号']!=''].reset_index(drop=True)
    need_north=need_use[need_use['核算项目号']==''].reset_index(drop=True)
    ####给华东还需做项目号整理
    if len(need_east[need_east['核算项目号'].str.contains('-')]) > 0:
        need_east['项目号整理'] = ''
        need_east['项目整'] = need_east['核算项目号'].str.split('-', expand=True)[0]
        need_east['项目整1'] = need_east['核算项目号'].str.split('-', expand=True)[1]
        need_east['项目整1'] = need_east['项目整1'].fillna('空值')
        need_east['项目号整理'] = need_east['项目整']
        need_east.loc[(need_east['项目整1'].str.isdigit()) | (need_east['项目整1'].str.contains('SH')), '项目号整理'] = \
            need_east['项目号整理'] + '-' + need_east['项目整1']
    if len(need_east[need_east['核算项目号'].str.contains('-')]) == 0:
        need_east['项目号整理'] = need_east['核算项目号']
    need_east.loc[(need_east['项目号整理'].str[0] == 'F') & (
        need_east['项目号整理'].str[:3].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX')), '项目号整理'] = need_east['项目号整理'].str[3:]
    need_east.loc[(need_east['项目号整理'].str[:2].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX', na=False)), '项目号整理'] = \
        need_east['项目号整理'].str[2:]
    need_north_group = need_north.groupby(['大项目名称']).agg({'交付工': "sum"}).add_suffix('').reset_index()####h核算项目号
    need_east_group = need_east.groupby(['项目号整理']).agg({'成本': "sum", '料': "sum", '工': "sum", '生产工': "sum", '交付工': "sum", '费': "sum", '设计工': "sum",'其他费': "sum"}).add_suffix('').reset_index()####大项目名称

    #####将汇总表排序并做拆分
    total_out['排个序']=0
    for i in range(len(total_out)):
        total_out.loc[i,'排个序']=i
        i=i+1
    ####拆出未出货、已出货未验收、已验收、子项目
    ##在产
    total_out_a=total_out[total_out['生产状态'].str.contains('在产')].reset_index(drop=True)#####预算-核算
    ###出货未验收
    total_out_b1 = total_out[total_out['生产状态'].str.contains('已出货')].reset_index(drop=True)
    total_out_b2 = total_out[(total_out['生产状态'].str.contains('已验收')) & (total_out['是否预验未终验'].str.contains('是'))].reset_index(drop=True)
    total_out_b=pd.concat([total_out_b1,total_out_b2]).reset_index(drop=True)
    # 子项目
    total_out_c = total_out[total_out['生产状态'].str.contains('子项目')].reset_index(drop=True)
    #已验收
    total_out_d= total_out[(total_out['生产状态'].str.contains('已验收'))&(total_out['是否预验未终验'].str.contains('是')==False)].reset_index(drop=True)

    #####华东
    total_out_b3=total_out_b[total_out_b['产品线名称'].str.contains('大装配线|干燥产品线')==False].reset_index(drop=True)
    #total_out_b3=total_out_b[total_out_b['项目号整理'].isin(need_east_group['项目号整理'])].reset_index(drop=True)
    total_out_b3_group=total_out_b3.groupby(['项目号整理']).agg({'项目数量': "sum"}).add_suffix('统计').reset_index()  ####h核算项目号
    total_out_b3['小项目数量']=pd.merge(total_out_b3,total_out_b3_group,on='项目号整理',how='left')['项目数量统计']
    #####华南
    #total_out_b4= total_out_b[~total_out_b['项目号整理'].isin(need_east_group['项目号整理'])].reset_index(drop=True)
    total_out_b4= total_out_b[total_out_b['产品线名称'].str.contains('大装配线|干燥产品线')].reset_index(drop=True)
    total_out_b4_group = total_out_b4.groupby(['大项目名称']).agg({'项目数量': "sum"}).add_suffix('统计').reset_index()  ####大项目名称
    total_out_b4['大项目数量'] = pd.merge(total_out_b4, total_out_b4_group, on='大项目名称', how='left')['项目数量统计']

    total_out_b3[['成本合计-还需', '料-还需', '工-还需', '生产工-还需', '交付工-还需', '费-还需', '设计工-还需', '其他费-还需']]= \
    pd.merge(total_out_b3, need_east_group, on='项目号整理', how='left')[['成本', '料', '工', '生产工', '交付工', '费', '设计工', '其他费']]

    total_out_b4['交付工-还需'] = pd.merge(total_out_b4, need_north_group, on='大项目名称', how='left')['交付工']
    total_out_b4['交付工-还需']=total_out_b4['交付工-还需'].fillna(0)
    total_out_b4['交付工-还需'] =total_out_b4['交付工-还需']/total_out_b4['大项目数量']
    total_out_b4['设计工-还需'] = total_out_b4['设计工-预算'] - total_out_b4['设计工-核算']
    total_out_b4['费-还需'] = total_out_b4['制费-预算']+total_out_b4['设计工-预算'] - total_out_b4['费-核算']
    total_out_b4['料-还需']=0
    total_out_b4['生产工-还需'] = 0
    total_out_b4['其他费-还需'] = total_out_b4['制费-预算'] - total_out_b4['其他费-核算']
    total_out_b4['其他费-还需'] = total_out_b4['制费-预算'] - total_out_b4['其他费-核算']
    total_out_b4['工-还需'] = total_out_b4['交付工-还需']+total_out_b4['生产工-还需']
    total_out_b4['成本合计-还需'] = total_out_b4['工-还需'] + total_out_b4['费-还需']

    total_out_b3[['成本合计-还需', '料-还需', '工-还需', '生产工-还需', '交付工-还需', '费-还需', '设计工-还需', '其他费-还需']] =total_out_b3[['成本合计-还需', '料-还需', '工-还需', '生产工-还需', '交付工-还需', '费-还需', '设计工-还需', '其他费-还需']].fillna(0)
    for num in ['成本合计', '料', '工', '生产工', '交付工', '费', '设计工', '其他费']:
        total_out_b3[num+'-还需']=total_out_b3[num+'-还需']/total_out_b3['小项目数量']
    for num in ['成本合计', '交付工', '设计工']:
        total_out_b3.loc[~total_out_b3['项目号整理'].isin(need_east_group['项目号整理']),num+'-还需']=total_out_b3[num+'-预算']-total_out_b3[num+'-核算']
    total_out_b3.loc[~total_out_b3['项目号整理'].isin(need_east_group['项目号整理']), '其他费-还需'] = total_out_b3['制费-预算'] - \
                                                                                           total_out_b3['其他费-核算']
    ###已验收
    for num in ['成本合计-还需', '料-还需', '工-还需', '生产工-还需', '交付工-还需', '费-还需', '设计工-还需', '其他费-还需']:
        total_out_d[num]=0
    ###子项目
    total_out_c['成本合计-还需'] = -1*total_out_c['成本合计-核算']
    total_out_c['料-还需'] = -1 * total_out_c['料-核算']
    total_out_c['工-还需'] = -1 * total_out_c['工-核算']
    total_out_c['生产工-还需'] = -1 * total_out_c['生产工-核算']
    total_out_c['交付工-还需'] = -1 * total_out_c['交付工-核算']
    total_out_c['费-还需'] = -1 * total_out_c['费-核算']
    total_out_c['设计工-还需'] = -1 * total_out_c['设计工-核算']
    total_out_c['其他费-还需'] = -1 * total_out_c['其他费-核算']
    for num in ['成本合计-还需', '料-还需', '工-还需', '生产工-还需', '交付工-还需', '费-还需', '设计工-还需', '其他费-还需']:
        total_out_c.loc[total_out_c['子项目状态'].str.contains('验收'),num]=0
    ###未出货
    total_out_a['成本合计-还需']=total_out_a['成本合计-预算'] - total_out_a['成本合计-核算']
    total_out_a['料-还需'] = total_out_a['料-预算'] - total_out_a['料-核算']

    total_out_a['生产工-还需'] = total_out_a['生产工-预算'] - total_out_a['生产工-核算']
    total_out_a['交付工-还需'] = total_out_a['交付工-预算'] - total_out_a['交付工-核算']


    total_out_a['设计工-还需'] = total_out_a['设计工-预算'] - total_out_a['设计工-核算']
    total_out_a['其他费-还需'] = total_out_a['制费-预算'] - total_out_a['其他费-核算']


    total_out=pd.concat([total_out_a,total_out_b3,total_out_b4,total_out_c,total_out_d]).reset_index(drop=True)
    total_out =total_out.sort_values(by=['排个序'],ascending=True).reset_index(drop=True)
    for num in ['成本合计-还需', '料-还需', '工-还需', '生产工-还需', '交付工-还需', '费-还需', '设计工-还需', '其他费-还需']:
        total_out.loc[(total_out[num]<0)&(total_out['生产状态'].str.contains('子项目')==False),num]=0

    total_out['费-还需'] = total_out['其他费-还需'] + total_out['设计工-还需']
    total_out['工-还需'] = total_out['生产工-还需'] + total_out['交付工-还需']
    total_out['成本合计-还需'] = total_out['工-还需'] + total_out['费-还需'] + total_out['料-还需']


    print("3.9：拉取存货表..")
    '''
    存货分两批：在产和已出货，方案：将存货表和汇总表分别按在产和已出货分开执行
    '''
    inventory_total = inventory[['项目号/批号', '存货编码', '料', '工', '费', '合计','结存数量','存货大类（重分类后）']]
    total_out['排序'] = 0
    for i in range(len(total_out)):
        total_out.loc[i, '排序'] = i + 1
    total_out_first=total_out[total_out['生产状态'].str.contains('在产')].reset_index(drop=True)
    total_out_second = total_out[(total_out['生产状态'].str.contains('在产')==False)].reset_index(drop=True)

    inventory_use=inventory_total[(inventory_total['存货大类（重分类后）'].str.contains('发出')==False)&(inventory_total['存货大类（重分类后）'].str.contains('原材料|委托加工')==False)].reset_index(drop=True)
    inventory_use_other= inventory_total[inventory_total['存货大类（重分类后）'].str.contains('发出')].reset_index(drop=True)
    inventory_use_mater = inventory_total[inventory_total['存货大类（重分类后）'].str.contains('原材料|委托加工')].reset_index(drop=True)
####################在产
    if len(inventory_use[inventory_use['项目号/批号'].str.contains('-')]) > 0:
        inventory_use['项目号整理'] = ''
        inventory_use['项目整'] = inventory_use['项目号/批号'].str.split('-', expand=True)[0]
        inventory_use['项目整1'] = inventory_use['项目号/批号'].str.split('-', expand=True)[1]
        inventory_use['项目整1'] = inventory_use['项目整1'].fillna('空值')
        inventory_use['项目号整理'] = inventory_use['项目整']
        inventory_use.loc[(inventory_use['项目整1'].str.isdigit()) | (inventory_use['项目整1'].str.contains('SH')), '项目号整理'] = \
        inventory_use['项目号整理'] + '-' + inventory_use['项目整1']
    if len(inventory_use[inventory_use['项目号/批号'].str.contains('-')]) == 0:
        inventory_use['项目号整理'] = inventory_use['项目号/批号']
    inventory_use.loc[(inventory_use['项目号整理'].str[0] == 'F')&(inventory_use['项目号整理'].str[:3].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX')), '项目号整理'] =inventory_use['项目号整理'].str[3:]
    inventory_use.loc[(inventory_use['项目号整理'].str[:2].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX', na=False)), '项目号整理'] = inventory_use['项目号整理'].str[2:]

    inventory_use['项目号+料号'] = inventory_use['项目号整理'] + inventory_use['存货编码']
    ######存货取单价
    inventory_use1 = inventory_use.drop_duplicates(subset=['项目号+料号']).reset_index(drop=True)
    inventory_use_group1 = inventory_use.groupby(['项目号+料号']).agg(
        {'合计': "sum", '料': "sum", '工': "sum", '费': "sum",'结存数量':"sum"}).add_suffix('').reset_index()
    inventory_use_group1['项目号整理'] = pd.merge(inventory_use_group1, inventory_use1, on='项目号+料号', how='left')['项目号整理']
    inventory_number=[ '料', '工', '费', '合计']
    for num in inventory_number:
        inventory_use_group1.loc[inventory_use_group1['结存数量'] != 0, num] =inventory_use_group1[num] / inventory_use_group1['结存数量']

    '''
    方案：存货需靠两层（项目号整理+料号，项目号整理）去给到金额，按第一层关系给到汇总表，然后将汇总表一拆为2，并将存货表扣掉按第一层拉过去的金额，把剩余金额均摊给汇总表拆出来的2表，后将被拆过的汇总表合并还原
    '''
    total_out_first['项目号+料号'] = total_out_first['项目号整理'] + total_out_first['成品料号']
    total_out_first[['成本合计-存货', '料-存货', '工-存货', '费-存货']] = pd.merge(total_out_first, inventory_use_group1, on='项目号+料号', how='left')[
        ['料', '工', '费', '合计']]
    total_out_first1 = total_out_first[total_out_first['项目号+料号'].isin(inventory_use_group1['项目号+料号']) & (total_out_first['成品料号'] != '') & (
                total_out_first['成品料号'] != '/')].reset_index(drop=True)
    '''
    total_out1_group = total_out1.groupby(['项目号+料号']).agg({'核算项目号': "count"}).add_suffix('数量').reset_index()
    total_out1['项目号+料号数量'] = pd.merge(total_out1, total_out1_group, on='项目号+料号', how='left')['核算项目号数量']
    inventory_num = ['成本合计-存货', '料-存货', '工-存货', '费-存货']
    for num in inventory_num:
        total_out1[num] = total_out1[num] / total_out1['项目号+料号数量']
    '''
    ##第二层
    total_out_first2 = total_out_first[~total_out_first['序列号'].isin(total_out_first1['序列号'])].reset_index(drop=True)
    inventory_use_group2 = inventory_use_group1[~inventory_use_group1['项目号+料号'].isin(total_out_first1['项目号+料号'])].reset_index(
        drop=True)
    inventory_use_group3 = inventory_use_group2.groupby(['项目号整理']).agg(
        {'合计': "sum", '料': "sum", '工': "sum", '费': "sum"}).add_suffix('').reset_index()
    total_out_first2[['成本合计-存货', '料-存货', '工-存货', '费-存货']] = \
    pd.merge(total_out_first2, inventory_use_group3, on='项目号整理', how='left')[['料', '工', '费', '合计']]
    '''
    total_out2_group = total_out2.groupby(['项目号整理']).agg({'核算项目号': "count"}).add_suffix('数量').reset_index()
    total_out2['项目号整理数量'] = pd.merge(total_out2, total_out2_group, on='项目号整理', how='left')['核算项目号数量']
    for num in inventory_num:
        total_out2[num] = total_out2[num] / total_out2['项目号整理数量']
    '''
    summary1 = pd.concat([total_out_first1, total_out_first2]).reset_index(drop=True)
####################出货
    if len(inventory_use_other[inventory_use_other['项目号/批号'].str.contains('-')]) > 0:
        inventory_use_other['项目号整理'] = ''
        inventory_use_other['项目整'] = inventory_use_other['项目号/批号'].str.split('-', expand=True)[0]
        inventory_use_other['项目整1'] = inventory_use_other['项目号/批号'].str.split('-', expand=True)[1]
        inventory_use_other['项目整1'] = inventory_use_other['项目整1'].fillna('空值')
        inventory_use_other['项目号整理'] = inventory_use_other['项目整']
        inventory_use_other.loc[(inventory_use_other['项目整1'].str.isdigit()) | (inventory_use_other['项目整1'].str.contains('SH')), '项目号整理'] = \
        inventory_use_other['项目号整理'] + '-' + inventory_use_other['项目整1']
    if len(inventory_use_other[inventory_use_other['项目号/批号'].str.contains('-')]) == 0:
        inventory_use_other['项目号整理'] = inventory_use_other['项目号/批号']
    inventory_use_other.loc[(inventory_use_other['项目号整理'].str[0] == 'F')&(inventory_use_other['项目号整理'].str[:3].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX')), '项目号整理'] =inventory_use_other['项目号整理'].str[3:]
    inventory_use_other.loc[(inventory_use_other['项目号整理'].str[:2].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX', na=False)), '项目号整理'] = inventory_use_other['项目号整理'].str[2:]

    inventory_use_other['项目号+料号'] = inventory_use_other['项目号整理'] + inventory_use_other['存货编码']
    ######存货取单价
    inventory_use_other1 = inventory_use_other.drop_duplicates(subset=['项目号+料号']).reset_index(drop=True)
    inventory_use_other_group1 = inventory_use_other.groupby(['项目号+料号']).agg(
        {'合计': "sum", '料': "sum", '工': "sum", '费': "sum",'结存数量':"sum"}).add_suffix('').reset_index()
    inventory_use_other_group1['项目号整理'] = pd.merge(inventory_use_other_group1, inventory_use_other1, on='项目号+料号', how='left')['项目号整理']
    inventory_number=[ '料', '工', '费', '合计']
    for num in inventory_number:
        inventory_use_other_group1.loc[inventory_use_other_group1['结存数量'] != 0, num] =inventory_use_other_group1[num] / inventory_use_other_group1['结存数量']

    '''
    方案：存货需靠两层（项目号整理+料号，项目号整理）去给到金额，按第一层关系给到汇总表，然后将汇总表一拆为2，并将存货表扣掉按第一层拉过去的金额，把剩余金额均摊给汇总表拆出来的2表，后将被拆过的汇总表合并还原
    '''
    total_out_second['项目号+料号'] = total_out_second['项目号整理'] + total_out_second['成品料号']
    total_out_second[['成本合计-存货', '料-存货', '工-存货', '费-存货']] = pd.merge(total_out_second, inventory_use_other_group1, on='项目号+料号', how='left')[
        ['料', '工', '费', '合计']]
    total_out_second1 = total_out_second[total_out_second['项目号+料号'].isin(inventory_use_other_group1['项目号+料号']) & (total_out_second['成品料号'] != '') & (
                total_out_second['成品料号'] != '/')].reset_index(drop=True)
    '''
    total_out1_group = total_out1.groupby(['项目号+料号']).agg({'核算项目号': "count"}).add_suffix('数量').reset_index()
    total_out1['项目号+料号数量'] = pd.merge(total_out1, total_out1_group, on='项目号+料号', how='left')['核算项目号数量']
    inventory_num = ['成本合计-存货', '料-存货', '工-存货', '费-存货']
    for num in inventory_num:
        total_out1[num] = total_out1[num] / total_out1['项目号+料号数量']
    '''
    ##第二层
    total_out_second2 = total_out_second[~total_out_second['序列号'].isin(total_out_second1['序列号'])].reset_index(drop=True)
    inventory_use_other_group2 = inventory_use_other_group1[~inventory_use_other_group1['项目号+料号'].isin(total_out_second1['项目号+料号'])].reset_index(
        drop=True)
    inventory_use_other_group3 = inventory_use_other_group2.groupby(['项目号整理']).agg(
        {'合计': "sum", '料': "sum", '工': "sum", '费': "sum"}).add_suffix('').reset_index()
    total_out_second2[['成本合计-存货', '料-存货', '工-存货', '费-存货']] = \
    pd.merge(total_out_second2, inventory_use_other_group3, on='项目号整理', how='left')[['料', '工', '费', '合计']]
    '''
    total_out2_group = total_out2.groupby(['项目号整理']).agg({'核算项目号': "count"}).add_suffix('数量').reset_index()
    total_out2['项目号整理数量'] = pd.merge(total_out2, total_out2_group, on='项目号整理', how='left')['核算项目号数量']
    for num in inventory_num:
        total_out2[num] = total_out2[num] / total_out2['项目号整理数量']
    '''
    summary2 = pd.concat([total_out_second1, total_out_second2]).reset_index(drop=True)
#############合并
    summary= pd.concat([summary1,summary2]).reset_index(drop=True)
    summary = summary.sort_values(by=['排序'], ascending=[True]).reset_index(drop=True)

#######原材料
    inventory_mater=inventory_use_mater[['项目号/批号','合计','结存数量','存货编码']]
    if len(inventory_mater[inventory_mater['项目号/批号'].str.contains('-')]) > 0:
        inventory_mater['项目号整理'] = ''
        inventory_mater['项目整'] = inventory_mater['项目号/批号'].str.split('-', expand=True)[0]
        inventory_mater['项目整1'] = inventory_mater['项目号/批号'].str.split('-', expand=True)[1]
        inventory_mater['项目整1'] = inventory_mater['项目整1'].fillna('空值')
        inventory_mater['项目号整理'] = inventory_mater['项目整']
        inventory_mater.loc[
            (inventory_mater['项目整1'].str.isdigit()) | (inventory_mater['项目整1'].str.contains('SH')), '项目号整理'] = \
            inventory_mater['项目号整理'] + '-' + inventory_mater['项目整1']
    if len(inventory_mater[inventory_mater['项目号/批号'].str.contains('-')]) == 0:
        inventory_mater['项目号整理'] = inventory_mater['项目号/批号']
    inventory_mater.loc[(inventory_mater['项目号整理'].str[0] == 'F') & (
        inventory_mater['项目号整理'].str[:3].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX')), '项目号整理'] = \
    inventory_mater['项目号整理'].str[3:]
    inventory_mater.loc[
        (inventory_mater['项目号整理'].str[:2].str.contains('JM|JS|SZ|jm|jM|Jm|js|Js|jS|Sz|sz|sZ|HX', na=False)), '项目号整理'] = \
    inventory_mater['项目号整理'].str[2:]

    inventory_mater['项目号+料号'] = inventory_mater['项目号整理'] + inventory_mater['存货编码']
    ######存货取单价
    inventory_mater1 = inventory_mater.drop_duplicates(subset=['项目号+料号']).reset_index(drop=True)
    inventory_mater_group1 = inventory_mater.groupby(['项目号+料号']).agg(
        {'合计': "sum",'结存数量': "sum"}).add_suffix('').reset_index()
    inventory_mater_group1['项目号整理'] = pd.merge(inventory_mater_group1, inventory_mater1, on='项目号+料号', how='left')['项目号整理']
    inventory_mater_group1.loc[inventory_mater_group1['结存数量'] != 0,'合计'] = inventory_mater_group1['合计']/inventory_mater_group1['结存数量']

    summary['项目号+料号'] = summary['项目号整理'] + summary['成品料号']
    summary[['原材料-存货']] = pd.merge(summary, inventory_mater_group1, on='项目号+料号', how='left')[['合计']]
    summary1 = summary[summary['项目号+料号'].isin(inventory_mater_group1['项目号+料号']) & (summary['成品料号'] != '') & (summary['成品料号'] != '/')].reset_index(drop=True)

    ##第二层
    summary2 = summary[~summary['序列号'].isin(summary1['序列号'])].reset_index(drop=True)
    inventory_mater_group2 = inventory_mater_group1[~inventory_mater_group1['项目号+料号'].isin(summary1['项目号+料号'])].reset_index(
        drop=True)
    inventory_mater_group3 = inventory_mater_group2.groupby(['项目号整理']).agg(
        {'合计': "sum"}).add_suffix('').reset_index()
    summary2[['原材料-存货']] =pd.merge(summary2, inventory_mater_group3, on='项目号整理', how='left')[['合计']]
    summary = pd.concat([summary1,summary2]).reset_index(drop=True)
    summary = summary.sort_values(by=['排个序'], ascending=[True]).reset_index(drop=True)
    time_data_catch=time.time()
    print('第三阶段【数据拉取】执行时长:%d秒' % (time_data_catch - time_data_format))

    ######################################################################################################################################四
    print("四、汇总表字段加工..")
    summary_num=['成本合计-概算', '料-概算','生产工-概算', '交付工-概算', '设计工-概算',   '制费-概算'
        , '成本合计-预算', '料-预算', '生产工-预算', '交付工-预算', '设计工-预算',  '制费-预算'
        ,'成本合计-财务', '料-财务', '工-财务', '费-财务'
        ,'采购PO','原材料-存货'
        , '成本合计-核算', '料-核算','工单料','设变料', '工-核算', '生产工-核算', '交付工-核算', '费-核算', '设计工-核算', '其他费-核算', '毛利-核算', '毛利率-核算'
        , '成本合计-年初', '料-年初', '工-年初','生产工-年初', '交付工-年初', '费-年初', '设计工-年初', '其他费-年初'
        , '成本合计-还需', '料-还需','工-还需', '生产工-还需', '交付工-还需', '费-还需', '设计工-还需', '其他费-还需'
        ,'成本合计-存货', '料-存货','工-存货',  '费-存货']
    summary[summary_num]=summary[summary_num].fillna('')

    print("4.1：加工出货&验收年月..")
    summary['出货年份'] = ''
    summary['出货月份'] = ''
    summary[['计划出货时间', '产品线计划验收时间']] = summary[['计划出货时间', '产品线计划验收时间']].fillna(default_date)
    summary.loc[summary['产品线计划验收时间'] == '','产品线计划验收时间'] = default_date
    summary.loc[summary['计划出货时间'] == '', '计划出货时间'] = default_date
    summary['计划出货时间']=pd.to_datetime(summary['计划出货时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
    summary['产品线计划验收时间'] = pd.to_datetime(summary['产品线计划验收时间'], errors='coerce').dt.strftime('%Y-%m-%d').astype(str)
    summary.loc[summary['计划出货时间'].str.contains('1990') == False, '出货年份'] = pd.to_datetime(summary['计划出货时间'],errors='coerce').dt.strftime('%Y').astype(int)
    summary.loc[summary['计划出货时间'].str.contains('1990') == False, '出货月份'] = pd.to_datetime(summary['计划出货时间'],errors='coerce').dt.strftime('%m').astype(int)
    summary['验收年份'] = ''
    summary['验收月份'] = ''
    summary.loc[summary['产品线计划验收时间'].str.contains('1990') == False, '验收年份'] = pd.to_datetime(summary['产品线计划验收时间'],errors='coerce').dt.strftime('%Y').astype(int)
    summary.loc[summary['产品线计划验收时间'].str.contains('1990') == False, '验收月份'] = pd.to_datetime(summary['产品线计划验收时间'],errors='coerce').dt.strftime('%m').astype(int)
    summary['产品线计划验收时间'] = ['' if i == '1990-01-01' else i for i in summary['产品线计划验收时间']]
    summary['计划出货时间'] = ['' if i == '1990-01-01' else i for i in summary['计划出货时间']]

    print("4.2：滚动预测..")##核算+还需
    summary['成本合计-滚动']=''
    summary['料-滚动'] = ''
    summary['工-滚动'] = ''
    summary['生产工-滚动'] = ''
    summary['交付工-滚动'] = ''
    summary['费-滚动'] = ''
    summary['设计工-滚动'] = ''
    summary['其他费-滚动'] = ''
    '''
    summary.loc[(summary['成本合计-核算']!='')&(summary['成本合计-还需']!=''),'成本合计-滚动']=summary['成本合计-核算'] + summary['成本合计-还需']
    summary.loc[(summary['料-核算']!='')&(summary['料-还需']!=''),'料-滚动'] = summary['料-核算'] + summary['料-还需']
    summary.loc[(summary['工-核算']!='')&(summary['工-还需']!=''),'工-滚动'] = summary['工-核算'] + summary['工-还需']
    summary.loc[(summary['生产工-核算']!='')&(summary['生产工-还需']!=''),'生产工-滚动'] = summary['生产工-核算'] + summary['生产工-还需']
    summary.loc[(summary['交付工-核算']!='')&(summary['交付工-还需']!=''),'交付工-滚动'] = summary['交付工-核算'] + summary['交付工-还需']
    summary.loc[(summary['费-核算']!='')&(summary['费-还需']!=''),'费-滚动'] = summary['费-核算'] + summary['费-还需']
    summary.loc[(summary['设计工-核算']!='')&(summary['设计工-还需']!=''),'设计工-滚动'] = summary['设计工-核算'] + summary['设计工-还需']
    summary.loc[(summary['其他费-核算']!='')&(summary['其他费-还需']!=''),'其他费-滚动'] = summary['其他费-核算'] + summary['其他费-还需']
    '''
    for i in range(len(summary)):
        if summary.loc[i,"成本合计-核算"]!='' and summary.loc[i,"成本合计-还需"]!='':
            summary.loc[i, "成本合计-滚动"]=summary.loc[i,'成本合计-核算'] + summary.loc[i,'成本合计-还需']
        if summary.loc[i,"料-核算"]!='' and summary.loc[i,"料-还需"]!='':
            summary.loc[i, "料-滚动"]=summary.loc[i,'料-核算'] + summary.loc[i,'料-还需']
        if summary.loc[i, "工-核算"] != '' and summary.loc[i, "工-还需"] != '':
            summary.loc[i, "工-滚动"] = summary.loc[i, '工-核算'] + summary.loc[i, '工-还需']
        if summary.loc[i, "生产工-核算"] != '' and summary.loc[i, "生产工-还需"] != '':
            summary.loc[i, "生产工-滚动"] = summary.loc[i, '生产工-核算'] + summary.loc[i, '生产工-还需']
        if summary.loc[i, "交付工-核算"] != '' and summary.loc[i, "交付工-还需"] != '':
            summary.loc[i, "交付工-滚动"] = summary.loc[i, '交付工-核算'] + summary.loc[i, '交付工-还需']
        if summary.loc[i, "费-核算"] != '' and summary.loc[i, "费-还需"] != '':
            summary.loc[i, "费-滚动"] = summary.loc[i, '费-核算'] + summary.loc[i, '费-还需']
        if summary.loc[i, "设计工-核算"] != '' and summary.loc[i, "设计工-还需"] != '':
            summary.loc[i, "设计工-滚动"] = summary.loc[i, '设计工-核算'] + summary.loc[i, '设计工-还需']
        if summary.loc[i, "其他费-核算"] != '' and summary.loc[i, "其他费-还需"] != '':
            summary.loc[i, "其他费-滚动"] = summary.loc[i, '其他费-核算'] + summary.loc[i, '其他费-还需']

    print("4.3：毛利&毛利率..")
    summary['毛利-概算']=''
    summary['毛利率-概算']=''
    for i in range(len(summary)):
        if summary.loc[i,"集团收入"]!='' and summary.loc[i,"成本合计-概算"]!='':
            summary.loc[i,"毛利-概算"]=summary.loc[i,'集团收入']-summary.loc[i,'成本合计-概算']
    for i in range(len(summary)):
        if summary.loc[i,"集团收入"]!='' and summary.loc[i,"集团收入"]!=0 and summary.loc[i,"毛利-概算"]!='':
            summary.loc[i,"毛利率-概算"]=summary.loc[i,'毛利-概算'] / summary.loc[i,'集团收入']

    summary['毛利-预算'] = ''
    summary['毛利率-预算'] = ''
    for i in range(len(summary)):
        if summary.loc[i, "集团收入"] != '' and summary.loc[i, "成本合计-预算"] != '':
            summary.loc[i, "毛利-预算"] = summary.loc[i, '集团收入']-summary.loc[i, '成本合计-预算']
    for i in range(len(summary)):
        if summary.loc[i, "集团收入"] != '' and summary.loc[i, "集团收入"] != 0 and summary.loc[i, "毛利-预算"] != '':
            summary.loc[i, "毛利率-预算"] = summary.loc[i, '毛利-预算'] / summary.loc[i, '集团收入']

    summary['毛利-核算'] = ''
    summary['毛利率-核算'] = ''
    for i in range(len(summary)):
        if summary.loc[i, "集团收入"] != '' and summary.loc[i, "成本合计-核算"] != '':
            summary.loc[i, "毛利-核算"] =summary.loc[i, '集团收入']-summary.loc[i, '成本合计-核算']
    for i in range(len(summary)):
        if summary.loc[i, "集团收入"] != '' and summary.loc[i, "集团收入"] != 0 and summary.loc[i, "毛利-核算"] != '':
            summary.loc[i, "毛利率-核算"] = summary.loc[i, '毛利-核算'] / summary.loc[i, '集团收入']

    summary['毛利-滚动'] = ''
    summary['毛利率-滚动'] = ''
    for i in range(len(summary)):
        if summary.loc[i, "集团收入"] != '' and summary.loc[i, "成本合计-滚动"] != '':
            summary.loc[i, "毛利-滚动"] = summary.loc[i, '集团收入']-summary.loc[i, '成本合计-滚动']
    for i in range(len(summary)):
        if summary.loc[i, "集团收入"] != '' and summary.loc[i, "集团收入"] != 0 and summary.loc[i, "毛利-滚动"] != '':
            summary.loc[i, "毛利率-滚动"] = summary.loc[i, '毛利-滚动'] / summary.loc[i, '集团收入']

    summary['毛利-财务'] = ''
    summary['毛利率-财务'] = ''
    for i in range(len(summary)):
        if summary.loc[i, "集团收入"] != '' and summary.loc[i, "成本合计-财务"] != '':
            summary.loc[i, "毛利-财务"] = summary.loc[i, '集团收入']-summary.loc[i, '成本合计-财务']
    for i in range(len(summary)):
        if summary.loc[i, "集团收入"] != '' and summary.loc[i, "集团收入"] != 0 and summary.loc[i, "毛利-财务"] != '':
            summary.loc[i, "毛利率-财务"] = summary.loc[i, '毛利-财务'] / summary.loc[i, '集团收入']

    print("4.3：整理格式..")
    summary['事业部收入']=''
    summary['大项目整理'] = ''
    summary=summary.reindex(columns=[
    '序列号', '区域', '行业中心', '客户简称', '大项目名称', '大项目号', '产品线名称', '核算项目号', '设备名称','项目数量', '已出货数量', '在产数量', '生产状态', '全面预算有无', '设备类型'
    , '项目经理','项目阶段','出货年份','出货月份', '验收年份', '验收月份'
    , '事业部收入','集团收入', '软件收入','硬件收入'
    , '成本合计-概算', '料-概算','生产工-概算', '交付工-概算', '设计工-概算', '制费-概算', '毛利-概算', '毛利率-概算'
    , '成本合计-预算','料-预算', '生产工-预算', '交付工-预算', '设计工-预算', '制费-预算', '毛利-预算', '毛利率-预算'
    ,'采购PO','原材料-存货'
    , '成本合计-核算', '料-核算','工单料','设变料', '工-核算', '生产工-核算','交付工-核算', '费-核算', '设计工-核算', '其他费-核算', '毛利-核算', '毛利率-核算'
    , '成本合计-还需', '料-还需','工-还需', '生产工-还需', '交付工-还需', '费-还需', '设计工-还需', '其他费-还需'
    , '成本合计-滚动', '料-滚动', '工-滚动', '生产工-滚动', '交付工-滚动','费-滚动', '设计工-滚动', '其他费-滚动','毛利-滚动', '毛利率-滚动'
    , '成本合计-财务', '料-财务', '工-财务', '费-财务','毛利-财务', '毛利率-财务'
    , '成本合计-存货', '料-存货', '工-存货', '费-存货'
    , '项目号整理','是否预验未终验','子项目状态', '成品料号', '一般工单号601/608', '返工工单号603'
    , '大项目整理'
    , '实际出货时间', '实际验收时间','系统验收时间', '是否预验收'
    , '是否有风险', '风险分类', '生产实际进度', '验收实际进度'
    , '成本合计-年初', '料-年初', '工-年初','生产工-年初', '交付工-年初', '费-年初', '设计工-年初', '其他费-年初'])
    summary_date=['实际出货时间', '实际验收时间','系统验收时间']
    for dat in summary_date:
        summary[dat]=['' if i == '1990-01-01' else i for i in summary[dat]]
    time_data_refresh=time.time()
    print('第三阶段【数据拉取】执行时长:%d秒' % (time_data_refresh - time_data_catch))

    ######################################################################################################################################五
    print('五、表格输出...')
    def writer_contents(sheet, array, start_row, start_col, format=None, percent_format=None, percentlist=[]):
        # start_col = 0
        for col in array:
            if percentlist and (start_col in percentlist):
                sheet.write_column(start_row, start_col, col, percent_format)
            else:
                sheet.write_column(start_row, start_col, col, format)
            start_col += 1
    def write_color(book, sheet, data, fmt, col_num='I'):
        start = 3
        format_red = book.add_format({'font_name': 'Arial',
                                    'font_size': 10,
                                    'bg_color':'#F86470'})
        format_red.set_align('center')
        format_red.set_align('vcenter')

        for item in data:
            if '找不到' in str(item)  in str(item):
                sheet.write(col_num + str(start), item, format_red)
            else:
                sheet.write(col_num + str(start), item, fmt)
            start += 1

    # 写入表格
    print('5.1：正在设置表格格式...')
    now_time = time.strftime("%Y-%m-%d-%H", time.localtime(time.time()))
    book_name = '概预核决汇总表'+ now_time
    workbook = xlsxwriter.Workbook(book_name + '.xlsx', {'nan_inf_to_errors': True})
    worksheet0 = workbook.add_worksheet('汇总表')
    worksheet1 = workbook.add_worksheet('01-底表基础数据')##base
    worksheet2 = workbook.add_worksheet('02-概预算')#budget_estimate
    worksheet3 = workbook.add_worksheet('03-核算表')##calcu
    worksheet4 = workbook.add_worksheet('04-财务成本')##finance
    worksheet5 = workbook.add_worksheet('05-周进度')##advance
    worksheet6 = workbook.add_worksheet('06-还需')##need
    worksheet7 = workbook.add_worksheet('07-存货')##inventory
    worksheet8= workbook.add_worksheet('08-年初预算表')  ##year_budget
    ######主色调
    title_format = workbook.add_format({'font_name': 'Arial',
                                            'font_size': 10,
                                            'font_color':'white',
                                            'bg_color':'#1F4E78',
                                            'bold': True,
                                            'align':'center',
                                            'valign':'vcenter',
                                            'border':1,
                                            'border_color':'white'
                                            })
    title_format.set_align('vcenter')
    ######项目&数量  |存货
    title_format1 = workbook.add_format({'font_name': 'Arial',
                                        'font_size': 10,
                                        'font_color': 'white',
                                        'bg_color': '#974706',
                                        'bold': True,
                                        'align': 'center',
                                        'valign': 'vcenter',
                                        'border': 1,
                                        'border_color': 'white'
                                        })
    title_format1.set_align('vcenter')
    ######出货&验收
    title_format2= workbook.add_format({'font_name': 'Arial',
                                         'font_size': 10,
                                         'font_color': 'white',
                                         'bg_color': '#215967',
                                         'bold': True,
                                         'align': 'center',
                                         'valign': 'vcenter',
                                         'border': 1,
                                         'border_color': 'white'
                                         })
    title_format2.set_align('vcenter')
    ######收入\滚动预测
    title_format3 = workbook.add_format({'font_name': 'Arial',
                                         'font_size': 10,
                                         'font_color': 'white',
                                         'bg_color': '#7030A0',
                                         'bold': True,
                                         'align': 'center',
                                         'valign': 'vcenter',
                                         'border': 1,
                                         'border_color': 'white'
                                         })
    title_format3.set_align('vcenter')
    ######预算
    title_format4 = workbook.add_format({'font_name': 'Arial',
                                         'font_size': 10,
                                         'font_color': 'white',
                                         'bg_color': '#BF6753',
                                         'bold': True,
                                         'align': 'center',
                                         'valign': 'vcenter',
                                         'border': 1,
                                         'border_color': 'white'
                                         })
    title_format4.set_align('vcenter')
    ######核算
    title_format5 = workbook.add_format({'font_name': 'Arial',
                                         'font_size': 10,
                                         'font_color': 'white',
                                         'bg_color': '#00B050',
                                         'bold': True,
                                         'align': 'center',
                                         'valign': 'vcenter',
                                         'border': 1,
                                         'border_color': 'white'
                                         })
    title_format5.set_align('vcenter')
    ######还需
    title_format6= workbook.add_format({'font_name': 'Arial',
                                         'font_size': 10,
                                         'font_color': 'white',
                                         'bg_color': '#E26B0A',
                                         'bold': True,
                                         'align': 'center',
                                         'valign': 'vcenter',
                                         'border': 1,
                                         'border_color': 'white'
                                         })
    title_format6.set_align('vcenter')
    ######财务|年初预算
    title_format7 = workbook.add_format({'font_name': 'Arial',
                                         'font_size': 10,
                                         'font_color': 'white',
                                         'bg_color': '#FF0000',
                                         'bold': True,
                                         'align': 'center',
                                         'valign': 'vcenter',
                                         'border': 1,
                                         'border_color': 'white'
                                         })
    title_format7.set_align('vcenter')
    ######项目进度
    title_format8 = workbook.add_format({'font_name': 'Arial',
                                         'font_size': 10,
                                         'font_color': 'white',
                                         'bg_color': '#00B0F0',
                                         'bold': True,
                                         'align': 'center',
                                         'valign': 'vcenter',
                                         'border': 1,
                                         'border_color': 'white'
                                         })
    title_format8.set_align('vcenter')
    col_format = workbook.add_format({'font_name': 'Arial',
                                        'font_size': 8,
                                        'font_color':'white',
                                        'bg_color':'#595959',
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
    data_format1 = workbook.add_format({'font_name': 'Arial',
                                        'font_size': 10,
                                        'align':'center',
                                        'valign':'vcenter'
                                        })
    data_format2 = workbook.add_format({'font_name': 'Arial',
                                        'font_size': 10,
                                        'align':'center',
                                        'valign':'vcenter'
                                        })
    data_format2.set_num_format('0.00')
    data_format3 = workbook.add_format({'font_name': 'Arial',
                                        'font_size': 10,
                                        'align':'center',
                                        'valign':'vcenter'
                                        })
    data_format3.set_num_format('0.00%')
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
    data_format_percent = workbook.add_format({'font_name': 'Arial',
                                               'font_size': 10,
                                               'align': 'center',
                                               'valign': 'vcenter'
                                               })
    data_format_percent.set_num_format('0.00%')

    print('5.2：正在写入EXCEL表格...')
    worksheet0.write_row("A2", summary.columns, title_format)
    writer_contents(sheet=worksheet0, array=summary.T.values, start_row=2,start_col=0)
   # end = len(report_work1) + 1
    worksheet0.merge_range('A1:Q1', '项目基础信息', title_format)
    worksheet0.merge_range('R1:U1', '计划出货/验收', title_format2)
    worksheet0.merge_range('V1:Y1', '收入数据', title_format3)
    worksheet0.merge_range('Z1:AG1', '概算数据', title_format)
    worksheet0.merge_range('AH1:AO1', '预算数据', title_format4)
    worksheet0.merge_range('AP1:BC1', '核算数据', title_format5)
    worksheet0.merge_range('BD1:BK1', '还需数据', title_format6)
    worksheet0.merge_range('BL1:BU1', '滚动预测数据', title_format3)
    worksheet0.merge_range('BV1:CA1', '财务成本数据', title_format7)
    worksheet0.merge_range('CB1:CE1', '存货数据', title_format1)
    worksheet0.merge_range('CF1:CL1', '辅助信息', title_format)
    worksheet0.merge_range('CM1:CT1', '项目进度信息', title_format8)
    worksheet0.merge_range('CU1:DB1', '年初预算数据', title_format7)

    worksheet0.write_row("J2:L2", ['项目数量','已出货数量','在产数量'], title_format1)
    worksheet0.write_row("R2:U2", ['出货年份','出货月份','验收年份','验收月份'], title_format2)
    worksheet0.write_row("V2:Y2", ['事业部收入','集团收入','软件收入','硬件收入'], title_format3)
    worksheet0.write_row("Z2:AG2", ['成本合计', '料', '生产工', '交付工', '设计工', '其他费', '毛利', '毛利率'], title_format)
    worksheet0.write_row("AH2:AO2", ['成本合计','料','生产工','交付工','设计工','其他费','毛利','毛利率'], title_format4)
    worksheet0.write_row("AP2:BC2", ['采购PO','原材料-存货','成本合计', '料','工单料','设变料','工', '生产工', '交付工','费','设计工', '其他费', '毛利', '毛利率'], title_format5)
    worksheet0.write_row("BD2:BK2", ['成本合计', '料', '工', '生产工', '交付工','费', '设计工', '其他费'], title_format6)
    worksheet0.write_row("BL2:BU2", ['成本合计', '料','工', '生产工', '交付工','费','设计工', '其他费', '毛利', '毛利率'], title_format3)
    worksheet0.write_row("BV2:CA2", ['成本合计', '料', '工',  '费','毛利', '毛利率'], title_format7)
    worksheet0.write_row("CB2:CE2", ['成本合计', '料', '工',  '费'], title_format1)
    worksheet0.write_row("CM2:CT2", ['实际出货时间','实际验收时间','系统验收时间','是否预验收','是否有风险','风险分类','生产实际进度','验收实际进度'], title_format8)
    worksheet0.write_row("CU2:DB2",['成本合计', '料', '工', '生产工', '交付工','费', '设计工', '其他费'],title_format7)
    worksheet0.set_row(0, 25)
    worksheet0.set_row(1, 22)
    worksheet0.set_column('AG:AG', 8, data_format_percent)
    worksheet0.set_column('AO:AO', 8, data_format_percent)
    worksheet0.set_column('BC:BC', 8, data_format_percent)
    worksheet0.set_column('BU:BU', 8, data_format_percent)
    worksheet0.set_column('CA:CA', 8, data_format_percent)

    worksheet1.write_row("A1", base.columns, title_format)
    writer_contents(sheet=worksheet1, array=base.T.values, start_row=1, start_col=0)
    worksheet1.set_row(0, 25)

    worksheet2.write_row("A1", budget_estimate.columns, title_format)
    writer_contents(sheet=worksheet2, array=budget_estimate.T.values, start_row=1, start_col=0)
    worksheet2.set_row(0, 25)

    worksheet3.write_row("A1", calcu.columns, title_format)
    writer_contents(sheet=worksheet3, array=calcu.T.values, start_row=1, start_col=0)
    worksheet3.set_row(0, 25)

    worksheet4.write_row("A1", finance.columns, title_format)
    writer_contents(sheet=worksheet4, array=finance.T.values, start_row=1, start_col=0)
    worksheet4.set_row(0, 25)

    worksheet5.write_row("A1", advance.columns, title_format)
    writer_contents(sheet=worksheet5, array=advance.T.values, start_row=1, start_col=0)
    worksheet5.set_row(0, 25)

    worksheet6.write_row("A1", need.columns, title_format)
    writer_contents(sheet=worksheet6, array=need.T.values, start_row=1, start_col=0)
    worksheet6.set_row(0, 25)

    worksheet7.write_row("A1", inventory.columns, title_format)
    writer_contents(sheet=worksheet7, array=inventory.T.values, start_row=1, start_col=0)
    worksheet7.set_row(0, 25)

    worksheet8.write_row("A1", year_budget.columns, title_format)
    writer_contents(sheet=worksheet8, array=year_budget.T.values, start_row=1, start_col=0)
    worksheet8.set_row(0, 25)

    print('明细表已写入。。。')
    workbook.close()
    time_excel = time.time()
    print('五阶段【输出表格】执行时长:%d秒' % (time_excel - time_data_refresh))

except Exception as f:
    # print('异常信息为:', e)  # 异常信息为: division by zero
    print('——#@*&程序报错，异常信息为:' + traceback.format_exc())


time_end=time.time()
print('执行完成！！！！！')
print('执行总时长:%d秒' % (time_end - time_start))
input('*************&&&请点击去关闭程序...')
















