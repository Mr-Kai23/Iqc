import pandas as pd
import time
import numpy as np
import math
import re

# 设置控制台的输出选项
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 5000)


# 信息归一
def onlyunique(df_form, m):
    heads_name = df_form.columns.tolist()
    sheet = df_form[heads_name[m]]
    unique = sheet.unique()
    if len(unique) == 0 or (type(unique[0]) == np.float64 and math.isnan(unique[0])) or unique[0] is np.nan:
        return 'unknow'
    else:
        return unique[0]


# 判定是否为有效数据项
def is_valid_date(strdate):
    if type(strdate) == pd._libs.tslibs.timestamps.Timestamp:
        return True
    elif type(strdate) == str:
        try:
            if ":" in strdate:
                time.strptime(strdate, "%Y-%m-%d %H:%M%S")
            elif "/" in strdate:
                pass
            else:
                time.strptime(strdate, "%Y-%m-%d")
            return True
        except:
            return False
    else:
        return False


# 日报合成
def merge_excel(excel_list, report_type):
    result = []
    sum_excel_message = []

    for excel_src in excel_list:
        content = pd.read_excel(excel_src, headers=[0, 1])
        df = pd.DataFrame(data=content)
        single_excel_message = []

        for row in df.iterrows():
            single_excel_row = []

            for i in row[1][1:]:
                single_excel_row.append(i)
            flag = is_valid_date(strdate=single_excel_row[0])

            if flag:
                if type(single_excel_row[0]) == pd._libs.tslibs.timestamps.Timestamp:
                    # print("转换", str(single_excel_row[1].date()), type(single_excel_row[1]))
                    single_excel_row[0] = str(single_excel_row[0].date())
                single_excel_message.append(single_excel_row)
        sum_excel_message = sum_excel_message+single_excel_message

        # 进行if判断选择哪种处理方式
        if report_type == 0:
            result = merge_top_data(sum_excel_message)
        elif report_type == 1:
            result = merge_other_data(sum_excel_message, True)
        else:
            result = merge_other_data(sum_excel_message, False)

    return result


def timezone(input_time):
    min_time = np.min(input_time)
    max_time = np.max(input_time)
    if max_time == min_time:
        return max_time
    else:
        return min_time+'~'+max_time


def shiftzone(input_message, m):
    heads_name = input_message.columns.tolist()
    sheet = input_message[heads_name[m]]
    unique = sheet.unique()
    unique = ','.join(unique)
    return unique


def errors_regroup(input_message, m):

    realSeries = input_message[input_message[m].isna() == False][m]
    nameList = []

    for value in realSeries:
        nameList += re.findall('(.+):\t(.+)x', value)

    nameFrame = pd.DataFrame(nameList, columns=['error', 'amount'])
    nameFrame['amount'] = nameFrame['amount'].astype('int64')
    sumFrame = nameFrame.groupby(['error']).sum()

    errors = ''
    for i, j in sumFrame.iterrows():
        errors += (i + ':\t' + str(j['amount']) + 'x\n')

    return errors


def zeroexcerpt(molecule, denom):
    if denom == 0:
        ngrate = 0
    else:
        ngrate = format(molecule/denom, '.0%')
    return ngrate


# top信息汇总
def merge_top_data(datalist):
    convert_data = pd.DataFrame(datalist)

    merge_lists = []
    data = convert_data.groupby([4, 12, 13])

    for i, j in data:

        child_list = j.values.tolist()
        data_group = data.get_group(i)
        sheets = j
        insp_date = timezone(data_group[0])
        shift = shiftzone(sheets, 1)
        project = onlyunique(sheets, 2)
        stage = onlyunique(sheets, 3)
        pn = onlyunique(sheets, 4)
        color = onlyunique(sheets, 5)
        m_type = 'Top Module'
        mac = onlyunique(sheets, 7)
        iar = onlyunique(sheets, 8)
        datecode = onlyunique(sheets, 9)
        ecode = onlyunique(sheets, 10)
        vendor = onlyunique(sheets, 11)
        cfg = onlyunique(sheets, 12)
        batch = onlyunique(sheets, 13)
        total_isp_qty = data_group[14].agg(np.sum)
        total_pass_qty = data_group[15].agg(np.sum)
        total_ng_qty = data_group[16].agg(np.sum)

        total_ng_rate = zeroexcerpt(total_ng_qty, total_isp_qty)
        # 计算cosmetic1信息
        cos1_isp_qty = data_group[18].agg(np.sum)
        cos1_pass_qty = data_group[19].agg(np.sum)
        cos1_ng_qty = data_group[20].agg(np.sum)
        cos1_ng_rate = zeroexcerpt(cos1_ng_qty, cos1_isp_qty)
        # 错误信息待处理
        cos1_errors = errors_regroup(sheets, 22)

        # 计算Fos工作台的信息
        fos_insp_qty = data_group[23].agg(np.sum)
        fos_pass_qty = data_group[24].agg(np.sum)
        fos_ng_qty = data_group[25].agg(np.sum)
        fos_ng_rate = zeroexcerpt(fos_ng_qty, fos_insp_qty)
        fos_errors = errors_regroup(sheets, 27)
        # 计算lcd工作台的信息
        lcd_total_insp = data_group[28].agg(np.sum)
        lcd_total_pass = data_group[29].agg(np.sum)
        lcd_ll_ng = data_group[30].agg(np.sum)
        lcd_ll_ngrate = zeroexcerpt(lcd_ll_ng, lcd_total_insp)
        lcd_ymi_ng = data_group[32].agg(np.sum)
        lcd_ymi_ngrate = zeroexcerpt(lcd_ymi_ng, lcd_total_insp)
        lcd_wu_ng = data_group[34].agg(np.sum)
        lcd_wu_ngrate = zeroexcerpt(lcd_wu_ng, lcd_total_insp)
        lcd_bbi_ng = data_group[36].agg(np.sum)
        lcd_bbi_ngrate = zeroexcerpt(lcd_bbi_ng, lcd_total_insp)
        lcd_bm_ng = data_group[38].agg(np.sum)
        lcd_bm_ngrate = zeroexcerpt(lcd_bm_ng, lcd_total_insp)
        # 计算touch工作台信息
        touch_insp = data_group[40].agg(np.sum)
        touch_pass = data_group[41].agg(np.sum)
        touch_ng = data_group[42].agg(np.sum)
        touch_ngrate = zeroexcerpt(touch_ng, touch_insp)
        # 计算current工作台信息
        current_insp = data_group[44].agg(np.sum)
        current_pass = data_group[45].agg(np.sum)
        current_ng = data_group[46].agg(np.sum)
        current_ngrate = zeroexcerpt(current_ng, current_insp)

        # 计算Fliker工作台信息
        flicker_insp = data_group[48].agg(np.sum)
        flicker_pass = data_group[49].agg(np.sum)
        flicker_ng = data_group[50].agg(np.sum)
        flicker_ngrate = zeroexcerpt(flicker_ng, flicker_insp)
        # 计算cosmetic信息
        cos2_insp_qty = data_group[52].agg(np.sum)
        cos2_pass_qty = data_group[53].agg(np.sum)
        cos2_ng_qty = data_group[54].agg(np.sum)

        cos2_ng_rate = zeroexcerpt(cos2_ng_qty, cos2_insp_qty)
        # 错误信息待处理
        cos2_errors = errors_regroup(sheets, 56)
        # remark信息
        remark = ''
        # 使用True标记是否为母项
        merge_list = [[insp_date, shift, project, stage, pn, color,
                      m_type, mac, iar, datecode, ecode, vendor, cfg, batch, total_isp_qty,
                      total_pass_qty, total_ng_qty, total_ng_rate, cos1_isp_qty, cos1_pass_qty, cos1_ng_qty,
                      cos1_ng_rate, cos1_errors, fos_insp_qty, fos_pass_qty,
                      fos_ng_qty, fos_ng_rate, fos_errors, lcd_total_insp, lcd_total_pass,
                      lcd_ll_ng, lcd_ll_ngrate, lcd_ymi_ng, lcd_ymi_ngrate, lcd_wu_ng, lcd_wu_ngrate,
                      lcd_bbi_ng, lcd_bbi_ngrate, lcd_bm_ng, lcd_bm_ngrate, touch_insp, touch_pass,
                      touch_ng, touch_ngrate, current_insp, current_pass, current_ng, current_ngrate,
                      flicker_insp, flicker_pass, flicker_ng, flicker_ngrate, cos2_insp_qty, cos2_pass_qty,
                      cos2_ng_qty, cos2_ng_rate, cos2_errors, remark, True]]

        if len(child_list) > 1:
            merge_list = merge_list+child_list

        merge_lists = merge_lists+merge_list

    return merge_lists


def merge_other_data(datalist, flag):
    convert_data = pd.DataFrame(datalist)
    merge_lists = []
    data = convert_data.groupby([5, 11, 12])
    for i, j in data:
        child_list = j.values.tolist()
        data_group = data.get_group(i)
        sheets = j
        insp_date = timezone(data_group[0])
        shift = shiftzone(sheets, 1)
        project = onlyunique(sheets, 2)
        stage = onlyunique(sheets, 3)
        vendor = onlyunique(sheets, 4)
        pn = onlyunique(sheets, 5)
        mtype = onlyunique(sheets, 6)
        mar = onlyunique(sheets, 7)
        iar = onlyunique(sheets, 8)
        datecode = onlyunique(sheets, 9)
        ecode = onlyunique(sheets, 10)
        cfg = onlyunique(sheets, 11)
        batch = onlyunique(sheets, 12)
        total_ins_qty = data_group[13].agg(np.sum)
        total_pass_qty = data_group[14].agg(np.sum)
        total_ng_qty = data_group[15].agg(np.sum)
        total_ngrate = zeroexcerpt(total_ng_qty, total_ins_qty)
        # cos1信息
        cos1_ins_qty = data_group[17].agg(np.sum)
        cos1_pass_qty = data_group[18].agg(np.sum)
        cos1_ng_qty = data_group[19].agg(np.sum)
        cos1_ngrate = zeroexcerpt(cos1_ng_qty, cos1_ins_qty)
        cos1_errors = errors_regroup(sheets, 21)
        # function工作台信息
        func_ins_qty = data_group[22].agg(np.sum)
        func_pass_qty = data_group[23].agg(np.sum)
        func_ng_qty = data_group[24].agg(np.sum)
        func_ngrate = zeroexcerpt(func_ng_qty, func_ins_qty)
        func_errors = errors_regroup(sheets, 26)

        if flag:
            # fit guage check 工作台
            fgc_ins_qty = data_group[27].agg(np.sum)
            fgc_pass_qty = data_group[28].agg(np.sum)
            fgc_ng_qty = data_group[29].agg(np.sum)
            fgc_ngrate = zeroexcerpt(fgc_ng_qty, fgc_ins_qty)
            fgc_errors = errors_regroup(sheets, 31)
            # cos2工作台信息
            cos2_insp_qty = data_group[32].agg(np.sum)
            cos2_pass_qty = data_group[33].agg(np.sum)
            cos2_ng_qty = data_group[34].agg(np.sum)
            cos2_ngrate = zeroexcerpt(cos2_ng_qty, cos2_insp_qty)
            cos2_errors = errors_regroup(sheets, 36)

            remark = ''

            merge_list = [[insp_date, shift, project, stage, vendor, pn, mtype, mar, iar, datecode,
                          ecode, cfg, batch, total_ins_qty, total_pass_qty, total_ng_qty, total_ngrate,
                          cos1_ins_qty, cos1_pass_qty, cos1_ng_qty, cos1_ngrate, cos1_errors,
                          func_ins_qty, func_pass_qty, func_ng_qty, func_ngrate, func_errors,
                          fgc_ins_qty, fgc_pass_qty, fgc_ng_qty, fgc_ngrate, fgc_errors,
                          cos2_insp_qty, cos2_pass_qty, cos2_ng_qty, cos2_ngrate, cos2_errors, remark, True]]

            if len(child_list) > 1:
                merge_list = merge_list + child_list

            merge_lists = merge_lists + merge_list
        else:
            # cos2工作台信息
            cos2_insp_qty = data_group[27].agg(np.sum)
            cos2_pass_qty = data_group[28].agg(np.sum)
            cos2_ng_qty = data_group[29].agg(np.sum)
            cos2_ngrate = zeroexcerpt(cos2_ng_qty, cos2_insp_qty)
            cos2_errors = errors_regroup(sheets, 31)

            remark = ''

            merge_list = [[insp_date, shift, project, stage, vendor, pn, mtype, mar, iar, datecode,
                          ecode, cfg, batch, total_ins_qty, total_pass_qty, total_ng_qty, total_ngrate,
                          cos1_ins_qty, cos1_pass_qty, cos1_ng_qty, cos1_ngrate, cos1_errors,
                          func_ins_qty, func_pass_qty, func_ng_qty, func_ngrate, func_errors,
                          cos2_insp_qty, cos2_pass_qty, cos2_ng_qty, cos2_ngrate, cos2_errors, remark, True]]

            if len(child_list) > 1:
                merge_list = merge_list + child_list

            merge_lists = merge_lists + merge_list

    return merge_lists
