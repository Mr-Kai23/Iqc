import pandas as pd
import numpy as np
from collections import Counter
import math


# 设置控制台的输出选项
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 5000)


# 工作台只检查一项的信息生成(根据表单和工作台名称)
def single_station(df_form, workstation_name):
    kk = df_form[workstation_name]
    name = workstation_name
    mes = []

    for row in kk.iterrows():
        mes.append(row[1][0])
    results = Counter(mes)
    isp_qty = len(kk)-results['/']
    pass_qty = results['OK']
    ng_qty = isp_qty-pass_qty

    if isp_qty == 0:
        ng_rate = 0
    else:
        ng_rate = format(ng_qty/isp_qty, '.0%')

    return {'isp_qty': isp_qty, 'pass_qty': pass_qty, 'ng_qty': ng_qty, 'ng_rate': ng_rate, 'name': name}


# 多级索引下单工作台的信息生成
def single_stations(df_form, workstation_name, workstation_option):
    kk = df_form[workstation_name][workstation_option]
    name = kk.name
    results = Counter(kk)
    isp_qty = len(kk)-results['/']
    pass_qty = results['OK']
    ng_qty = isp_qty-pass_qty

    if isp_qty == 0:
        ng_rate = 0
    else:
        ng_rate = format(ng_qty/isp_qty, '.0%')

    return {'isp_qty': isp_qty, 'pass_qty': pass_qty, 'ng_qty': ng_qty, 'ng_rate': ng_rate, 'name': name}


# 多项检查工作台信息生成(根据表单和工作台名称)
def multiple_station(df_form, workstation):
    kk = df_form[workstation]
    ng_qty = 0
    pass_qty = 0

    for row in kk.iterrows():
        dic = np.unique(row[1])

        if len(dic) == 1:
            if dic[0].lower() == "ok":
                pass_qty += 1
            else:
                ng_qty += 1
                # print("not ok")
        else:
            ng_qty += 1
            # print("not ok")
    isp_qty = pass_qty+ng_qty

    if isp_qty == 0:
        ng_rate = 0
    else:
        ng_rate = format(ng_qty/isp_qty, '.0%')

    return {'pass_qty': pass_qty, 'ng_qty': ng_qty, 'ng_rate': ng_rate, 'isp_qty': isp_qty}


# 检查工作台错误信息,返回一个list,并计算出错误信息的数量
def wst_errors(df_form, workstation):
    errors = []
    kk = df_form[workstation]
    results = ''

    for row in kk.iterrows():
        for j in row[1]:
            if j.lower() != "ok":
                errors.append(j)

    results_pro = Counter(errors)

    for key, value in results_pro.items():
        results = results + (key + ':\t' + str(value) + 'x\n')

    return results


# 信息归一
def onlyunique(df_form, m):
    heads_name = df_form.columns.tolist()
    sheet = df_form[heads_name[m][0]][heads_name[m][1]]
    unique = np.unique(sheet)[0]

    if type(unique) == np.float64 and math.isnan(unique):
        return 'unknow'
    else:
        return unique


# 根据检查日期获取工作日报所需内容
def generate_sheets(insp_date, df):
    sheets = []
    # 获取表格的head名
    heads_name = df.columns.tolist()
    check_all = df[df[heads_name[1][0]][heads_name[1][1]] == insp_date]
    # 日期下面的多级检索(白班，日班)
    day_or_night = check_all[heads_name[4][0]].drop_duplicates()

    for i in day_or_night.values:

        # 得出当天白班或晚班的sheet
        check_parts = check_all[check_all[heads_name[4][0]][heads_name[4][1]] == i[0]]
        # 得出班次下所质检的不同pn
        p_n = check_parts[heads_name[5][0]].drop_duplicates()

        for j in p_n.values:

            check_pns = check_parts[check_parts[heads_name[5][0]][heads_name[5][1]] == j[0]]
            # 不同pn下的不同config的多级检索
            cfgs = check_pns[heads_name[12][0]].drop_duplicates()

            for k in cfgs.values:
                sheet = check_pns[check_pns[heads_name[12][0]][heads_name[12][1]] == k[0]]
                # cfg值获取
                cfg_num = k[0]
                # 班次获取
                shift = i[0]
                # pn值获取
                pn_num = j[0]
                # 专案名获取
                project_name = onlyunique(sheet, 2)
                # 专案阶段获取
                stage = onlyunique(sheet, 3)
                # 颜色获取
                color = onlyunique(sheet, 6)
                # 物料类型
                type = onlyunique(sheet, 7)
                # 厂商名称
                vendor = onlyunique(sheet, 8)
                # print("22522", type(vendor))
                # Material Actual Rev
                MAR = onlyunique(sheet, 9)
                # IQC Actual Rev
                IAR = onlyunique(sheet, 10)
                # Date code
                datecode = onlyunique(sheet, 11)
                # 批量值获取
                rec_qty = onlyunique(sheet, 13)
                # E-code
                ecode = onlyunique(sheet, 17)

                # 检查量获取
                ins_qty = len(sheet)
                # 获取该表单下的pass和ng量
                p = check_is_quality(sheet)
                pass_qty = p['pass_qty']
                ng_qty = p['ng_qty']

                # ng率计算
                if ins_qty == 0:
                    ng_rate = 0
                else:
                    ng_rate = format(ng_qty/ins_qty, '.0%')

                # 检查日期获取
                inspect_date = insp_date.strftime('%Y-%m-%d')
                # 将分类出的信息打包传给list
                sheets.append([sheet, inspect_date, shift, project_name,
                               stage, pn_num, color, type,
                               MAR, IAR, datecode, ecode, vendor,
                               cfg_num, rec_qty, ins_qty, pass_qty, ng_qty, ng_rate])

    return sheets


def check_is_quality(df_form):
    heads_name = df_form.columns.tolist()

    check_list = [heads_name[22][0], heads_name[26][0], heads_name[27][0],
                  heads_name[31][0], heads_name[36][0],
                  heads_name[37][0], heads_name[38][0],
                  heads_name[39][0], heads_name[40][0]]

    dataform = df_form[check_list]

    match_list = ['ok', 'OK', '/', '\'', np.nan]
    ng_qty = 0
    pass_qry = 0
    check_list_info = []

    for row in dataform.iterrows():
        infos = []

        for info in row[1]:
            if not pd.isnull(info):
                infos.append(info)
        check_list_info.append(infos)

    for i in check_list_info:
        mes = set(i)
        if mes.issubset(set(match_list)):
            pass_qry += 1
        else:
            ng_qty += 1

    return {"pass_qty": pass_qry, "ng_qty": ng_qty}


def get_daily_reports_top(insp_date, df):
    daily_reports = []
    heads_name = df.columns.tolist()

    sheets = generate_sheets(insp_date, df)

    for i in sheets:

        sheet = i[0]

        # cos1工作台信息 调用函数得到数据
        cos1 = multiple_station(sheet, heads_name[22][0])

        # cos1工作台报错信息 调用函数得到数据
        cos1_errors = wst_errors(sheet, heads_name[22][0])

        # cos2工作台信息
        cos2 = single_station(sheet, heads_name[26][0])

        # cos2工作台的错误信息
        cos2_errors = wst_errors(sheet, heads_name[26][0])

        # FOS 工作台信息 调用函数得到数据
        fos = multiple_station(sheet, heads_name[27][0])

        # FOS工作台错误信息
        fos_errors = wst_errors(sheet, heads_name[27][0])

        # TOUCH工作台信息
        touch = multiple_station(sheet, heads_name[36][0])

        # Current Flicker Ink工作台信息
        current = single_station(sheet, heads_name[37][0])
        flicker = single_station(sheet, heads_name[38][0])
        ink = single_station(sheet, heads_name[39][0])

        # LCD-Uniformity(MP6)工作台信息
        lcd_sum = multiple_station(sheet, heads_name[31][0])
        lcd_ll = single_stations(sheet, heads_name[31][0], heads_name[31][1])
        lcd_ymi = single_stations(sheet, heads_name[31][0], heads_name[32][1])
        lcd_wu = single_stations(sheet, heads_name[31][0], heads_name[33][1])
        lcd_bbi = single_stations(sheet, heads_name[31][0], heads_name[34][1])
        lcd_bm = single_stations(sheet, heads_name[31][0], heads_name[35][1])

        messages = [cos1['isp_qty'], cos1['pass_qty'], cos1['ng_qty'], cos1['ng_rate'], cos1_errors,
                    fos['isp_qty'], fos['pass_qty'], fos['ng_qty'], fos['ng_rate'], fos_errors,
                    lcd_sum['isp_qty'], lcd_sum['pass_qty'], lcd_ll['ng_qty'], lcd_ll['ng_rate'],
                    lcd_ymi['ng_qty'], lcd_ymi['ng_rate'], lcd_wu['ng_qty'], lcd_wu['ng_rate'],
                    lcd_bbi['ng_qty'], lcd_bbi['ng_rate'], lcd_bm['ng_qty'], lcd_bm['ng_rate'],
                    touch['isp_qty'], touch['pass_qty'], touch['ng_qty'], touch['ng_rate'],
                    current['isp_qty'], current['pass_qty'], current['ng_qty'], current['ng_rate'],
                    flicker['isp_qty'], flicker['pass_qty'], flicker['ng_qty'], flicker['ng_rate'],
                    cos2['isp_qty'], cos2['pass_qty'], cos2['ng_qty'], cos2['ng_rate'], cos2_errors,
                    ]

        # 删除第一个sheet
        i.pop(0)

        for m in messages:
            i.append(m)
        daily_reports.append(i)

    return daily_reports


if __name__ == '__main__':
    src = '/Users/jarvis01/Desktop/PF P1 Top Module Daily Report.xlsx'
    content = pd.read_excel(src, sheet_name=None)
    sheetnames = content.keys()

    # f = open('reports.json', 'wb')
    # report_sheets = []

    for sheetname in sheetnames:
        if sheetname.strip() != "Summary" and sheetname.strip() !="不良中英對比":
            print(sheetname)
            content = pd.read_excel(src, sheet_name=sheetname, header=[0, 1])
            df = pd.DataFrame(data=content)
            heads_name = df.columns.tolist()
            # print(heads_name)
            date_lists = df[heads_name[1][0]][heads_name[1][1]].drop_duplicates()
            for date in date_lists:
                if str(date) != "NaT":
                    reports = get_daily_reports_top(date, df)
                    for report in reports:
                        print(report)
                        # report_sheets.append(report)

    # pickle.dump(report_sheets,f)
    print("日报已经生成！")
    # f.close()








