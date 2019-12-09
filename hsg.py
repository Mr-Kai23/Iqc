import pandas as pd
import numpy as np
from collections import Counter
import math
import datetime

# 设置控制台的输出选项
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 5000)


# 工作台只检查一项的信息生成(根据表单和工作台名称)
def single_station(df_form, workstation_name):
    heads_name = df_form.columns.tolist()
    name = workstation_name
    mes = []
    check_col = [heads_name[21]]
    df_form = pd.concat([df_form[workstation_name], df_form[check_col]], axis=1)
    for row in df_form.iterrows():
        sqeconfirm = row[1][len(row[1])-1]
        if type(sqeconfirm) == str and sqeconfirm.lower()=='ok':
            mes.append('OK')
        else:
            mes.append(row[1][0])
    results = Counter(mes)
    isp_qty = len(df_form)-results['/']
    pass_qty = results['OK']+results['ok']
    ng_qty = isp_qty-pass_qty

    if isp_qty == 0:
        ng_rate = 0
    else:
        ng_rate = format(ng_qty/isp_qty, '.0%')

    return {'isp_qty': isp_qty, 'pass_qty': pass_qty, 'ng_qty': ng_qty, 'ng_rate': ng_rate, 'name': name}


# 检查工作台错误信息,返回一个list,并计算出错误信息的数量
def wst_errors(df_form, workstation):
    heads_name = df_form.columns.tolist()
    errors = []
    check_col = [heads_name[21]]
    df_form = pd.concat([df_form[workstation], df_form[check_col]],axis=1)
    results = ''
    for row in df_form.iterrows():
        sqeconfirm = row[1][len(row[1])-1]
        if type(sqeconfirm) == str and sqeconfirm.lower() == 'ok':
            continue
        else:
            for j in row[1][:len(row[1])-1]:
                if type(j) == str and j.lower() != "ok":
                    errors.append(j)
    results_pro = Counter(errors)
    for key, value in results_pro.items():
        results = results + (key + ':\t' + str(value) + 'x\n')
    return results


# 信息归一
def onlyunique(df_form, m):
    heads_name = df_form.columns.tolist()
    sheet = df_form[heads_name[m]]
    unique = sheet.unique()
    if len(unique) == 0 or (type(unique[0]) == np.float64 and math.isnan(unique[0])):
        return 'unknow'
    else:
        return unique[0]


def dayToDate(sheet):
    datecode = onlyunique(sheet, 15)
    startYear = 1900
    d2 = datetime.datetime.now().strftime('%Y/%m/%d')
    if type(datecode) == 'numpy.int64':
        while datecode > 0:
            if datecode < 366:
                d1 = datetime.date(startYear, 1, 1)
                d2 = d1 + datetime.timedelta(int(datecode) - 2)
                d2 = datetime.datetime.strftime(d2, '%Y/%m/%d')
            else:
                startYear += 1

            if (startYear % 4 == 0 and startYear % 100 != 0) or (startYear % 400 == 0 and startYear % 3200 != 0):
                datecode -= 366
            else:
                datecode -= 365

    elif type(datecode) == 'datetime.datetime':
        np.datetime64(datecode).astype(datetime.datetime)
        d2 = datetime.datetime.strftime(datecode, '%Y/%m/%d')

    return d2


# 根据检查日期获取工作日报所需内容
def generate_sheets(insp_date, df):
    # 获取表格的head名
    heads_name = df.columns.tolist()
    sheets = []
    # 获取表格的head名
    check_all = df[df[heads_name[0]] == insp_date]
    day_or_night = check_all[heads_name[3]].drop_duplicates()

    for i in day_or_night.values:
        check_parts = check_all[check_all[heads_name[3]] == i]
        p_n = check_parts[heads_name[5]].drop_duplicates()
        for j in p_n.values:
            check_pns = check_parts[check_parts[heads_name[5]] == j]
            cfgs = check_pns[heads_name[7]].drop_duplicates()
            for k in cfgs.values:
                check_cfgs = check_pns[check_pns[heads_name[7]] == k]
                # 不同cfg下不同batch的多级索引
                batchs = check_cfgs[heads_name[8]].drop_duplicates()
                for batch in batchs.values:
                    sheet = check_cfgs[check_cfgs[heads_name[8]] == batch]
                    sheet = sheet.reset_index(drop=True)
                    nanIndexList = sheet[sheet[heads_name[10]].isnull()].index.tolist()
                    sheet.drop(nanIndexList, inplace = True)
                    # 获取质检日期
                    inspect_date = insp_date.strftime('%Y-%m-%d')
                    # 获取班次
                    shift = i
                    # 获取专案名称,专案阶段
                    project_name = onlyunique(sheet, 1)
                    stage = onlyunique(sheet, 2)
                    # 获取厂商名称，pn，tpye
                    vendor = onlyunique(sheet, 4)
                    pn_num = j
                    type = onlyunique(sheet, 6)
                    MAR = onlyunique(sheet, 13)
                    IAR = onlyunique(sheet, 14)
                    datecode = dayToDate(sheet)
                    ecode = onlyunique(sheet, 12)
                    # 获取cfg
                    cfg_num = k
                    rec_qty = onlyunique(sheet, 8)
                    ins_qty = len(sheet)
                    p = check_is_quality(sheet)
                    ng_qty = p['ng_qty']
                    pass_qty = p['pass_qty']
                    if ins_qty == 0:
                        ng_rate = 0
                    else:
                        ng_rate = format(ng_qty/ins_qty, '.0%')
                    # 将分类出的信息打包传给list
                    sheets.append([sheet, inspect_date, shift, project_name, stage, vendor, pn_num, type, MAR, IAR,
                                    datecode, ecode, cfg_num, rec_qty, ins_qty, pass_qty, ng_qty, ng_rate])

    return sheets


def check_is_quality(df_form):
    heads_name = df_form.columns.tolist()

    check_list = [heads_name[16], heads_name[17], heads_name[18],
                  heads_name[19], heads_name[20], heads_name[21]]
    dataform = df_form[check_list]
    match_list = ['ok', 'OK', '/', np.nan]
    ng_qty = 0
    pass_qry = 0

    for row in dataform.iterrows():
        check_list_info = []
        infos = []
        sqeconfirm = row[1][len(row[1])-1]
        if type(sqeconfirm) == np.float:
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
        elif type(sqeconfirm)==str:
            if sqeconfirm.lower() == 'ok':
                pass_qry += 1
            else:
                ng_qty += 1

    return {"pass_qty": pass_qry, "ng_qty": ng_qty}


def get_daily_reports_hsg(insp_date, sheet):
    daily_reports = []
    heads_name = sheet.columns.tolist()
    sheets = generate_sheets(insp_date, sheet)

    for i in sheets:
        sheet = i[0]
        # cos1工作台信息 调用函数得到数据，错误信息
        cos1 = single_station(sheet, heads_name[16])
        cos1_errors = wst_errors(sheet, heads_name[16])

        # cos2工作台信息，错误信息
        cos2 = single_station(sheet, heads_name[19])
        cos2_errors = wst_errors(sheet, heads_name[19])

        # FUN 工作台信息 调用函数得到数据,错误信息
        func = single_station(sheet, heads_name[17])
        func_errors = wst_errors(sheet, heads_name[17])

        # Fit guage check工作台信息，错误信息
        fit = single_station(sheet, heads_name[18])
        fit_errors = wst_errors(sheet, heads_name[18])

        # SQE 工作台信息，错误信息
        sqe = single_station(sheet, heads_name[20])
        seq_errors = wst_errors(sheet, heads_name[20])

        messages = [cos1['isp_qty'], cos1['pass_qty'], cos1['ng_qty'], cos1['ng_rate'], cos1_errors,
                    func['isp_qty'], func['pass_qty'], func['ng_qty'], func['ng_rate'], func_errors,
                    fit['isp_qty'], fit['pass_qty'], fit['ng_qty'], fit['ng_rate'], fit_errors,
                    cos2['isp_qty'], cos2['pass_qty'], cos2['ng_qty'], cos2['ng_rate'], cos2_errors,
                    ]

        # 删除第一个sheet
        i.pop(0)

        for m in messages:
            i.append(m)
        daily_reports.append(i)

    return daily_reports


if __name__ == '__main__':
    src = '/Users/andy/Desktop/PG-HSG Inspec report - - .xlsx'
    content = pd.read_excel(src, sheet_name=None)
    sheetnames = content.keys()
    for sheetname in sheetnames:
        print(sheetname)
        if (sheetname.strip() != "Summary") and (sheetname.strip() != "不良中英對比"):
            content = pd.read_excel(src, sheet_name=sheetname, header=[0])
            df = pd.DataFrame(data=content)
            heads_name = df.columns.tolist()
            # print(heads_name)
            date_lists = df[heads_name[0]].drop_duplicates()
            for date in date_lists:
                if str(date) != "NaT":
                    reports = get_daily_reports_hsg(date, df)
                    for report in reports:
                        print(report)









