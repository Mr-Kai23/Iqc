from hsg import *
from eeparts import *
from topmodule import *


def data_analysis_top(filepath):
    reports_sum = []

    for single_excel_src in filepath:
        content = pd.read_excel(single_excel_src, sheet_name=None)
        sheetnames = content.keys()

        for sheetname in sheetnames:
            if sheetname.strip() != "Summary" and sheetname.strip() != "不良中英對比":
                content = pd.read_excel(single_excel_src, sheet_name=sheetname, header=[0, 1])
                df = pd.DataFrame(data=content)
                heads_name = df.columns.tolist()
                date_lists = df[heads_name[1][0]][heads_name[1][1]].drop_duplicates()

                for date in date_lists:
                    if str(date) != "NaT":
                        reports = get_daily_reports_top(date, df)

                        reports_sum = reports_sum+reports

    return reports_sum


def data_analysis_hsg(filepath):
    reports_sum = []

    for single_excel_src in filepath:
        content = pd.read_excel(single_excel_src, sheet_name=None)
        sheetnames = content.keys()

        for sheetname in sheetnames:
            if sheetname.strip() != "Summary" and sheetname.strip() != "不良中英對比":
                content = pd.read_excel(single_excel_src, sheet_name=sheetname, header=[0])
                df = pd.DataFrame(data=content)
                heads_name = df.columns.tolist()
                date_lists = df[heads_name[0]].drop_duplicates()

                for date in date_lists:
                    if str(date) != "NaT":
                        reports = get_daily_reports_hsg(date, df)

                        reports_sum = reports_sum + reports

    return reports_sum


def data_analysis_eeparts(filepath):
    reports_sum = []

    for single_excel_src in filepath:
        content = pd.read_excel(single_excel_src, sheet_name=None)
        sheetnames = content.keys()

        for sheetname in sheetnames:
            if sheetname.strip() != "Summary" and sheetname.strip() != "不良中英對比":
                content = pd.read_excel(single_excel_src, sheet_name=sheetname, header=[0])
                df = pd.DataFrame(data=content)
                heads_name = df.columns.tolist()
                date_lists = df[heads_name[0]].drop_duplicates()

                for date in date_lists:
                    if str(date) != "NaT":
                        reports = get_daily_reports_eeparts(date, df)

                        reports_sum = reports_sum + reports

    return reports_sum






