from __future__ import absolute_import
from Data_analysis import *

from Merge_excel import *
from read_excel_edited import make_sheet
import os
import re
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from TkinterDnD2 import TkinterDnD
from threading import Thread
import tkinter as tk


class Gui:
    # 初始化Gui布局
    def __init__(self, master):
        master.title("IQC日報自動生成軟體--beta2.0")
        # 设置主窗体大小
        master.geometry("760x805+600+150")
        master.resizable(False, False)
        # master.configure(background='#e5e5e6')
        self.master = master

        # 日期列表
        self.cddate = []
        self.date_list = []

        # 选择日报的类型 0为topmodule，1 为hsg和eeparts
        self.report_type = 0

        # 将筛选值传入list获取新的筛选条件及 条件下的报表信息
        self.date_request_list = []
        self.project_request_list = []
        self.stage_request_list = []
        self.dn_request_list = []
        self.pn_request_list = []
        self.cfg_request_list = []
        self.vendor_request_list = []
        self.type_request_list = []

        self.filter_settings = [self.date_request_list, self.project_request_list, self.stage_request_list,
                                self.dn_request_list, self.pn_request_list, self.cfg_request_list,
                                self.type_request_list, self.vendor_request_list]
        # 将方法放入列表
        self.function_list = [self.project_activate, self.stage_activate, self.dn_activate,
                              self.pn_activate, self.cfg_activate, self.type_activate, self.vendor_activate]
        self.buttons_value = []
        self.buttons = []
        self.radios = []

        # 进度条设置
        self.style = ttk.Style()
        self.style.theme_use("alt")
        self.style.configure("RGB(0,102,208).Horizontal.TProgressbar")

        self.data_state = False
        self.flag = 0

        # 文件操作面板
        self.frame_content = tk.LabelFrame(master, text="文件操作:", height=40, relief=tk.RIDGE)
        self.frame_content.pack(fill=tk.X, padx=5, pady=5)

        # 设置打开路径和保存路径的类型
        self.openpath = tk.StringVar()
        self.savepath = tk.StringVar()
        self.xfile = tk.StringVar()
        self.result_list = []

        # 打开文件框（可实现拖拽效果）
        default_value = '選擇文件路徑，拖入文件即可'
        tk.entry = self.build_label_value_block(default_value)
        # 给 widget 添加拖放方法
        self.add_drop_handle(tk.entry, self.handle_drop_on_local_path_entry)

        # 保存文件框
        self.save_name = tk.Entry(self.frame_content, width=55, font=('Verdana', 15), state="normal",
                                  textvariable=self.savepath)
        self.save_name.grid(padx=10, pady=3, row=3, column=0)

        # 打开文件按钮
        ttk.Button(self.frame_content, text="打開文件", width=15, command=self.openfile).grid(row=1, column=1, padx=5, pady=2)

        # 数据分析按钮
        self.data_analyze_btn = ttk.Button(self.frame_content, text="分析數據", width=15, state='disable',
                                           command=lambda: self.thread_it(self.analyze_data))
        self.data_analyze_btn.grid(row=2, column=1, padx=5, pady=2)

        # 保存文件按钮
        self.save_path_btn = ttk.Button(self.frame_content, text="保存路徑", width=15, state='disable', command=self.save_file)
        self.save_path_btn.grid(row=3, column=1, padx=5, pady=3)

        # 这里添加进度条：
        self.pbr = ttk.Progressbar(self.frame_content, length=560, maximum=100, mode='indeterminate', orient=tk.HORIZONTAL)
        self.pbr.grid(row=2, column=0)
        self.pbr.pack_forget()

        # 報表類型選擇
        self.type_frame = tk.LabelFrame(master, text='類型選擇:', height=30, relief=tk.RIDGE)
        self.type_frame.pack(fill=tk.X, padx=5, pady=1)
        self.lab = tk.Label(self.type_frame, width=18)
        self.lab.grid(row=0, column=0, sticky=tk.E, padx=2)

        self.deal = ttk.Button(self.type_frame, text='數據處理生成日報', width=15)
        self.deal.grid(row=0, column=0, padx=8, pady=3)

        self.merge = ttk.Button(self.type_frame, text='日報合成', width=15)
        self.merge.grid(row=1, column=0, padx=8, pady=3)

        # 數據處理單選框生成
        self.radio_var = tk.IntVar()
        self.radio_var.set('')
        self.radio_name = ['Top Module', 'HSG', 'EEPARTS', 'Top Module合成', 'HSG合成', 'EEPARTS合成']
        for radio in self.radio_name:
            radio_index = self.radio_name.index(radio)
            self.radio = ttk.Radiobutton(self.type_frame,
                                         variable=self.radio_var,
                                         value=radio_index,
                                         text=self.radio_name[radio_index],
                                         command=self.get_radio)
            if radio_index < 3:
                self.radio.grid(row=0, column=radio_index + 1, padx=20, pady=3)
            else:
                self.radio.grid(row=1, column=radio_index-2, padx=40, pady=3)
            self.radios.append(self.radio)

        # 条件面板1
        self.apply_frame = tk.LabelFrame(master, text='條件查詢:', height=30, relief=tk.RIDGE)
        self.apply_frame.pack(fill=tk.X, padx=5, pady=1)
        self.lab = tk.Label(self.apply_frame, width=18)
        self.lab.grid(row=0, column=0, sticky=tk.E, padx=2)

        # 滚动板
        self.scrollbar_frame = tk.LabelFrame(master, relief=tk.RIDGE,
                                             # height=280
                                             )
        self.scrollbar_frame.pack(fill=tk.BOTH, padx=1, pady=1)

        self.canvas = tk.Canvas(self.scrollbar_frame, borderwidth=0, background="#fff",
                                # height=280
                                )
        self.frame = tk.Frame(self.canvas, background="#fff",
                              # height=280
                              )

        self.vsb = tk.Scrollbar(self.scrollbar_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas.create_window((4, 4), window=self.frame, anchor=tk.NW)

        self.frame.bind("<Configure>", lambda event, canvas=self.canvas: self.onFrameConfigure(canvas))

        self.check_frame_group = []
        for i in range(7):
            self.check_frame = tk.LabelFrame(self.frame, relief=tk.RIDGE)
            self.check_frame.pack(fill=tk.X, padx=2, pady=1)
            self.check_frame_group.append(self.check_frame)

        # 生成
        self.apply6_frame = tk.LabelFrame(master, relief=tk.RIDGE)
        self.apply6_frame.pack(fill=tk.X, padx=5, pady=2)

        # 选择条件铭牌
        check_frame_name = ['專案', '階段', '班次', '料號', 'CONFIG', '物料名称', '厂商']
        for i in range(7):
            self.tag_label = ttk.Button(self.check_frame_group[i],
                                        text=check_frame_name[i],
                                        width=15,
                                        )
            self.tag_label.grid(row=0, column=0, padx=7, pady=3)

        # 日期框
        self.date = ttk.Button(self.apply_frame, text='日期选择', width=15)
        self.date.grid(row=0, column=0, pady=2)

        # 创建一个开始日期下拉列表
        self.start_date = tk.StringVar()
        self.stChosen = ttk.Combobox(self.apply_frame, width=28, textvariable=self.start_date)
        self.stChosen.grid(row=0, column=1, padx=2)  # 设置其在界面中出现的位置  column代表列   row 代表行
        self.stChosen.bind('<<ComboboxSelected>>', self.get_date)

        # 创建一个结束日期下拉列表
        self.end_date = tk.StringVar()
        self.edChosen = ttk.Combobox(self.apply_frame, width=28, textvariable=self.end_date)
        self.edChosen.grid(row=0, column=2, padx=4)  # 设置其在界面中出现的位置  column代表列   row 代表行
        self.edChosen.bind('<<ComboboxSelected>>', self.get_date)

        # 全选按钮
        self.select_all_var = tk.IntVar()
        self.select_all_var.set(0)
        self.select_all_cb = ttk.Checkbutton(self.apply6_frame,
                                             text='全选条件',
                                             width=37,
                                             variable=self.select_all_var,
                                             command=self.select_all)
        self.select_all_cb.grid(row=6, column=5, padx=13, pady=3)

        # 生成日报按钮
        self.build_report_btn = ttk.Button(self.apply6_frame,
                                           text='生成日報',
                                           width=38,
                                           state='disable',
                                           command=self.get_results)
        self.build_report_btn.grid(row=6, column=6, padx=1, pady=3)

        # 底部面板
        self.bottom_frame = tk.LabelFrame(master, relief=tk.RIDGE)
        self.bottom_frame.pack(fill=tk.X, padx=5, pady=5)

        # 提示面板
        self.remind_frame = tk.LabelFrame(self.bottom_frame, text='溫馨提示:', height=30, relief=tk.RIDGE)
        self.remind_frame.pack(fill=tk.X, padx=5, pady=5)

        self.lb = ttk.Label(self.remind_frame, text='數據越多處理時間越長，請耐心等待！！')
        self.lb.pack(fill=tk.X, padx=10, pady=5)

        self.lb1 = ttk.Label(self.remind_frame, text='')
        self.lb1.pack(fill=tk.X, padx=10, pady=5)

        # logo面板
        self.logo_frame = tk.LabelFrame(self.bottom_frame, relief=tk.RIDGE)
        self.logo_frame.pack(fill=tk.X, padx=5, pady=5)

        # 添加logo画布
        # (150,60),(360,80)
        self.can = tk.Canvas(self.logo_frame,
                             # width=240,
                             height=160,
                             bg='white',
                             )
        self.im = tk.PhotoImage(file='logo-v2-6.png')
        self.can.create_image(360, 80, image=self.im)
        self.can.pack(fill=tk.X, expand=True, side='bottom')

    # 打开文件函数
    def openfile(self):
        # 获取到文件路径
        self.flag = 1
        names = filedialog.askopenfilenames(parent=self.frame_content,
                                            filetypes=[('xlsx file', '.xlsx'), ('xlsx file', '.xlsx')])
        oldopenfilepath = self.openpath.get()

        # 提示框
        self.lb["text"] = '您已选择 ' + str(len(names)) + ' 份文件！'

        if names != oldopenfilepath:
            self.openpath.set(names)
            self.filter_settings.clear()

        self.radio_var.set('')
        self.data_analyze_btn.state(statespec=('disabled', ))
        self.save_path_btn.state(statespec=('disabled', ))
        for r in self.radios:
            r.state(statespec=('!disabled', ))

        self.reprint_canvas()

    # 文件拖拽lable的值
    def build_label_value_block(self, default_value, entry_size=55):
        entry = tk.Entry(self.frame_content, width=entry_size, font=('Verdana', 15), textvariable=self.openpath)
        entry.grid(row=1, column=0, padx=3, pady=1)
        entry.insert(0, default_value)
        return entry

    # 文件推拽处理
    def add_drop_handle(self, widget, handle):
        widget.drop_target_register('DND_Files')
        def drop_enter(event):
            event.widget.focus_force()
            return event.action

        def drop_position(event):
            return event.action

        def drop_leave(event):
            # leaving 应该清除掉之前 drop_enter 的 focus 状态, 怎么清?
            return event.action

        widget.dnd_bind('<<DropEnter>>', drop_enter)
        widget.dnd_bind('<<DropPosition>>', drop_position)
        widget.dnd_bind('<<DropLeave>>', drop_leave)
        widget.dnd_bind('<<Drop>>', handle)

    def handle_drop_on_local_path_entry(self, event):
        self.flag = 0
        if event.data:
            files = event.widget.tk.splitlist(event.data)
            for f in files:
                if os.path.exists(f):
                    event.widget.delete(0, 'end')
                    event.widget.insert('end', f)
                else:
                    print('Not dropping file "%s": file does not exist.' % f)

        for r in self.radios:
            r.state(statespec=('!disabled',))

        return event.action

    def thread_it(self, func):
        # 将函数打包进线程
        # 创建
        t = Thread(target=func)
        # 守护 !!!
        # t.setDaemon(True)
        # 启动
        t.start()
        # 阻塞--卡死界面！
        # t.join()

    # 分析数据
    def analyze_data(self):
        # 判断是否打开文件
        if len(self.openpath.get()) == 0:
            messagebox.showerror(title="錯誤", message="請輸入文件路徑！！", parent=self.frame_content)
            return
        # 获取时间
        t1 = time.time()

        # 获取打开文件的路径
        openpath = self.openpath.get()
        if self.flag == 1:
            # 正则表达式将路径转换为list, re.I忽略大小写，re.S多行处理
            path = re.findall('\'(.*?.xlsx)\'', openpath, re.S | re.I)
        else:
            path=[]
            path.append(openpath)

        # 获取单选框和复选框的值
        radio_value = self.radio_var.get()
        self.pbr.start(interval=10)
        self.reprint_canvas()
        try:
            # 判断是什么部件的表格
            if radio_value == 0:
                self.report_type = 0
                self.lb["text"] = 'Top_module 數據正在分析中，請耐心等待！！'
                self.lb1["text"] = ''
                self.xfile = 'Top_module.xlsx'

                # 调用数据分析方法
                self.result_list = data_analysis_top(path)
                filter_options = self.get_filter(False)
                self.cddate = filter_options[0]

            elif radio_value == 1:
                self.report_type = 1
                self.lb["text"] = 'HSG 數據正在分析中，請耐心等待！！'
                self.lb1["text"] = ''
                self.xfile = 'HSG.xlsx'

                self.result_list = data_analysis_hsg(path)
                filter_options = self.get_filter(True)
                self.cddate = filter_options[0]

            elif radio_value == 2:
                self.report_type = 2
                self.lb["text"] = 'EEPARTS 數據正在分析中，請耐心等待！！'
                self.lb1["text"] = ''
                self.xfile = 'EEPARTS.xlsx'

                self.result_list = data_analysis_eeparts(path)
                filter_options = self.get_filter(True)
                self.cddate = filter_options[0]

            elif radio_value == 3:
                self.report_type = 0
                self.lb["text"] = 'Top_module 報表正在合成中，請耐心等待！！'
                self.lb1["text"] = ''
                self.xfile = 'Top_module.xlsx'

                self.result_list = merge_excel(path, self.report_type)
                filter_options = self.get_filter(True)
                self.cddate = filter_options[0]

            elif radio_value == 4:
                self.report_type = 1
                self.lb["text"] = 'HSG 報表正在合成中，請耐心等待！！'
                self.lb1["text"] = ''
                self.xfile = 'HSG.xlsx'

                self.result_list = merge_excel(path, self.report_type)
                filter_options = self.get_filter(True)
                self.cddate = filter_options[0]

            elif radio_value == 5:
                self.report_type = 2
                self.lb["text"] = 'EEPARTS 報表正在合成中，請耐心等待！！'
                self.lb1["text"] = ''
                self.xfile = 'EEPARTS.xlsx'

                self.result_list = merge_excel(path, self.report_type)

                filter_options = self.get_filter(True)
                self.cddate = filter_options[0]

            # 运行时间
            self.speed1 = time.time() - t1

            # 对时间排序
            self.date_list = sorted(self.cddate)

            self.get_buttons()

            # 设置日期下拉列表的值
            self.stChosen["values"] = self.date_list
            self.edChosen["values"] = self.date_list

            self.data_state = True
            self.pbr.stop()

            self.save_path_btn.state(statespec=('!disabled',))
            self.build_report_btn.state(statespec=('!disabled',))
            for r in self.radios:
                r.state(statespec=('disabled',))

            # 设置下拉列表默认值
            self.stChosen.current(0)
            self.edChosen.current(len(self.date_list)-1)

            self.lb["text"] = '數據分析完成,輸入保存路徑,條件選擇生成報表！！'
            self.lb1["text"] = "一共分析出", str(self.amount_of_sheets()), "條日報信息！！"

        except:
            messagebox.showerror(title="錯誤", message="請選擇正確的文件進行數據分析！！", parent=self.frame_content)
            self.pbr.stop()
            self.date.state(statespec=('disabled',))

    def onFrameConfigure(self, event):
        '''Reset the scroll region to encompass the inner frame'''
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    # radio点击触发函数
    def get_radio(self):
        radio_value = self.radio_var.get()
        if radio_value == 0:
            self.lb["text"] = '您已選擇生成 Top_module 數據報表！！'
        elif radio_value == 1:
            self.lb["text"] = '您已選擇生成 HSG 數據報表！！'
        elif radio_value == 2:
            self.lb["text"] = '您已選擇生成 EEPARTS 數據報表！！'
        elif radio_value == 3:
            self.lb["text"] = '您已選擇合成 Top_module 數據報表！！'
        elif radio_value == 4:
            self.lb["text"] = '您已選擇合成 HSG 數據報表！！'
        elif radio_value == 5:
            self.lb["text"] = '您已選擇合成 EEPARTS 數據報表！！'

        self.data_analyze_btn.state(statespec=('!disabled',))

    # 文件保存路徑
    def save_file(self):
        name = filedialog.asksaveasfilename(parent=self.frame_content, filetypes=[('excel file', '.xlsx')],
                                            initialfile='_Daily_Reports'+'('+self.stChosen.get()+'至'+self.edChosen.get()+')')
        self.savepath.set(name)

    # 導出文件
    def get_results(self):
        if self.openpath.get() and self.data_state:
            t = time.time()
            # 是否有保存路径
            if self.savepath.get():
                result = self.filter_sheets()
                fname = gui.savepath.get()
                sheet = make_sheet(self.xfile, fname, result)

                if self.radio_var.get() == 0 or self.radio_var.get() == 3:
                    self.lb["text"] = 'Top_module 報表正在生成中，請耐心等待！！'
                    sheet.top_module_sum()
                    sheet.ttl_sum_top()

                elif self.radio_var.get() == 1 or self.radio_var.get() == 4:
                    self.lb["text"] = 'HSG 報表正在生成中，請耐心等待！！'
                    sheet.Other_module_sum()
                    sheet.ttl_sum_other(True)

                else:
                    self.lb["text"] = 'EEPARTS 報表正在生成中，請耐心等待！！'
                    sheet.Other_module_sum()
                    sheet.ttl_sum_other(False)

                # 运行耗时
                self.speed = time.time() - t + self.speed1
                # 设置提示框运行耗时的值
                self.lb["text"] = '報表生成成功！！'
                self.lb1["text"] = '本次數據處理一共耗時:' + str(self.speed) + 's'

                messagebox.showerror(title="Success", message='報表生成成功！！', parent=self.frame_content)

                self.savepath.set('')
            else:
                self.save_name.focus()

        else:
            messagebox.showerror(title="錯誤", message="請先打開文件，並進行數據分析！！", parent=self.frame_content)

    def check_list_regetting(self):
        self.date_request_list = self.filter_settings[0]
        self.project_request_list = self.filter_settings[1]
        self.stage_request_list = self.filter_settings[2]
        self.dn_request_list = self.filter_settings[3]
        self.pn_request_list = self.filter_settings[4]
        self.cfg_request_list = self.filter_settings[5]
        self.vendor_request_list = self.filter_settings[6]
        self.type_request_list = self.filter_settings[7]

    # 获取日期至下拉框
    def get_date(self, event):
        self.date_request_list.clear()
        self.clear_button_state()

        st_selected = self.stChosen.get()
        ed_selected = self.edChosen.get()

        date_list = self.get_filter(False)[0]
        index_date = sorted([st_selected, ed_selected])

        for date in date_list:
            if date >= index_date[0] and date <= index_date[1]:
                self.date_request_list.append(date)

        self.change_button_state('select_part')

        self.lb1["text"] = "該條件下一共有", str(self.amount_of_sheets()), "條日報信息！！"

    # 摧毁所有已经生成的条件按钮
    def reprint_canvas(self):
        for button in self.buttons:
            button.destroy()
        self.buttons.clear()
        self.clear_all()
        self.cddate.clear()
        self.stChosen.set('')
        self.edChosen.set('')

    # 数据数量
    def amount_of_sheets(self):
        sheets = self.filter_sheets()
        amount = len(sheets)
        return amount

    # 判断过滤项是否满足表单
    def inornot(self, sheet, request_list, pos_num):
        if request_list:
            flag = sheet[pos_num] in request_list
        else:
            flag = True
        return flag

    # 获取该条件下的日报信息
    def filter_sheets(self):
        sheets_data = self.result_list
        sheets = []
        if self.report_type == 0:
            for sheet in sheets_data:
                data_request = self.inornot(sheet, self.date_request_list, 0)
                project_request = self.inornot(sheet, self.project_request_list, 2)
                stage_request = self.inornot(sheet, self.stage_request_list, 3)
                dn_request = self.inornot(sheet, self.dn_request_list, 1)
                pn_request = self.inornot(sheet, self.pn_request_list, 4)
                cfg_request = self.inornot(sheet, self.cfg_request_list, 12)
                type_request = self.inornot(sheet, self.type_request_list, 6)
                vendor_request = self.inornot(sheet, self.vendor_request_list, 11)
                if data_request and dn_request and pn_request and type_request and vendor_request and cfg_request and project_request and stage_request:
                    sheets.append(sheet)
        else:
            for sheet in sheets_data:
                data_request = self.inornot(sheet, self.date_request_list, 0)
                project_request = self.inornot(sheet, self.project_request_list, 2)
                stage_request = self.inornot(sheet, self.stage_request_list, 3)
                dn_request = self.inornot(sheet, self.dn_request_list, 1)
                pn_request = self.inornot(sheet, self.pn_request_list, 5)
                type_request = self.inornot(sheet, self.type_request_list, 6)
                vendor_request = self.inornot(sheet, self.vendor_request_list, 4)
                cfg_request = self.inornot(sheet, self.cfg_request_list, 11)
                if data_request and dn_request and pn_request and type_request and vendor_request and cfg_request and project_request and stage_request:
                    sheets.append(sheet)
        return sheets

    # 获取该信息的筛选条件项
    def get_filter(self, flag):
        date_options = []
        project_options = []
        stage_options = []
        dn_options = []
        pn_options = []
        cfg_options = []
        vendor_options = []
        type_options = []
        filter_options = [date_options, project_options, stage_options,
                          dn_options, pn_options, cfg_options,
                          type_options, vendor_options]
        if self.report_type == 0:
            report_index = [0, 2, 3, 1, 4, 12, 6, 11]
        else:
            report_index = [0, 2, 3, 1, 5, 11, 6, 4]

        if flag:
            sheets_data = self.filter_sheets()
        else:
            sheets_data = self.result_list

        for sheet in sheets_data:
            for filter_option in filter_options:
                index = filter_options.index(filter_option)
                pos_num = report_index[index]
                if not sheet[pos_num] in filter_option:
                    filter_option.append(sheet[pos_num])

        return filter_options

    # 取消全选
    def clear_all(self):
        self.date_request_list.clear()
        self.project_request_list.clear()
        self.stage_request_list.clear()
        self.dn_request_list.clear()
        self.pn_request_list.clear()
        self.cfg_request_list.clear()
        self.vendor_request_list.clear()
        self.type_request_list.clear()
        self.lb1["text"] = "該條件下一共有", str(self.amount_of_sheets()), "條日報信息！！"

    # 全选
    def fill_all(self):
        self.filter_settings = self.get_filter(False)
        self.date_request_list = self.filter_settings[0]
        self.project_request_list = self.filter_settings[1]
        self.stage_request_list = self.filter_settings[2]
        self.dn_request_list = self.filter_settings[3]
        self.pn_request_list = self.filter_settings[4]
        self.cfg_request_list = self.filter_settings[5]
        self.type_request_list = self.filter_settings[6]
        self.vendor_request_list = self.filter_settings[7]
        self.lb1["text"] = "該條件下一共有", str(self.amount_of_sheets()), "條日報信息！！"
        return self.filter_settings

    # 获取过滤项的长度
    def filter_result_col(self):
        num = 0
        results = self.filter_settings
        for result in results:
            if len(result) > 0:
                num = num + 1
        return num

    def mix_get_filter(self):

        filter_results = self.get_filter(True)
        mix_result = []
        for i in filter_results:
            for j in i:
                mix_result.append(j)
        return mix_result

    def mix_filter_settings(self):
        filter_results = self.filter_settings
        mix_result = []
        for i in filter_results:
            for j in i:
                mix_result.append(j)
        return mix_result

    def reset_settings(self, result_list):

        for button in self.buttons_value:
            v, n = button
            if self.filter_result_col() == 0 or (self.filter_result_col() == 1 and len(result_list) != 0):
                if v.get() == 1 and (n not in result_list):
                    result_list.append(n)
                elif v.get() == 0 and (n in result_list):
                    result_list.remove(n)
            else:
                if v.get() == 1 and (n not in self.mix_filter_settings()):
                    result_list.append(n)
                elif v.get() == 0 and (n in self.mix_filter_settings()):
                    result_list.remove(n)
        self.lb1["text"] = "該條件下一共有", str(self.amount_of_sheets()), "條日報信息！！"

    def change_button_state(self, model):
        if model == 'select_part':
            for button in self.buttons:
                if button.cget('text') not in self.mix_get_filter() and button.cget('text') not in self.get_filter(False)[0]:
                    button.configure(state='disable', cursor='circle')
                else:
                    button.configure(state='normal', cursor='arrow')

        elif model == 'select_all':
            # 选择全部
            for button in self.buttons:
                button.configure(state='normal', cursor='arrow')
                button.select()

        elif model == 'clear_all':
            # 取消全部
            for button in self.buttons:
                button.configure(state='normal', cursor='arrow')
                button.deselect()

    def clear_button_state(self):
        for button in self.buttons:
            button.configure(state='normal')
            button.deselect()
        for setting in self.filter_settings[1:]:
            setting.clear()

    def project_activate(self):
        self.reset_settings(self.project_request_list)

    def stage_activate(self):
        self.reset_settings(self.stage_request_list)

    def dn_activate(self):
        self.reset_settings(self.dn_request_list)

    def pn_activate(self):
        self.reset_settings(self.pn_request_list)

    def cfg_activate(self):
        self.reset_settings(self.cfg_request_list)

    def vendor_activate(self):
        self.reset_settings(self.vendor_request_list)

    def type_activate(self):
        self.reset_settings(self.type_request_list)

    def select_all(self):
        check_box = []
        for button in self.buttons_value:
            v, n = button
            var = v.get()
            if not var in check_box:
                check_box.append(v.get())

        # 如果选项不一，先全选
        if len(check_box) != 1:
            for button in self.buttons_value:
                v, n = button
                v.set(1)
                self.select_all_var.set(1)
                self.fill_all()
                self.change_button_state('select_all')

        elif len(check_box) == 1 and check_box[0] == 0:
            # 全选
            for button in self.buttons_value:
                v, n = button
                v.set(1)
            self.fill_all()
            self.change_button_state('select_all')

        else:
            # 取消
            for button in self.buttons_value:
                v, n = button
                v.set(0)
            self.clear_all()

    def get_buttons(self):
        all_settings = self.get_filter(True)
        for part_settings in all_settings[1:]:
            index = all_settings[1:].index(part_settings)
            row = 0
            col = 1
            ckbtn_len = len(max(part_settings))
            count = 20/ckbtn_len
            for setting in part_settings:
                var = tk.IntVar()
                var.set(0)
                self.button = tk.Checkbutton(self.check_frame_group[index],
                                             text=setting,
                                             width=int(ckbtn_len)+8,
                                             command=self.function_list[index],
                                             variable=var,
                                             )
                self.button.grid(row=row, column=col, sticky=tk.W, padx=2)

                col = col+1
                if int(count) < 6:
                    if col == int(count)+3:
                        row = row + 1
                        col = 1
                else:
                    if col == 6:
                        row = row+1
                        col = 1

                self.button.deselect()
                self.buttons_value.append((var, setting))
                self.buttons.append(self.button)


if __name__ == '__main__':
    root = TkinterDnD.Tk()
    gui = Gui(root)
    root.mainloop()

