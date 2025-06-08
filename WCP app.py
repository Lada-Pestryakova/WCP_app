# coding: utf-8
import os
import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog
from tkinter.messagebox import showerror, showwarning
import matplotlib
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import numpy as np
from scipy.signal import find_peaks  # Для поиска пиков
import struct
import json
import xlsxwriter

# Настройка matplotlib на использование шрифта, поддерживающего Unicode
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Tahoma']
matplotlib.rcParams['axes.unicode_minus'] = False  # Исправление отображения минуса


class App:
    def __init__(self, window):
        self.ratio_dict = {}
        self.window = window
        self.window.title("WCP app")
        self.window.geometry('1160x600+30+120')
        self.window["bg"] = "whitesmoke"
        self.additional_windows = []  # Список для хранения дополнительных окон

        # Имя файла
        self.file_name_label = ttk.Label(window, text="File name", background="whitesmoke", anchor="c")
        self.file_name_label.place(relx=0.22, rely=0.005, relwidth=0.2, relheight=0.04)

        # Номер комплекса
        self.complex_name_label = ttk.Label(window, text="Complex name", background="whitesmoke", anchor="c")
        self.complex_name_label.place(relx=0.76, rely=0.005, relwidth=0.1, relheight=0.04)

        # Кнопка открытия файла
        self.start_button = ttk.Button(window, text="Открыть", command=self.open_file)
        self.start_button.place(relx=0.01, rely=0.95, relwidth=0.06, relheight=0.04)

        # Кнопка выбора сегментов записи
        self.part_button = ttk.Button(window, text="Просмотр записи", state=tk.DISABLED, command=self.part_graph)
        self.part_button.place(relx=0.09, rely=0.95, relwidth=0.1, relheight=0.04)

        # Перелистывание сегментов
        self.prev_rec_button = ttk.Button(window, text="<", state=["disabled"])
        self.prev_rec_button.place(relx=0.09, rely=0.91, relwidth=0.02, relheight=0.04)
        self.next_rec_button = ttk.Button(window, text=">", state=["disabled"])
        self.next_rec_button.place(relx=0.17, rely=0.91, relwidth=0.02, relheight=0.04)

        # Номер сегмента
        self.n_record_entry = Entry(justify=CENTER)
        self.n_record_entry.insert(0, 00)
        self.n_record_entry.place(relx=0.12, rely=0.915, relwidth=0.02, relheight=0.03)

        # Общее количество сегментов
        self.n_records_label = ttk.Label(window, text="/00", background="whitesmoke", anchor="c")
        self.n_records_label.place(relx=0.14, rely=0.91, relwidth=0.03, relheight=0.04)

        # Кнопка выбора участков записи
        self.zoom_in_button = ttk.Button(window, text="Выбрать ПД", state=["disabled"], command=self.zoom_sig)
        self.zoom_in_button.place(relx=0.20, rely=0.95, relwidth=0.07, relheight=0.04)

        # Кнопка установки линии нуля
        self.zerolevel_button = ttk.Button(window, text="0 уровень", state=["disabled"], command=self.zerolevel)
        self.zerolevel_button.place(relx=0.27, rely=0.95, relwidth=0.07, relheight=0.04)

        # Масштабирование сигнала
        self.frame1 = ttk.Frame(borderwidth=1, relief=SOLID)
        self.frame1.place(relx=0.35, rely=0.91, relwidth=0.04, relheight=0.08)
        self.multi_var = StringVar(value="x1")
        self.x1_radiobutton = ttk.Radiobutton(self.frame1, text="", value="x1", variable=self.multi_var,
                                              state=["disabled"], command=self.x1)
        self.x1_radiobutton.place(relx=0, rely=0)
        self.x10_radiobutton = ttk.Radiobutton(self.frame1, text="", value="x0.1", variable=self.multi_var,
                                               state=["disabled"], command=self.x10)
        self.x10_radiobutton.place(relx=0, rely=0.5)
        self.x1_label = tk.Label(self.frame1, text="x1", relief=tk.FLAT, borderwidth=0,
                                 background="whitesmoke")
        self.x1_label.place(relx=0.45, rely=0.05)
        self.x10_label = tk.Label(self.frame1, text="x0.1", relief=tk.FLAT, borderwidth=0,
                                  background="whitesmoke")
        self.x10_label.place(relx=0.45, rely=0.55)

        # Инвертирование сигнала + -
        self.frame2 = ttk.Frame(borderwidth=1, relief=SOLID)
        self.frame2.place(relx=0.395, rely=0.91, relwidth=0.03, relheight=0.08)
        self.sign_var = StringVar(value="+")
        self.original_radiobutton = ttk.Radiobutton(self.frame2, text="", value="+", variable=self.sign_var,
                                                    state=["disabled"], command=self.original)
        self.original_radiobutton.place(relx=0, rely=0)
        self.inverse_radiobutton = ttk.Radiobutton(self.frame2, text="", value="-", variable=self.sign_var,
                                                   state=["disabled"], command=self.inverse)
        self.inverse_radiobutton.place(relx=0, rely=0.5)
        self.plus_label = tk.Label(self.frame2, text="+", relief=tk.FLAT, borderwidth=0,
                                   background="whitesmoke")
        self.plus_label.place(relx=0.55, rely=0.05)
        self.minus_label = tk.Label(self.frame2, text="-", relief=tk.FLAT, borderwidth=0,
                                    background="whitesmoke")
        self.minus_label.place(relx=0.58, rely=0.55)

        # График записи
        self.fig1 = Figure(facecolor="whitesmoke")
        self.ax1 = self.fig1.add_subplot(111)
        self.fig1.tight_layout()
        self.canvas1 = FigureCanvasTkAgg(self.fig1, master=window)
        self.canvas1.draw()
        self.canvas1.get_tk_widget().configure(background="whitesmoke", highlightcolor='whitesmoke',
                                               highlightbackground='whitesmoke')
        self.canvas1.get_tk_widget().place(relx=0, rely=0.07, relwidth=0.61, relheight=0.83)
        self.toolbar1 = NavigationToolbar2Tk(self.canvas1, window)
        self.toolbar1.config(background="whitesmoke")
        self.toolbar1._message_label.config(background="whitesmoke", fg="whitesmoke")
        self.toolbar1.update()
        self.toolbar1.place(relx=0.042, rely=0.05, relwidth=0.2, relheight=0.05)

        # График пика
        self.fig2 = Figure(facecolor="whitesmoke")
        self.ax2 = self.fig2.add_subplot(111)
        self.ax2.plot()
        self.ax2.set_ylim(0, 1)
        self.ax2.set_xlim(0, 1)
        self.fig2.tight_layout()
        self.canvas2 = FigureCanvasTkAgg(self.fig2, master=window)
        self.canvas2.draw()
        self.canvas2.get_tk_widget().configure(background="whitesmoke", highlightcolor='whitesmoke',
                                               highlightbackground='whitesmoke')
        self.canvas2.get_tk_widget().place(relx=0.6, rely=0.077, relwidth=0.4, relheight=0.63)
        self.toolbar2 = NavigationToolbar2Tk(self.canvas2, window)
        self.toolbar2.config(background="whitesmoke")
        self.toolbar2._message_label.config(background="whitesmoke", fg="whitesmoke")
        self.toolbar2.update()
        self.toolbar2.place(relx=0.627, rely=0.05, relwidth=0.2, relheight=0.05)

        # Ввод пользователем сопротивления электрода
        self.matR_button = ttk.Button(window, text="R электрода", state=["disabled"], command=self.mat_R)
        self.matR_button.place(relx=0.43, rely=0.95, relwidth=0.068, relheight=0.04)
        self.R_insert = Entry(justify=CENTER)
        self.R_insert.insert(0, 1)
        self.R_insert.place(relx=0.43, rely=0.91, relwidth=0.03, relheight=0.04)
        self.R_txt = ttk.Label(window, text="MOhm", background="whitesmoke", anchor="c")
        self.R_txt.place(relx=0.46, rely=0.91, relwidth=0.038, relheight=0.04)

        # Работа с комплексом
        # Радиокнопки для определения типа ПД
        self.frame3 = ttk.Frame(borderwidth=1, relief=SOLID)
        self.frame3.place(relx=0.6, rely=0.7, relwidth=0.08, relheight=0.08)
        self.state_var = StringVar(value="spont")
        self.evoked_radiobutton = ttk.Radiobutton(self.frame3, text="", value="evoked",
                                                  state=["disabled"])
        self.evoked_radiobutton.place(relx=0, rely=0)
        self.spont_radiobutton = ttk.Radiobutton(self.frame3, text="", value="spont",
                                                 state=["disabled"])
        self.spont_radiobutton.place(relx=0, rely=0.5)
        self.evoked_label = tk.Label(self.frame3, text="вызванный", relief=tk.FLAT, borderwidth=0,
                                     background="whitesmoke")
        self.evoked_label.place(relx=0.2, rely=0.05)
        self.spont_label = tk.Label(self.frame3, text="спонтанный", relief=tk.FLAT, borderwidth=0,
                                    background="whitesmoke")
        self.spont_label.place(relx=0.2, rely=0.55)

        # Кнопки анализа комплекса
        self.analys_in_button = ttk.Button(window, text="Анализ ПД", state=["disabled"], command=self.analys_in)
        self.analys_in_button.place(relx=0.70, rely=0.74, relwidth=0.07, relheight=0.04)
        self.start_point_button = ttk.Button(window, text="Метка", state=["disabled"], command=self.start_point)
        self.start_point_button.place(relx=0.78, rely=0.74, relwidth=0.07, relheight=0.04)

        # Таблица
        self.columns = (
            'N', 'rec', 'sp_ev', 'R', 'RMP', 'Amp', 'Overshoot', 'Tr', 'T10', 'T50', 'T90', 'lat_per',
            'Comments')
        self.table = ttk.Treeview(columns=self.columns, show='headings', selectmode="extended")
        self.table.place(relx=0.6, rely=0.8, relwidth=0.385, relheight=0.2)

        self.table.heading('N', text='N')
        self.table.heading('rec', text='rec')
        self.table.heading('sp_ev', text='sp_ev')
        self.table.heading('R', text='R')
        self.table.heading('RMP', text='RMP')
        self.table.heading('Amp', text='Amp')
        self.table.heading('Overshoot', text='Overshoot')
        self.table.heading('Tr', text='Tr')
        self.table.heading('T10', text='T10')
        self.table.heading('T50', text='T50')
        self.table.heading('T90', text='T90')
        self.table.heading('lat_per', text='lat_per')
        self.table.heading('Comments', text='Comments')

        self.table.column("#1", width=130)
        self.table.column("#2", width=45)
        self.table.column("#3", width=40)
        self.table.column("#4", width=45)
        self.table.column("#5", width=40)
        self.table.column("#6", width=40)
        self.table.column("#7", width=70)
        self.table.column("#8", width=40)
        self.table.column("#9", width=40)
        self.table.column("#10", width=40)
        self.table.column("#11", width=40)
        self.table.column("#12", width=50)
        self.table.column("#13", width=170)

        self.scrollbar_h = ttk.Scrollbar(orient=HORIZONTAL, command=self.table.xview)
        self.table.configure(xscroll=self.scrollbar_h.set)
        self.scrollbar_h.place(relx=0.6, rely=0.97, relwidth=0.385, relheight=0.03)
        self.scrollbar_v = ttk.Scrollbar(orient=VERTICAL, command=self.table.yview)
        self.table.configure(yscroll=self.scrollbar_v.set)
        self.scrollbar_v.place(relx=0.985, rely=0.8, relwidth=0.015, relheight=0.2)

        self.table.bind("<Button-1>", self.toggle_row)
        self.table.bind("<Double-1>", self.on_double_click)

        self.editable_column = 'Comments'
        self.current_edit = None
        self.h_scroll_binding_id = None
        self.v_scroll_binding_id = None

        # Кнопки работы с таблицей
        self.delete_str_button = ttk.Button(window, text="Удалить строку", command=self.delete_str)
        self.delete_str_button.place(relx=0.5, rely=0.88, relwidth=0.1, relheight=0.04)
        self.delete_table_button = ttk.Button(window, text="Удалить таблицу", command=self.delete_table)
        self.delete_table_button.place(relx=0.5, rely=0.92, relwidth=0.1, relheight=0.04)
        self.save_data_button = ttk.Button(window, text="Сохранить", command=self.save_data)
        self.save_data_button.place(relx=0.5, rely=0.96, relwidth=0.1, relheight=0.04)

        # --- Аттрибуты класса ---
        self.out = None  # Сохраняет загруженные данные
        self.wcp = None  # Сохраняет данные WCP
        self.zero_level = 0.0  # Начальное значение уровня нуля
        self.part_sig = None

    def open_file(self):
        # Загружаем последний открытый путь из файла
        try:
            with open('last_path.json', 'r') as f:
                last_path = json.load(f).get('last_path', '')
        except:
            last_path = ''

        """Открывает диалоговое окно выбора файла."""
        root = tk.Tk()
        root.withdraw()  # Скрывает главное окно
        fn = filedialog.askopenfilename(
            title="Выберите файл",
            initialdir=os.path.dirname(last_path) if last_path else '',
            filetypes=(("Файлы WCP", "*.wcp"), ("Все файлы", "*.*"))
        )
        # Если файл был выбран, сохраняем его путь
        if fn:
            with open('last_path.json', 'w') as f:
                json.dump({'last_path': fn}, f)
        root.destroy()  # Закрывает диалоговое окно выбора файла

        if fn:  # Если файл был выбран
            try:
                self.out = self.load_wcp_data(fn)  # Функция для обработки
                self.part_sig = None
            except:
                showerror(title="Ошибка", message="Ошибка при чтении файла.")
        else:  # Если файл не был выбран
            if self.out is None:  # Если ранее не было данных
                showwarning(title="Внимание", message="Файл не выбран и нет предыдущих данных.")
            else:
                fn = last_path
                showwarning(title='Внимание', message="Файл не выбран, используется предыдущий результат.")

        if self.out:
            self.enable_buttons()  # Включает кнопки при успешной загрузке данных
            self.file_name_label.config(text=os.path.basename(fn))
            self.n_records_label.config(text='/' + str(self.wcp['nr']))

            # Передает данные для визуализации
            self.visualizer = self.WCPVisualizer(self.out, self.canvas1, self.prev_rec_button, self.next_rec_button,
                                                 self.fig1, self.ax1, self.n_record_entry)

    def load_wcp_data(self, fn, recordings=None, debug=False):
        """Загружает WCP данные"""
        wcp = {}
        # Парсинг заголовочной части файла
        try:
            with open(fn, 'r') as fid:
                tline = fid.readline().strip()
                while len(tline) < 80:
                    in_indices = [i for i, char in enumerate(tline) if char == '=']
                    if in_indices:
                        in_index = in_indices[0]
                        pre = tline[:in_index].strip()
                        post = tline[in_index + 1:].strip()
                        in2_indices = [i for i, char in enumerate(post) if char == '=']
                        if in2_indices:
                            post = post[in2_indices[-1] + 1:].strip()
                        try:
                            if pre == 'VER':
                                wcp['version'] = float(post)
                            elif pre == 'CTIME':
                                wcp['ctime'] = post
                            elif pre == 'NC':
                                wcp['nc'] = int(post)
                            elif pre == 'NR':
                                wcp['nr'] = int(post)
                            elif pre == 'NBH':
                                wcp['nbh'] = int(post)
                            elif pre == 'NBA':
                                wcp['nba'] = int(post)
                            elif pre == 'NBD':
                                wcp['nbd'] = int(post)
                            elif pre == 'AD':
                                wcp['ad'] = int(post)
                            elif pre == 'ADCMAX':
                                wcp['adcmax'] = int(post)
                            elif pre == 'NP':
                                wcp['np'] = int(post)
                            elif pre == 'DT':
                                dt = post.replace(',', '.')
                                wcp['dt'] = float(dt)
                            elif pre == 'NZ':
                                wcp['nz'] = int(post)
                            elif pre == 'TU':  # Time units, единицы измерения времени
                                wcp['tu'] = post
                            elif pre == 'ID':
                                wcp['id'] = post
                                break
                            if pre[:2] == 'YN':  # Название канала X
                                N = int(pre[2])
                                if 'channel_info' not in wcp:
                                    wcp['channel_info'] = {}
                                if N + 1 not in wcp['channel_info']:
                                    wcp['channel_info'][N + 1] = {}
                                wcp['channel_info'][N + 1]['yn'] = post
                            elif pre[:2] == 'YU':  # Единицы измерения канала X
                                N = int(pre[2])
                                if 'channel_info' not in wcp:
                                    wcp['channel_info'] = {}
                                if N + 1 not in wcp['channel_info']:
                                    wcp['channel_info'][N + 1] = {}
                                wcp['channel_info'][N + 1]['yu'] = post
                            elif pre[:2] == 'YG':  # Коэффициент усиления канала X
                                N = int(pre[2])
                                post = post.replace(',', '.')
                                if 'channel_info' not in wcp:
                                    wcp['channel_info'] = {}
                                if N + 1 not in wcp['channel_info']:
                                    wcp['channel_info'][N + 1] = {}
                                wcp['channel_info'][N + 1]['yg'] = float(post)
                            elif pre[:2] == 'YZ':  # Нулевой уровень канала X
                                N = int(pre[2])
                                post = post.replace(',', '.')
                                if 'channel_info' not in wcp:
                                    wcp['channel_info'] = {}
                                if N + 1 not in wcp['channel_info']:
                                    wcp['channel_info'][N + 1] = {}
                                wcp['channel_info'][N + 1]['yz'] = float(post)
                            elif pre[:2] == 'YO':  #
                                N = int(pre[2])
                                post = post.replace(',', '.')
                                if 'channel_info' not in wcp:
                                    wcp['channel_info'] = {}
                                if N + 1 not in wcp['channel_info']:
                                    wcp['channel_info'][N + 1] = {}
                                wcp['channel_info'][N + 1]['yo'] = float(post)
                            elif pre[:2] == 'YR':
                                N = int(pre[2])
                                if 'channel_info' not in wcp:
                                    wcp['channel_info'] = {}
                                if N + 1 not in wcp['channel_info']:
                                    wcp['channel_info'][N + 1] = {}
                                wcp['channel_info'][N + 1]['yr'] = post
                        except ValueError as e:
                            showerror(title='Ошибка', message="Ошибка преобразования данных.")
                    tline = fid.readline().strip()
        except:
            showerror(title='Ошибка', message='Файл не найден')
            return None

        self.wcp = wcp

        if not recordings:
            recordings = list(range(1, wcp['nr'] + 1))
        else:
            recordings = [rec for rec in recordings if rec <= wcp['nr']]
            if not recordings:
                showerror(title='Ошибка', message='Записи не выбраны.')

        # Чтение бинарных данных
        RAB = [None] * len(recordings)
        DAB = [None] * len(recordings)
        rec_index = [0] * len(recordings)
        rab_pos = [0] * len(recordings)
        db_pos = [0] * len(recordings)

        try:
            with open(fn, 'rb') as fid:  # Открытие в бинарном виде для чтения блоков данных
                for i, rec_number in enumerate(recordings):
                    rec_index[i] = rec_number
                    rab_pos[i] = wcp['nbh'] + ((rec_number - 1) * (wcp['nba'] + wcp['nbd'])) * 512
                    fid.seek(rab_pos[i])

                    # Record Analysis Block
                    RAB[i] = {}
                    RAB[i]['status'] = fid.read(8).decode('ascii').strip()
                    RAB[i]['type'] = fid.read(4).decode('ascii').strip()

                    RAB[i]['group_no'] = struct.unpack('<f', fid.read(4))[0]
                    RAB[i]['time_rec'] = struct.unpack('<f', fid.read(4))[0]
                    RAB[i]['sampling_interval'] = struct.unpack('<f', fid.read(4))[0]

                    RAB[i]['max_pos_voltage'] = []
                    for j in range(wcp['nc']):
                        RAB[i]['max_pos_voltage'].append(struct.unpack('<f', fid.read(4))[0])

                    RAB[i]['marker'] = fid.read(16).decode('ascii').strip()

                    db_pos[i] = wcp['nbh'] + (wcp['nba'] + (rec_number - 1) * (wcp['nba'] + wcp['nbd'])) * 512
                    fid.seek(db_pos[i])

                    # Data Block
                    num_samples = wcp['nbd'] * 256 // 2
                    data = fid.read(wcp['nc'] * num_samples * 2)
                    DB = np.array(struct.unpack("<{}h".format(wcp['nc'] * num_samples), data)).reshape(
                        (wcp['nc'], num_samples), order='F')
                    DAB[i] = DB.astype(float)

        except:
            showerror(title='Ошибка', message='Файл не найден')
            return None

        T = np.arange(1, DAB[0].shape[1] + 1) * wcp['dt']  # [s]

        # Преобразование данных в физические величины
        S = [[] for _ in range(wcp['nc'])]
        num_recordings = len(recordings)
        for i in range(num_recordings):
            for j in range(wcp['nc']):
                converted_data = (RAB[i]['max_pos_voltage'][j] / (
                        wcp['adcmax'] * wcp['channel_info'][j + 1]['yg'] * 1000)) * \
                                 DAB[i][j, :]

                S[j].append(converted_data)

        # Организация выходной структуры
        out = {}
        out['S'] = S
        out['DAB'] = DAB
        out['T'] = T
        out['time'] = wcp['ctime']
        out['rec_index'] = rec_index
        out['channel_no'] = wcp['nc']
        out['t_interval'] = wcp['dt']
        out['file_name'] = fn

        out['channel_info'] = {}
        for i in range(1, wcp['nc'] + 1):
            out['channel_info'][i] = {}
            out['channel_info'][i]['unit'] = wcp['channel_info'][i]['yu']
            out['channel_info'][i]['name'] = wcp['channel_info'][i]['yn']

        out['rec_info'] = {}
        for i in range(len(recordings)):
            out['rec_info'][i] = {}
            out['rec_info'][i]['status'] = RAB[i]['status']
            out['rec_info'][i]['type'] = RAB[i]['type']
            out['rec_info'][i]['time_recorded'] = RAB[i]['time_rec']
            out['rec_info'][i]['group_no'] = RAB[i]['group_no']

        # Перестановка канала с данными о напряжении на первую позицию
        for key in out['channel_info']:
            if out['channel_info'][key]['unit'] == "mV":
                out['S'][0] = out['S'][key - 1]
        for j in range(wcp['nc']):
            out['S'][j] = [data * 1000 for data in out['S'][j]]  # Перевод из В в мВ

        return out

    def enable_buttons(self):
        """Включает кнопки после загрузки файла WCP."""
        if str(self.part_button.cget("state")) == tk.DISABLED:
            self.part_button.config(state=tk.NORMAL)
            self.prev_rec_button.config(state=tk.NORMAL)
            self.next_rec_button.config(state=tk.NORMAL)
            self.zoom_in_button.config(state=tk.NORMAL)
            self.zerolevel_button.config(state=tk.NORMAL)
            self.x10_radiobutton.config(state=tk.NORMAL)
            self.inverse_radiobutton.config(state=tk.NORMAL)

        if str(self.x1_radiobutton.cget("state")) == tk.NORMAL:
            self.multi_var = StringVar(value="x1")
            self.x1_radiobutton.config(variable=self.multi_var, state=tk.DISABLED)
            self.x10_radiobutton.config(variable=self.multi_var, state=tk.NORMAL)

        if str(self.original_radiobutton.cget("state")) == tk.NORMAL:
            self.sign_var = StringVar(value="+")
            self.original_radiobutton.config(variable=self.sign_var, state=tk.DISABLED)
            self.inverse_radiobutton.config(variable=self.sign_var, state=tk.NORMAL)

    class WCPVisualizer:
        def __init__(self, out, canvas1, prev_rec, next_rec, fig1, ax1, n_record):
            if not out or 'S' not in out:
                showerror(title='Ошибка', message='Нет данных для визуализации.')
                return
            self.canvas1 = canvas1
            self.fig = fig1
            self.ax = ax1
            self.root = self.canvas1.get_tk_widget().master
            self.n_record = n_record
            self.out = out

            # Получаем список всех комбинаций записи и канала, у которых unit = mV
            self.plot_data = []
            for rec_index in range(len(out['S'][0])):
                for channel_index in range(len(out['S'])):
                    if out['channel_info'][channel_index + 1]['unit'] == 'mV':
                        self.plot_data.append((rec_index, channel_index, out['S'][channel_index][rec_index]))

            if not self.plot_data:
                showerror(title='Ошибка', message='Нет данных для отображения с unit mV.')
                return
            self.current_index = 0
            self.canvas = self.canvas1
            self.update_plot()

            # Кнопки навигации
            prev_rec.config(command=self.prev_plot)
            next_rec.config(command=self.next_plot)

            # Привязываем событие нажатия клавиши Enter к полю ввода
            self.n_record.bind('<Return>', self.go_to_record)

        def update_plot(self):
            self.ax.clear()
            rec_index, channel_index, channel_data = self.plot_data[self.current_index]
            self.ax.plot(self.out['T'], channel_data)
            self.ax.grid(True)
            self.ax.set_xlim(min(self.out['T']), max(self.out['T']))
            if isinstance(channel_data, np.ndarray):
                self.ax.set_ylim(min(channel_data), max(channel_data))
                ymin = min(channel_data)
                ymax = max(channel_data)
                yrange = ymax - ymin
                padding = 0.1 * yrange  # 10% отступ
                self.ax.set_ylim(ymin - padding, ymax + padding)  # Устанавливаем пределы с отступом
            else:
                channel_data_np = np.array(channel_data)
                self.ax.set_ylim(min(channel_data_np), max(channel_data_np))
                ymin = min(channel_data_np)
                ymax = max(channel_data_np)
                yrange = ymax - ymin
                padding = 0.1 * yrange  # 10% отступ
                self.ax.set_ylim(ymin - padding, ymax + padding)  # Устанавливаем пределы с отступом

            self.fig.canvas.draw()
            self.canvas1.get_tk_widget().update_idletasks()

            self.n_record.delete(0, tk.END)
            self.n_record.insert(0, str(self.current_index + 1))

        def next_plot(self):
            self.current_index = (self.current_index + 1) % len(self.plot_data)
            self.update_plot()

        def prev_plot(self):
            self.current_index = (self.current_index - 1) % len(self.plot_data)  # Обработка отрицательных индексов
            self.update_plot()

        def go_to_record(self, event):
            try:
                new_index = int(self.n_record.get()) - 1  # Получаем номер записи от пользователя
                if new_index < 0:
                    new_index = 0  # если номер меньше 1, устанавливаем на первую запись
                elif new_index >= len(self.plot_data):
                    new_index = len(self.plot_data) - 1  # если номер больше доступных, переходим на последнюю запись
                self.current_index = new_index
                self.update_plot()
            except ValueError:
                showerror(title='Ошибка', message='Введите корректное целое число.')

    def x10(self):
        if not self.out or 'S' not in self.out:
            showerror(title='Ошибка', message='Нет данных для изменения.')
            return

        for i in range(len(self.visualizer.plot_data)):
            rec_index, channel_index, channel_data = self.visualizer.plot_data[i]
            updated_data = [data / 10 for data in channel_data]
            self.visualizer.plot_data[i] = (rec_index, channel_index, updated_data)

        self.visualizer.update_plot()
        self.multi_var = StringVar(value="x10")
        self.x10_radiobutton.config(state=tk.DISABLED)
        self.x1_radiobutton.config(state=tk.NORMAL)

    def x1(self):
        if not self.out or 'S' not in self.out:
            showerror(title='Ошибка', message='Нет данных для изменения.')
            return

        for i in range(len(self.visualizer.plot_data)):
            rec_index, channel_index, channel_data = self.visualizer.plot_data[i]
            updated_data = [data * 10 for data in channel_data]
            self.visualizer.plot_data[i] = (rec_index, channel_index, updated_data)

        self.visualizer.update_plot()
        self.multi_var = StringVar(value="x1")
        self.x1_radiobutton.config(state=tk.DISABLED)
        self.x10_radiobutton.config(state=tk.NORMAL)

    def inverse(self):
        if not self.out or 'S' not in self.out:
            showerror(title='Ошибка', message='Нет данных для изменения.')
            return

        for i in range(len(self.visualizer.plot_data)):
            rec_index, channel_index, channel_data = self.visualizer.plot_data[i]
            updated_data = [-data for data in channel_data]
            self.visualizer.plot_data[i] = (rec_index, channel_index, updated_data)

        self.visualizer.update_plot()
        self.inverse_radiobutton.config(state=tk.DISABLED)
        self.original_radiobutton.config(state=tk.NORMAL)

    def original(self):
        if not self.out or 'S' not in self.out:
            showerror(title='Ошибка', message='Нет данных для изменения.')
            return

        for i in range(len(self.visualizer.plot_data)):
            rec_index, channel_index, channel_data = self.visualizer.plot_data[i]
            updated_data = [-data for data in channel_data]
            self.visualizer.plot_data[i] = (rec_index, channel_index, updated_data)

        self.visualizer.update_plot()
        self.original_radiobutton.config(state=tk.DISABLED)
        self.inverse_radiobutton.config(state=tk.NORMAL)

    def part_graph(self):
        """Отображает окно выбора сегментов записи."""
        try:
            new_window = tk.Toplevel(self.window)  # Используем self.window
            new_window.title("Выбор сегментов записи")
            new_window.geometry('1200x420+50+300')
            new_window["bg"] = "whitesmoke"
            self.additional_windows.append(new_window)  # Добавляем окно в список
            new_window.protocol("WM_DELETE_WINDOW", lambda: new_window.destroy())

            # Создаем Figure и Axes
            fig3 = plt.figure(2)
            fig3.set_facecolor("whitesmoke")
            fig3.clf()
            fig3.set_size_inches(10, 2)  # (ширина, высота)
            ax = fig3.add_axes([0.05, 0.15, 0.925, 0.83])

            # Создаем Canvas и Toolbar
            canvas = FigureCanvasTkAgg(fig3, master=new_window)
            canvas.draw()
            canvas.get_tk_widget().pack(side=TOP, fill=BOTH, expand=1)

            toolbar_frame = Frame(new_window, bg="whitesmoke")
            toolbar_frame.pack(side=TOP, fill=X)

            try:
                toolbar3 = NavigationToolbar2Tk(canvas, toolbar_frame)
                toolbar3.update()
                toolbar3._message_label.destroy()
                toolbar3.pack(side=RIGHT)
                canvas._tkcanvas.pack(side=BOTTOM, fill=BOTH, expand=1)
            except:
                showerror(title='Ошибка', message="Ошибка при создании NavigationToolbar.")

            # Создаем Comboboxes
            N = list(range(1, self.wcp['nr'] + 1))
            np1_label = ttk.Label(new_window, text="Начало:", background="whitesmoke")
            np1_label.place(x=100, rely=0.92)  # Чуть ниже слайдеров

            NP1 = ttk.Combobox(new_window, values=N,
                               state="readonly")  # state="readonly" делает невозможным ввод текста
            NP1.place(x=150, rely=0.92, width=60, height=20)
            if N:
                NP1.set(1)  # Значение по умолчанию
            else:
                NP1.set('')

            np2_label = ttk.Label(new_window, text="Конец:", background="whitesmoke")
            np2_label.place(x=250, rely=0.92)

            NP2 = ttk.Combobox(new_window, values=N, state="readonly")
            NP2.place(x=300, rely=0.92, width=60, height=20)
            if N:
                NP2.set(1)  # Значение по умолчанию
            else:
                NP2.set('')

            NP1.bind("<<ComboboxSelected>>", lambda event: self.change_NP1Value(NP1, NP2))
            NP2.bind("<<ComboboxSelected>>", lambda event: self.change_NP2Value(NP1, NP2))

            select_button = tk.Button(new_window, text="Выбрать",
                                      command=lambda: self.plot_part(ax, fig3, NP1,
                                                                     NP2))
            select_button.place(relx=0.02, rely=0.92, width=60, height=20)

            self.plot_part(ax, fig3, NP1, NP2)

        except:
            showerror(title='Ошибка', message="Ошибка в part_graph.")

    def change_NP1Value(self, NP1, NP2):
        try:
            np1_value = int(NP1.get())
            np2_value = int(NP2.get())

            if np1_value > np2_value:
                NP2.set(np1_value)
        except ValueError:
            pass

    def change_NP2Value(self, NP1, NP2):
        try:
            np1_value = int(NP1.get())
            np2_value = int(NP2.get())
            if np2_value < np1_value:
                NP1.set(np2_value)
        except ValueError:
            pass

    def plot_part(self, ax, fig3, NP1, NP2):
        """Отображает выбранные сегменты на графике."""
        try:
            np1_value = int(NP1.get())
            np2_value = int(NP2.get())
        except ValueError:
            print("Ошибка: Некорректные значения в списках выбора.")
            return

        # Преобразуем в массив NumPy, если это список
        if isinstance(self.out['S'][0], list):
            self.out['S'][0] = np.array(self.out['S'][0])

        # Инициализируем данные
        full_sig = self.out['S'][0].flatten()  # Преобразуем в одномерный массив
        xS, yS = self.out['S'][0].shape

        start_index = (np1_value - 1) * yS  # Индекс начала сегмента
        end_index = np2_value * yS  # Индекс конца сегмента

        # Безопасность индексов
        if end_index > len(full_sig):
            end_index = len(full_sig)
        if start_index < 0:
            start_index = 0

        ax.cla()  # Очищаем оси
        ax.plot(full_sig[start_index:end_index])

        if len(full_sig[start_index:end_index]) > 0:  # Проверка, что слайс не пустой
            ymin = min(full_sig[start_index:end_index])
            ymax = max(full_sig[start_index:end_index])
            yrange = ymax - ymin
        else:
            ymin = 0
            ymax = 1
            yrange = 1

        # Добавляем сетку
        ax.grid(True)
        ax.margins(x=0)  # убираем отступы по оси x
        ax.autoscale(tight=True)
        ax.axis('tight')
        ax.set_xticks([])

        padding = 0.1 * yrange  # 10% отступ
        ax.set_ylim(ymin - padding, ymax + padding)
        fig3.canvas.draw_idle()

    def zerolevel(self):
        if not self.out or 'S' not in self.out:
            showerror(title='Ошибка', message='Нет данных для установки уровня нуля.')
            return

        # Отключаем кнопку zerolevel
        self.zerolevel_button.config(state=tk.DISABLED)

        # Меняем курсор на перекрестие
        self.canvas1.get_tk_widget().configure(cursor='crosshair')

        # Привязываем функцию к движению мыши и клику на графике
        self.cid_motion = self.canvas1.mpl_connect('motion_notify_event', self.update_cursor_position_1)
        self.cid_click = self.canvas1.mpl_connect('button_press_event', self.set_zero_level)
        self.cursor_line_vertical = None
        self.cursor_line_horizontal = None

    def update_cursor_position_1(self, event):
        """Обновляет положение курсора и отображает вертикальную и горизонтальную линии."""
        if event.inaxes is self.ax1:
            # Удаляем старые линии
            if self.cursor_line_vertical:
                self.cursor_line_vertical.remove()
            if self.cursor_line_horizontal:
                self.cursor_line_horizontal.remove()

            # Рисуем новые линии
            self.cursor_line_vertical = self.ax1.axvline(x=event.xdata, color='red', linestyle='-',
                                                         linewidth=0.5)
            self.cursor_line_horizontal = self.ax1.axhline(y=event.ydata, color='red', linestyle='-',
                                                           linewidth=0.5)

            self.canvas1.draw_idle()

    def set_zero_level(self, event):
        if event.inaxes is not self.ax1:  # Проверяем, что клик был на нужном графике
            return

        # Проверяем, что щелчок был на линии графика
        contains, _ = self.ax1.lines[0].contains(event)
        if not contains:
            showerror(title='Ошибка', message='Щелчок должен быть на линии графика.')
            return

        x1 = int(event.xdata * 10000)
        y1 = event.ydata

        # Получаем текущий канал данных для отображения
        rec_index, channel_index, self.part_sig = self.visualizer.plot_data[self.visualizer.current_index]

        # Расчет нового уровня нуля
        self.zero_level = y1
        self.part_sig -= self.zero_level

        # Обновляем данные в visualizer.plot_data
        self.visualizer.plot_data[self.visualizer.current_index] = (rec_index, channel_index, self.part_sig)

        # Перерисовываем график
        self.visualizer.update_plot()

        # Отключаем обработчики событий после выполнения коррекции
        self.disconnect_cursor_1()

    def disconnect_cursor_1(self):
        """Отключает обработчики событий и убирает линии курсора."""
        if self.cid_click is not None:
            self.ax1.figure.canvas.mpl_disconnect(self.cid_click)
            self.cid_click = None
        if self.cid_motion is not None:
            self.ax1.figure.canvas.mpl_disconnect(self.cid_motion)
            self.cid_motion = None

        if self.cursor_line_vertical:
            self.cursor_line_vertical = None

        if self.cursor_line_horizontal:
            self.cursor_line_horizontal = None

        self.canvas1.get_tk_widget().configure(cursor='arrow')
        self.canvas1.draw_idle()

        # Включаем кнопку zerolevel обратно
        self.zerolevel_button.config(state=tk.NORMAL)

    def zoom_sig(self):
        if not self.out or 'S' not in self.out:
            showerror(title='Ошибка', message='Нет данных для приближения.')
            return
        # Отключаем кнопку zoom_in
        self.zoom_in_button.config(state=tk.DISABLED)
        # Меняем курсор на перекрестие
        self.canvas1.get_tk_widget().configure(cursor='crosshair')

        # Привязываем функцию к движению мыши и клику на графике
        self.cid_motion = self.canvas1.mpl_connect('motion_notify_event', self.update_cursor_position_2)
        self.cid_click = self.canvas1.mpl_connect('button_press_event', self.set_zoom_area)
        self.cursor_line_vertical = None
        self.cursor_line_horizontal = None
        self.markers = []
        self.first_click = True

    def update_cursor_position_2(self, event):
        """Обновляет положение курсора и отображает вертикальную и горизонтальную линии."""
        if event.inaxes is self.ax1:
            # Удаляем старые линии
            if self.cursor_line_vertical:
                self.cursor_line_vertical.remove()
            if self.cursor_line_horizontal:
                self.cursor_line_horizontal.remove()

            # Рисуем новые линии
            self.cursor_line_vertical = self.ax1.axvline(x=event.xdata, color='green', linestyle='-',
                                                         linewidth=0.8)
            self.cursor_line_horizontal = self.ax1.axhline(y=event.ydata, color='green', linestyle='-',
                                                           linewidth=0.8)
            self.canvas1.draw_idle()

    def set_zoom_area(self, event):
        if event.inaxes is not self.ax1:  # Проверяем, что клик был на нужном графике
            return

        # Проверяем, что щелчок был на линии графика
        contains, _ = self.ax1.lines[0].contains(event)
        if not contains:
            showerror(title='Ошибка', message='Щелчок должен быть на линии графика.')
            return
        # Получаем текущий канал данных для отображения
        rec_index, channel_index, self.part_sig = self.visualizer.plot_data[self.visualizer.current_index]

        # Получаем x и y координаты клика
        x1 = event.xdata
        y1 = event.ydata

        # Добавляем маркер в месте клика
        marker = self.ax1.plot(x1, y1, marker='v', color='green', markersize=7)[0]  # ^ - треугольник
        self.markers.append(marker)  # Сохраняем маркер для удаления позже
        self.canvas1.draw_idle()

        if self.first_click:
            # Первый клик: сохраняем координату x и y
            self.start_x = x1
            self.start_y = y1
            self.first_click = False
        else:
            # Второй клик: сохраняем координату x и y
            self.end_x = x1
            self.end_y = y1

            #  логика для построения нового графика
            self.plot_selected_range()

            # Отключаем обработчики событий после выполнения коррекции
            self.disconnect_cursor_2()
            self.st_point = 0

        self.complex_name_label.config(text=('rec_', str(rec_index + 1)))

    def plot_selected_range(self):
        """     Отображает выделенный участок графика на self.ax2.  """
        # Проверяем, что координаты кликов установлены
        if self.start_y is None or self.end_y is None:
            showerror(title='Ошибка', message='Необходимо выбрать два значения на графике.')
            return

        s_values = self.part_sig
        t_values = self.out['T']

        if len(s_values) != len(t_values):
            showerror(title='Ошибка', message='Длина out["S"] и out["T"] не совпадают.')
            return None

        for i in range(len(s_values)):
            self.ratio_dict[t_values[i]] = s_values[i]

        # Определяем границы диапазона x
        start_x = min(self.start_x, self.end_x)
        end_x = max(self.start_x, self.end_x)

        # Извлекаем значения y, соответствующие выбранному диапазону x
        x_values = sorted(self.ratio_dict.keys())
        self.y_values = []

        for x in x_values:
            if start_x <= x <= end_x:
                self.y_values.append(self.ratio_dict[x])

        # Создаем массив x_values_selected для графика:
        self.x_values_selected = [x for x in x_values if start_x <= x <= end_x]

        self.complex_data = {}
        for i in range(len(self.y_values)):
            self.complex_data[self.x_values_selected[i]] = self.y_values[i]

        # Очищаем предыдущий график
        self.ax2.clear()

        # Строим график
        self.ax2.plot(self.x_values_selected, self.y_values)
        # Добавляем сетку
        self.ax2.grid(True)
        # Убираем отступы по бокам
        self.ax2.margins(x=0)  # убираем отступы по оси x

        # Обновляем canvas2
        self.fig2.tight_layout()
        self.canvas2.draw()

    def disconnect_cursor_2(self):
        """     Отключает обработчики событий и убирает линии курсора.  """
        if self.cid_click is not None:
            self.ax1.figure.canvas.mpl_disconnect(self.cid_click)
            self.cid_click = None
        if self.cid_motion is not None:
            self.ax1.figure.canvas.mpl_disconnect(self.cid_motion)
            self.cid_motion = None
        if self.cursor_line_vertical:
            self.cursor_line_vertical.remove()
            self.cursor_line_vertical = None
        if self.cursor_line_horizontal:
            self.cursor_line_horizontal.remove()
            self.cursor_line_horizontal = None
        self.canvas1.get_tk_widget().configure(cursor='arrow')
        self.canvas1.draw_idle()

        # Включаем кнопку zoom_in обратно
        self.zoom_in_button.config(state=tk.NORMAL)
        self.matR_button.config(state=tk.NORMAL)
        self.evoked_radiobutton.config(variable=self.state_var)
        self.spont_radiobutton.config(variable=self.state_var)
        self.analys_in_button.config(state=tk.NORMAL)
        self.start_point_button.config(state=tk.NORMAL)

    def mat_R(self):
        """     Добавление сопротивления электрода в таблицу.   """
        try:
            R = int(self.R_insert.get())  # Получаем значение сопротивления из поля ввода
        except ValueError:
            showerror(title='Ошибка', message='Неверный формат сопротивления.')
            return

        # Проверяем, есть ли строки в таблице
        if self.table.get_children():  # Есть ли строки в таблице
            last_row = self.table.get_children()[-1]  # Получаем ID последней строки
            values = self.table.item(last_row, 'values')  # Получаем значения последней строки

            # Проверяем, есть ли какие-либо данные в последней строке (кроме R)
            # и что столбец R пуст
            if any(values[i] != '' for i in range(len(values)) if i != 2) and values[
                3] == '':  # Проверяем столбцы, кроме R (индекс 2)
                values = list(values)  # Преобразуем tuple в list, чтобы изменить его
                values[3] = str(R)  # Обновляем значение R
                self.table.item(last_row, values=tuple(values))  # Обновляем строку в таблице
            else:
                showerror(title='Ошибка', message="В последней строке таблицы уже есть данные R. "
                                                  "R не добавлен.")
        else:
            showerror(title='Ошибка', message="В таблице нет строк. R не добавлен.")

    def start_point(self):
        """    Функция для выбора начальной точки стимуляции (артефакта).    """
        if self.complex_data is None:  # Проверяем, что complex_data не пустой
            showerror(title='Ошибка', message='Необходимо загрузить данные.')
            return

        # Если уже есть линия, удаляем её
        if hasattr(self, 'st_point_line'):
            if self.st_point_line in self.ax2.lines:  # Проверяем наличие линии на графике
                self.st_point_line.remove()
            del self.st_point_line  # Удаляем атрибут

        # Подключаем функцию onclick к событию клика мыши на рисунке
        self.cid_click = self.fig2.canvas.mpl_connect('button_press_event', self.onclick)
        self.cid_motion = self.fig2.canvas.mpl_connect('motion_notify_event',
                                                       self.on_motion)  # Подключаем для движения мыши

    def on_motion(self, event):
        """    Обработчик движения мыши для отображения временных линий.    """
        if event.inaxes == self.ax2:
            x = event.xdata
            y = event.ydata
            if x is not None and y is not None:
                # Удаляем предыдущие временные линии
                if hasattr(self, 'temp_vline'):
                    self.temp_vline.remove()
                if hasattr(self, 'temp_hline'):
                    self.temp_hline.remove()

                # Рисуем новые временные линии
                self.temp_vline = self.ax2.axvline(x=x, linestyle=':', color='gray', linewidth=0.5)
                self.temp_hline = self.ax2.axhline(y=y, linestyle=':', color='gray', linewidth=0.5)

                self.fig2.canvas.draw_idle()

    def onclick(self, event):
        """    Обработчик события клика мыши для start_point.    """
        if event.inaxes == self.ax2:
            x = event.xdata
            if x is not None:
                # Удаляем временные линии
                if hasattr(self, 'temp_vline'):
                    self.temp_vline.remove()
                    del self.temp_vline
                if hasattr(self, 'temp_hline'):
                    self.temp_hline.remove()
                    del self.temp_hline

                # Устанавливаем st_point
                self.st_point = x

                # Если линия уже существует, удаляем её
                if hasattr(self, 'st_point_line'):
                    if self.st_point_line in self.ax2.lines:  # Проверяем наличие линии на графике
                        self.st_point_line.remove()

                # Рисуем постоянную вертикальную линию
                self.st_point_line = self.ax2.axvline(x=self.st_point, linestyle='--', color='black', linewidth=0.5)

                self.fig2.canvas.draw()

                # Отключаем обработчик событий движения мыши и клика
                self.fig2.canvas.mpl_disconnect(self.cid_click)
                self.fig2.canvas.mpl_disconnect(self.cid_motion)

                self.state_var = StringVar(value='evoked')
                self.evoked_radiobutton.config(variable=self.state_var)
                self.spont_radiobutton.config(variable=self.state_var)
            else:
                print("Клик был вне области данных.")
        else:
            print("Клик был не на целевых осях.")

    def max_peak(self):
        """
        Находит положение пика в данных complex_data.
        Если self.st_point == 0, ищет максимум 2 пика от начала.
        Если self.st_point != 0, ищет первый пик в окрестности self.st_point, а второй - дальше.
        """
        if self.y_values is None or self.st_point is None:
            showerror(title='Ошибка', message='Необходимо сначала выбрать начальную точку.')
            return None

        y = np.array(self.y_values)
        x = np.array(self.x_values_selected)
        start_index = np.argmin(np.abs(x - self.st_point))

        if self.st_point == 0:
            # Случай 1: self.st_point == 0, ищем два пика от начала
            data_subset = y[start_index:]
            peaks, _ = find_peaks(data_subset, height=5, prominence=7, distance=2000)
            peaks = peaks[:2]
            xpd = peaks + start_index  # Смещаем индексы пиков

        else:
            # Случай 2: self.st_point != 0, ищем первый пик рядом с self.st_point, второй - дальше.
            # Поиск первого пика в окрестности self.st_point (например, +- 10% от длины x)
            search_range = int(len(x) * 0.01)  # Определяем окрестность
            start_search = max(0, start_index - search_range)  # Начало поиска не может быть меньше 0
            end_search = min(len(x), start_index + search_range)  # Конец поиска не может быть больше len(x)

            data_subset_1 = y[start_search:end_search]
            x_subset_1 = x[start_search:end_search]
            if str(self.multi_var.get()) == 'x10':
                peaks_1, _ = find_peaks(data_subset_1, height=0.4, prominence=0.01, distance=2000)
            else:
                peaks_1, _ = find_peaks(data_subset_1, height=1, prominence=3, distance=2000)

            # Если найдены пики в окрестности, берем первый.  Если нет - что-то идет не так, возвращаем None
            if len(peaks_1) > 0:
                peak_1_index_relative = peaks_1[0]  # Индекс относительно data_subset_1
                peak_1_index_absolute = start_search + peak_1_index_relative  # Индекс в исходном y
                peak_indices = [peak_1_index_absolute]

                # Поиск второго пика за пределами окрестности первого.
                data_subset_2 = y[end_search:]
                x_subset_2 = x[end_search:]
                peaks_2, _ = find_peaks(data_subset_2, height=1, prominence=3, distance=2000)
                if len(peaks_2) > 0:
                    peak_2_index_relative = peaks_2[0]
                    peak_2_index_absolute = end_search + peak_2_index_relative
                    peak_indices.append(peak_2_index_absolute)  # Добавляем к найденным пикам
                xpd = np.array(peak_indices)  # Преобразуем в NumPy массив
            else:
                showerror(title='Ошибка', message='Не найдено пиков в окрестности self.st_point.')
                return None

        xpd = [int(i) for i in xpd.tolist()]
        return xpd

    def analys_in(self):
        """ Анализ ПД.  """
        if not hasattr(self, 'y_values') or self.y_values is None:
            showerror(title='Ошибка', message='Необходимо загрузить данные.')
            return

        RMP = np.mean(self.y_values[:500])

        self.y_values = [y - RMP for y in self.y_values]  # Вычитаем RMP из каждого значения y

        # Корректировка дрейфа базовой линии с помощью линейной зависимости
        last = np.mean(self.y_values[(len(self.y_values) - 500):(len(self.y_values) - 5)])
        b = last / len(self.y_values)
        for ij in range(len(self.y_values)):
            self.y_values[ij] = self.y_values[ij] - b * (ij + 1)

        x_artf = 0
        xpd = self.max_peak()

        try:
            # если определился пик артефакта стимуляции
            if isinstance(xpd, list) and len(xpd) > 1:  # Проверка, что xpd - список и содержит более одного элемента
                x_artf = xpd[0]
                xpd = xpd[1]
                self.state_var = StringVar(value='evoked')
                self.evoked_radiobutton.config(variable=self.state_var)
                self.spont_radiobutton.config(variable=self.state_var)
            else:
                self.state_var = StringVar(value='spont')
                self.evoked_radiobutton.config(variable=self.state_var)
                self.spont_radiobutton.config(variable=self.state_var)
                xpd = xpd[0]

            dt = self.out['t_interval'] * 1000
            Apd = self.y_values[xpd]
            peak_x = self.x_values_selected[xpd]
            peak_y = self.y_values[xpd]

            # поиск Tr
            i = xpd
            Apd01 = 0.1 * Apd
            while self.y_values[i] > Apd01:
                i = i - 1
            i10 = i
            tr = (xpd - i) * dt

            # T10
            i = xpd
            Apd09 = 0.9 * Apd
            while self.y_values[i] > Apd09:
                i = i + 1
            i10s = i
            t10 = (i - xpd) * dt

            # T50
            i = xpd
            Apd05 = 0.5 * Apd
            while self.y_values[i] > Apd05:
                i = i + 1
            i50 = i
            t50 = (i - xpd) * dt

            # T90
            i = xpd
            while self.y_values[i] > Apd01:
                i = i + 1
            i90 = i
            t90 = (i - xpd) * dt

            # ------ вывод графика ------

            self.ax2.clear()
            self.ax2.plot(self.x_values_selected, self.y_values)  # Отображаем значения x и y
            self.ax2.plot(peak_x, peak_y, "rv", markerfacecolor='b',
                          label=u'Пик ПД')  # используем self.y_values

            if hasattr(self, 'st_point') and self.st_point != 0:  # Проверяем, был ли установлен st_point
                self.ax2.axvline(x=self.st_point, linestyle='--', color='black', linewidth=0.5, label=u'Метка')

            self.ax2.plot(self.x_values_selected[i10], self.y_values[i10], 'rv', markerfacecolor='g',
                          label=u'10% рост ПД')
            self.ax2.plot(self.x_values_selected[i10s], self.y_values[i10s], 'rv', markerfacecolor='y',
                          label=u'10% спад ПД')
            self.ax2.plot(self.x_values_selected[i50], self.y_values[i50], 'rv', markerfacecolor=[0.9, 0.6, 0.3],
                          label=u'50% спад ПД')
            self.ax2.plot(self.x_values_selected[i90], self.y_values[i90], 'rv', markerfacecolor='r',
                          label=u'90% спад ПД')

            if x_artf != 0:  # если есть артефакт
                xx_artf = self.x_values_selected[x_artf]
                yy_artf = self.y_values[x_artf]
                # поиск 10% от артефакта
                i10_artf = yy_artf * 0.1
                i = x_artf
                while self.y_values[i] > i10_artf:
                    i = i - 1
                i_artf = i
                lat_per = (i10 - i) * dt

                self.ax2.plot(self.x_values_selected[x_artf], self.y_values[x_artf], 'ro', markerfacecolor='b',
                              label=u'Пик артефакта')  # используем self.y_values
                self.ax2.plot(self.x_values_selected[i_artf], self.y_values[i_artf], 'ro', markerfacecolor='g',
                              label=u'10% артефакта')  # используем self.y_values
            else:
                lat_per = None

            self.ax2.legend(fontsize='small')
            self.ax2.margins(x=0)
            self.ax2.grid(True)
            self.fig2.canvas.draw()  # перерисовываем график

            sp_ev_state = str(self.state_var.get())
            overshoot = RMP + Apd
            # Добавляем информацию в таблицу
            values = (
                str(self.file_name_label['text']),  # N
                str(self.complex_name_label['text']),  # rec
                sp_ev_state,  # sp_ev
                '',  # R
                RMP,  # RMP
                Apd,  # Amp
                overshoot,  # Overshoot
                tr,  # Tr
                t10,  # T10
                t50,  # T50
                t90,  # T90
                lat_per,  # lat_per
                ''  # Comments
            )
            self.table.insert('', tk.END, values=values)

        except:
            showerror(title='Ошибка', message='Пики не найдены.')

    def on_double_click(self, event):
        """Handles double click on the table.  Only allows editing the Comments column."""
        if self.current_edit:
            return

        item = self.table.identify_row(event.y)
        column = self.table.identify_column(event.x)

        if not item or column != '#13':
            return

        # Unselect the row before editing
        self.table.selection_remove(item)

        x, y, width, height = self.get_cell_rectangle(item, column)

        try:
            value = self.table.item(item, 'values')[12]
        except IndexError:
            value = ""

        entry = tk.Entry(self.table,  # Entry создается на Treeview
                         background="white",
                         foreground="black",
                         relief=tk.SOLID,
                         borderwidth=1
                         )

        entry.insert(0, value)
        # Используем координаты относительно Treeview
        entry.place(x=x, y=y, width=width, height=height)
        entry.lift()
        entry.focus_set()

        self.current_edit = (item, entry, value)

        # Отключаем прокрутку
        self.disable_scroll()

        self.table.bind("<<TreeviewSelect>>", self.unselect_row_during_edit)
        entry.bind("<FocusOut>", self.on_edit_finish)
        entry.bind("<Return>", self.on_edit_finish)
        entry.bind("<Escape>", self.on_edit_cancel)

    def disable_scroll(self):
        """Отключает прокрутку Treeview."""
        self.h_scroll_binding_id = self.scrollbar_h.bind("<B1-Motion>", lambda event: "break")
        self.v_scroll_binding_id = self.scrollbar_v.bind("<B1-Motion>", lambda event: "break")
        self.table.configure(xscrollcommand=None, yscrollcommand=None)

    def enable_scroll(self):
        """Включает прокрутку Treeview."""
        self.scrollbar_h.unbind("<B1-Motion>", self.h_scroll_binding_id)
        self.scrollbar_v.unbind("<B1-Motion>", self.v_scroll_binding_id)
        self.table.configure(xscrollcommand=self.scrollbar_h.set, yscrollcommand=self.scrollbar_v.set)

    def unselect_row_during_edit(self, event):
        """Отменяет выбор строки при редактировании."""
        if self.current_edit:
            for item in self.table.selection():
                self.table.selection_remove(item)

    def on_edit_finish(self, event=None):
        """Вызывается при завершении редактирования."""
        if not self.current_edit:
            return
        item, entry, old_value = self.current_edit
        new_value = entry.get()

        if new_value != old_value:
            values = list(self.table.item(item, 'values'))
            try:
                values[13] = new_value
            except IndexError:
                while len(values) < 13:
                    values.append("")
                values[12] = new_value

            self.table.item(item, values=values)
        entry.destroy()
        self.current_edit = None

        # Включаем прокрутку
        self.enable_scroll()

        self.table.unbind("<<TreeviewSelect>>")

    def on_edit_cancel(self, event=None):
        """Вызывается, когда редактирование закончено."""
        if not self.current_edit:
            return
        item, entry, old_value = self.current_edit
        entry.destroy()
        self.current_edit = None

        # Включаем прокрутку
        self.enable_scroll()

        self.table.unbind("<<TreeviewSelect>>")

    def get_cell_rectangle(self, item, column):
        """Возвращает местоположение и размер ячейки."""
        bbox = self.table.bbox(item, column)
        x = bbox[0]
        y = bbox[1]
        width = bbox[2]
        height = bbox[3]
        return x, y, width, height

    def toggle_row(self, event):
        """Переключает состояние выделения строки при щелчке мыши."""
        item = self.table.identify_row(event.y)
        if item:
            if item in self.table.selection():
                self.table.selection_remove(item)
            else:
                self.table.selection_add(item)

    def delete_str(self):
        selected_items = self.table.selection()

        for item in selected_items:
            self.table.delete(item)

    def delete_table(self):
        for item in self.table.get_children():
            self.table.delete(item)

    def save_data(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            workbook = xlsxwriter.Workbook(filename)
            sheet = workbook.add_worksheet("Table Data")

            # Записывает заголовки
            for col, header in enumerate(self.table['columns']):
                sheet.write(0, col, self.table.heading(header)['text'])

            # Записывает данные
            for row, item in enumerate(self.table.get_children()):
                values = self.table.item(item, 'values')
                for col, value in enumerate(values):
                    sheet.write(row + 1, col, value)

            workbook.close()


def close_all(app_instance):
    """Закрывает все дополнительные окна и основное окно."""
    for window in app_instance.additional_windows:
        if tk.Toplevel.winfo_exists(window):
            window.destroy()
    app_instance.window.quit()
    app_instance.window.destroy()  # Закрываем основное окно


if __name__ == "__main__":
    window = tk.Tk()
    app = App(window)
    window.protocol("WM_DELETE_WINDOW", lambda: close_all(app))
    window.mainloop()
