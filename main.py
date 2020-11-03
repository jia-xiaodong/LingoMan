#!/usr/bin/python3
# -*- coding: utf-8 -*-
import codecs

import pandas as pd
import os
import re
import sqlite3
import tkinter as tk
import tkinter.messagebox as messagebox
import tkinter.filedialog as filedialog
import subprocess
import json

game_root = r"D:\Projects\B2\UnityExperiment"
database_path = r'D:\tools\LingoMan\text_stats.sqlite3'
roslyn_finder = r'D:\demos\MySlnFindRef\FindTextRef\bin\Release\net472\FindTextRef.exe'


def str_split(separators, target):
    for c in separators:
        if c in target:
            return target.split(c)
    else:
        return [target]


def try_read_text_file(filename):
    ifs = open(filename, 'rb')
    raw = ifs.read()
    ifs.close()

    content = None
    utf_boms = {
        'utf-8-sig': [codecs.BOM_UTF8],
        'utf-32': [codecs.BOM_UTF32_LE, codecs.BOM_UTF32_BE],
        'utf-16': [codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE]
    }
    # 1. try BOM
    for enc, boms in utf_boms.items():
        if any(raw.startswith(bom) for bom in boms):
            content = raw.decode(encoding=enc)
            break
    else:
        # 2. without BOM
        for enc in ['utf-8', 'utf-16', 'utf-32', 'gb2312', 'big5', 'big5hkscs', 'gbk', 'gb18030', 'ansi']:
            try:
                content = raw.decode(encoding=enc)
                break
            except UnicodeDecodeError as e:
                pass
    return content


class StringSplit:
    """
    中间的间隔符可以有多个。str的内置split只支持一个间隔符。
    Coding Example:
        target = "abc;| def;     ghi |;"
        splitor = StringSplit('|;, ')
        strings = splitor.split(target)
    Output:
        abc
        def
        ghi
    """
    def __init__(self, separators):
        pattern = '([^%s]+)[%s]*?(.*)' % (separators, separators)
        self._regex = re.compile(pattern)

    def split(self, target):
        strings = []
        while len(target) > 0:
            result = self._regex.search(target)
            if result is None:
                return strings
            strings.append(result.group(1))
            target = result.group(2)
        return strings


class TextStats:
    def __init__(self):
        self._locations = {}

    def add_entry(self, text_id, location):
        if text_id in self._locations:
            self._locations[text_id].append(location)
        else:
            self._locations[text_id] = [location]

    @property
    def text_ids(self):
        return self._locations.keys()

    def locations(self, text_id):
        if text_id not in self._locations:
            return None
        return self._locations[text_id]

    def __contains__(self, text_id):
        return text_id in self._locations


class TextDataBase:
    def __init__(self, filename, connection=None):
        self._filename = filename
        self._con = sqlite3.connect(filename) if connection is None else connection

    def read_all(self):
        try:
            cur = self._con.cursor()
            cur.execute('SELECT tid, loc FROM used')
            return [(t, l) for t, l in cur.fetchall()]
        except Exception as e:
            print('Error on reading: %s' % e)
            return None

    def insert(self, text_id: str, location: str):
        """
        :param text_id:
        :param location:
        :return: -1 when error occurs.
        """
        try:
            sql = 'INSERT INTO used (tid, loc) VALUES(?,?)'
            arg = (text_id, location)
            cur = self._con.cursor()
            cur.execute(sql, arg)
            self._con.commit()
            return cur.lastrowid
        except Exception as e:
            print('Error on insertion (batch): %s' % e)
            return -1

    def insert_batch(self, texts):
        """
        :param texts: sequence of tuple(text_id, location)
        """
        try:
            sql = 'INSERT INTO used (tid, loc) VALUES(?,?)'
            self._con.executemany(sql, texts)
            self._con.commit()
        except Exception as e:
            print('Error on insertion (batch): %s' % e)

    def read_all_unused(self):
        try:
            cur = self._con.cursor()
            cur.execute('SELECT tid FROM unused')
            return [t[0] for t in cur.fetchall()]
        except Exception as e:
            print('Error on reading: %s' % e)
            return None

    def update_unused(self, text_ids):
        """
        :param text_ids: sequence of text_id
        """
        try:
            cur = self._con.cursor()
            cur.execute('DELETE FROM unused')  # delete all old rows

            sql = 'INSERT INTO unused (tid) VALUES(?)'
            args = [(i,) for i in text_ids]
            cur.executemany(sql, args)
            self._con.commit()
        except Exception as e:
            print('Error on insertion of unused texts: %s' % e)

    def append_unused(self, text_ids):
        """
        :param text_ids: sequence of text_id
        """
        try:
            cur = self._con.cursor()
            sql = 'INSERT INTO unused (tid) VALUES(?)'
            args = [(i,) for i in text_ids]
            cur.executemany(sql, args)
            self._con.commit()
        except Exception as e:
            print('Error on insertion of unused texts: %s' % e)

    def is_used(self, text_id):
        records = []
        try:
            cur = self._con.cursor()
            cur.execute('SELECT * FROM used WHERE tid=?', (text_id,))
            records = cur.fetchall()
        except Exception as e:
            print('Error on insertion of unused texts: %s' % e)
        finally:
            return len(records) > 0

    def close(self):
        self._con.close()
        self._filename = None

    @staticmethod
    def clear_database(filename):
        try:
            con = sqlite3.connect(filename)
            cur = con.cursor()
            cur.execute('DELETE FROM unused')
            cur.execute('DELETE FROM used')
            con.commit()
        except Exception as e:
            print('Error on clear of DB: %s' % e)

    @staticmethod
    def create_new(filename: str):
        """
        Make sure there's no such a file with the specified name.\n
        :param filename: database file name.
        :return: TextDataBase obj.
        """
        try:
            sql = '''CREATE TABLE used (
                id    INTEGER PRIMARY KEY UNIQUE NOT NULL,
                tid   TEXT NOT NULL,
                loc   text);
                CREATE TABLE unused (
                id    INTEGER PRIMARY KEY UNIQUE NOT NULL,
                tid   TEXT);'''
            con = sqlite3.connect(filename)
            con.executescript(sql)
            con.commit()
            return TextDataBase(filename, con)
        except Exception as e:
            print('Error on creation of DB: %s' % e)
            return None

    @staticmethod
    def open_old(filename):
        try:
            if not TextDataBase.validate(filename):  # wrong database (different structure)
                return None
            return TextDataBase(filename)
        except Exception as e:
            print('Error on opening DB: %s' % e)
            return None

    @staticmethod
    def is_col_name_same(cursor, table, columns):
        """
        check if column names are correct as expected.\n
        :param cursor: sqlite cursor obj
        :param table: table name
        :param columns: column names
        :return: True if all column names are correct
        """
        cursor.execute('PRAGMA table_info (%s)' % table)
        structs = cursor.fetchall()
        result = (structs[i][1] == columns[i] for i in range(len(columns)))
        return all(result)

    @staticmethod
    def validate(filename):
        """
        Check if specified file has expected table structure.\n
        :param filename: database file.
        :return: True if its structure is correct.
        """
        with sqlite3.connect(filename) as con:
            cur = con.cursor()
            tables = {
                'used': ['id', 'tid', 'loc'],
                'unused': ['id', 'tid']
            }
            correct = all(TextDataBase.is_col_name_same(cur, tbl, cols) for tbl, cols in tables.items())
        return correct


class MainApp(tk.Tk):
    TITLE = 'LingoMan'

    def __init__(self, *a, **kw):
        tk.Tk.__init__(self, *a, **kw)
        frame = tk.LabelFrame(self, text='（Step 1/3）记录多语言文本的使用情况（used）', padx=5, pady=5)
        frame.pack(side=tk.TOP, padx=5, pady=5, fill=tk.BOTH, expand=tk.YES)

        frame1 = tk.LabelFrame(frame, text='创建新的数据库（慢）', padx=5, pady=5)
        frame1.pack(side=tk.TOP, padx=5, pady=5, fill=tk.BOTH, expand=tk.YES)

        sub_frame = tk.Frame(frame1)
        sub_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=tk.YES)

        label = tk.Label(sub_frame, text='活动模板的列表：')
        label.pack(side=tk.LEFT, fill=tk.X, expand=tk.NO)

        self._activity_list = tk.StringVar()
        entry = tk.Entry(sub_frame, textvariable=self._activity_list)
        entry.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=tk.YES)

        btn = tk.Button(sub_frame, text='浏览', command=self.on_btn_find_activity_list)
        btn.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=tk.NO)

        btn = tk.Button(frame1, text='开始扫描：活动模板、代码、Prefab、数据表', command=self.on_button_create_new)
        btn.pack(side=tk.TOP, padx=5, pady=5)

        btn = tk.Button(frame, text='打开旧的数据库（快）', command=self.on_button_open_old)
        btn.pack(side=tk.TOP, padx=5, pady=5)

        frame = tk.LabelFrame(self, text='（Step 2/3）统计无用的多语言文本（unused）', padx=5, pady=5)
        frame.pack(side=tk.TOP, padx=5, pady=5, fill=tk.BOTH, expand=tk.YES)

        btn = tk.Button(frame, text='第一步：自动分析', command=self.on_btn_find_error)
        btn.pack(side=tk.TOP, padx=5, pady=5)

        sub_frame = tk.Frame(frame)
        sub_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=tk.YES)

        label = tk.Label(sub_frame, text='第二步：人工补充（可选）')
        label.pack(side=tk.LEFT, fill=tk.X, expand=tk.NO)

        self._manual_analysis = tk.StringVar()
        entry = tk.Entry(sub_frame, textvariable=self._manual_analysis)
        entry.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=tk.YES)

        btn = tk.Button(sub_frame, text='浏览', command=self.on_btn_browse_manual_file)
        btn.pack(side=tk.LEFT, padx=5, pady=5)

        btn = tk.Button(sub_frame, text='添加到数据库', command=self.on_btn_update_unused_manually)
        btn.pack(side=tk.LEFT, padx=5, pady=5)

        frame = tk.LabelFrame(self, text='（Step 3/3）导出Excel文件', padx=5, pady=5)
        frame.pack(side=tk.TOP, padx=5, pady=5, fill=tk.BOTH, expand=tk.YES)

        btn = tk.Button(frame, text='开始', command=self.dump_result)
        btn.pack(side=tk.TOP, padx=5, pady=5)
        #
        self.title(MainApp.TITLE)

        self._used_strings = None
        self._all_strings = set()
        self._database = None
        self._xlsx_sheets = []

    def on_button_create_new(self):
        list_file = self._activity_list.get()
        if len(list_file) == 0 or not os.path.exists(list_file):
            messagebox.showerror(MainApp.TITLE, 'Please specify a valid activity-template-list file (JSON)!')
            return

        if not os.path.exists(database_path):
            text_db = TextDataBase.create_new(database_path)
        elif messagebox.askyesno(MainApp.TITLE, 'Do you want to clear old data?'):
            TextDataBase.clear_database(database_path)
            text_db = TextDataBase.open_old(database_path)
        else:
            return
        #
        self.read_all_strings_from_xlsx()
        #
        used_strings = set()
        used_strings |= self.scan_activity_list()
        used_strings |= self.scan_solution()
        used_strings |= self.scan_prefab()
        used_strings |= self.scan_game_data()
        text_db.insert_batch(used_strings)
        self._database = text_db
        self._used_strings = MainApp.create_stats(used_strings)
        #
        messagebox.showinfo(MainApp.TITLE, '[Create Database] Job done!')

    def on_button_open_old(self):
        if not os.path.exists(database_path):
            messagebox.showerror(MainApp.TITLE, 'No database is found!')
            return
        #
        text_db = TextDataBase.open_old(database_path)
        if text_db is None:
            messagebox.showerror(MainApp.TITLE, 'Wrong database format!')
            return None

        used_strings = text_db.read_all()
        if len(used_strings) == 0:  # 正常情况下，游戏肯定会用到大量文本
            messagebox.showerror(MainApp.TITLE, 'No data is found!')
            return None
        #
        self.read_all_strings_from_xlsx()
        self._database = text_db
        self._used_strings = MainApp.create_stats(used_strings)
        #
        messagebox.showinfo(MainApp.TITLE, '[Load Database] Job done!')

    def read_all_strings_from_xlsx(self):
        self._xlsx_sheets.clear()
        self._all_strings.clear()
        #
        workbook = os.path.join(game_root, r'Assets\Text\LOC.xlsx')
        frame_dict = pd.read_excel(workbook, sheet_name=None)
        for sheet, frame in frame_dict.items():
            self._xlsx_sheets.append(sheet)
            try:
                for cell in frame['ID']:  # 暂时只考虑ID这一列。在精简了文本以后，可以全读出来做深入分析
                    text_id = '{0}_{1}'.format(sheet, cell)
                    self._all_strings.add(text_id.strip())
            except Exception as e:
                print(e)

    def on_btn_find_error(self):
        if self._used_strings is None:
            messagebox.showerror(MainApp.TITLE, 'Must load data from database or collect data from scratch at first!')
            return

        # 1. 第一次，粗筛
        id_used = set(self._used_strings.text_ids)
        undefined = id_used - self._all_strings  # 虽然是“使用”状态，但并未在LOC.xlsx中定义
        unused = self._all_strings - id_used  # 找不到引用之处
        # 2. 第二次，组合式的文本
        possible_defined_total = set()
        possible_used_total = set()
        regex_csharp = re.compile(r'(\w+)({.+})(\w*)')  # 搜索C#的Interpolated String，形如：$"LC_COMMON_{agentName}"
        print('--- possible used:')
        for each in undefined:
            possible_used = set()
            if regex_csharp.match(each):
                pattern = regex_csharp.sub(r'\1([a-zA-Z0-9_]+)\3', each)  # 实时创建出正则匹配器
                regex = re.compile(pattern)
            else:
                regex = None
            for full in self._all_strings:
                if full.find(each) > -1:
                    possible_used.add(full)
                elif regex is not None and regex.match(full):
                    possible_used.add(full)
            if len(possible_used) > 0:
                possible_defined_total.add(each)  # 汇总：
                print(each)                       # 1. 可以认为该项是“已经定义”（在多语言文本里）
                for i in possible_used:           # 2. 所有相关的匹配项都认为是“被使用”状态
                    print('\t' + i)
                possible_used_total |= possible_used  # 汇总
        undefined = undefined - possible_defined_total
        unused = unused - possible_used_total
        # 3. 第三次，大小写拼写错误
        print('--- possible spelling mistake:')
        all_strings_pairs = [(i.lower(), i) for i in self._all_strings]
        possible_defined_total.clear()
        possible_used_total.clear()
        for each in undefined:
            each_lowercase = each.lower()
            possible_used = set()
            if regex_csharp.match(each):
                pattern = regex_csharp.sub(r'\1([a-zA-Z0-9_]+)\3', each_lowercase)  # 实时创建出正则匹配器
                regex = re.compile(pattern)
            else:
                regex = None
            for full_lowercase, full in all_strings_pairs:
                if full_lowercase.find(each_lowercase) > -1:
                    possible_used.add(full)
                elif regex is not None and regex.match(full_lowercase):
                    possible_used.add(full)
            if len(possible_used) > 0:
                possible_defined_total.add(each)  # 汇总
                print(each)
                locations = self._used_strings.locations(each)
                for location in locations:
                    print('\t' + location)
                for i in possible_used:
                    print('\t' + i)
                possible_used_total |= possible_used  # 汇总
        undefined = undefined - possible_defined_total
        unused = unused - possible_used_total
        # 1/2. 输出最后结果：未找到定义
        if len(undefined) > 0:
            print('--- undefined text IDs:')
            for each in undefined:
                print(each)
                locations = self._used_strings.locations(each)
                for location in locations:
                    print('\t' + location)
        # 2/2. 输出最后结果：未找到引用
        if len(unused) > 0:
            print('--- unused text IDs:')
            for each in unused:
                print(each)
            self._database.update_unused(unused)
        # 输出结果到Excel表里
        if messagebox.askyesno(MainApp.TITLE, 'Do you want to output Used and Unused strings (to excel)?'):
            self.dump_result(unused)
        #
        messagebox.showinfo(MainApp.TITLE, '[Check Error] Job is done.')

    def on_btn_find_activity_list(self):
        options = {"title": "Open File: Activity Template List", "filetypes": [("JSON text", ("*.json")), ("Text file", ("*.txt"))]}
        filename = filedialog.askopenfilename(**options)
        if len(filename) == 0:
            return
        self._activity_list.set(filename)

    def on_btn_browse_manual_file(self):
        options = {"title": "打开文件：无用的多语言文本ID列表", "filetypes": [("Text file", ("*.txt"))]}
        filename = filedialog.askopenfilename(**options)
        if len(filename) == 0:
            return
        self._manual_analysis.set(filename)

    def on_btn_update_unused_manually(self):
        filename = self._manual_analysis.get()
        if len(filename) == 0:
            messagebox.showerror(MainApp.TITLE, '文件路径无效')
            return
        if self._database is None:
            messagebox.showerror(MainApp.TITLE, '数据库无效')
            return
        id_list = try_read_text_file(filename)
        id_list = id_list.split('\n')
        id_list = [i.strip() for i in id_list]
        id_list = [i for i in id_list if not self._database.is_used(i)]
        self._database.append_unused(id_list)

    def on_btn_double_check(self):
        """
        Too slow!
        :return:
        """
        if self._database is None:
            messagebox.showwarning(MainApp.TITLE, 'Database connection is required!')
            return
        unused = self._database.read_all_unused()
        if len(unused) == 0:
            messagebox.showinfo(MainApp.TITLE, '[Double Check] Table <unused> is empty!')
            return
        #
        unused = set(unused)
        print('\n--- possible referenced places:')
        for root, dirs, files in os.walk(game_root):
            for file in files:
                _, ext = os.path.splitext(file)
                if ext != '.cs':  # 只检查C#源码文件
                    continue
                full_file_name = os.path.join(root, file)
                content = try_read_text_file(full_file_name)
                if content is None:
                    print('Unknown encoding: ' + full_file_name)
                    continue
                for each in unused:
                    if re.search(each, content):
                        print('%s: %s' % (each, file))
        messagebox.showinfo(MainApp.TITLE, '[Double Check] Job done!')

    def has_section_only(self, name):
        for section in self._xlsx_sheets:
            if name.startswith(section) and len(name) - len(section) < 2:
                return True
        return False

    def scan_prefab(self):
        """
        Scan all text IDs in Unity prefab files.
        """
        strings = set()
        regex = re.compile(r'stringLocKey:\s+(\w+)')
        for root, dirs, files in os.walk(game_root):
            for file in files:
                _, ext = os.path.splitext(file)
                if ext != '.prefab':
                    continue
                full_file_name = os.path.join(root, file)
                ifs = open(full_file_name, 'r')
                for line in ifs:
                    result = regex.search(line)
                    if result is not None:
                        text_id = result.group(1).strip()
                        if self.has_section_only(text_id):
                            print('Error text ID: %s in %s' % (text_id, full_file_name))
                        else:
                            strings.add((text_id, file))
                ifs.close()
        return strings

    def scan_game_data(self):
        strings = set()
        pattern = '|'.join([i + '_' for i in self._xlsx_sheets])
        pattern = '((%s).+)' % pattern
        regex = re.compile(pattern)
        splitor = StringSplit('|;, ')  # 现暂时只发现了两种间隔符：;（分号）,（逗号）
        data_root = os.path.join(game_root, r'config')
        folders = ['GameDatasNew/Client', 'GameDatasNew/Server', 'GameDatasNew/Share', 'Campaign']
        for each_dir in folders:
            folder = os.path.join(data_root, each_dir)
            workbooks = [i for i in os.listdir(folder) if i.endswith('.xls')]
            for each_book in workbooks:
                # read all sheets at once [to a dictionary {sheet : data_frame}]
                frame_dict = pd.read_excel(os.path.join(folder, each_book), sheet_name=None)
                for sheet, frame in frame_dict.items():
                    for each_row in frame.index.values:
                        for each_col in range(len(frame.columns.values)):
                            cell = frame.values[each_row, each_col]
                            if not isinstance(cell, str):
                                continue
                            for each_str in splitor.split(cell):  # cell可以容纳多条文本，以分号间隔开。
                                result = regex.search(each_str)
                                if result is None:
                                    continue
                                location = '%s - %s' % (each_dir, each_book)
                                text_id = result.group(1)
                                if self.has_section_only(text_id):
                                    print('Error text ID: %s in <%s>' % (text_id, location))
                                else:
                                    strings.add((text_id, location))
        return strings

    def scan_solution(self):
        sections = ','.join(self._xlsx_sheets)
        solution = os.path.join(game_root, 'UnityExperiment.sln')
        proc = subprocess.run([roslyn_finder, solution, 'Assembly-CSharp', sections], capture_output=True, encoding='gbk')
        if proc.returncode != 0:
            print('Error on collecting symbols: %s' % proc.stderr)
            return set()
        strings = set()
        regex = re.compile(r'TEXT:\s+(.+),(.+)')
        outputs = proc.stdout.split('\n')
        for line in outputs:
            result = regex.match(line)
            if result is None:
                continue
            tid, location = result.group(1), result.group(2)
            strings.add((tid, location))
        return strings

    def scan_activity_list(self):
        filename = self._activity_list.get()
        if not os.path.exists(filename):
            return set()
        try:
            template_file = try_read_text_file(filename)
            templates = json.loads(template_file)
            if templates["type"] != "rule_aty":
                return set()
            strings = set()
            for activity in templates["data"]:
                identity = activity["Id"]
                # 1/5
                title = activity["Title"]
                if len(title) > 0:
                    strings.add((title, 'activity: %s, title' % identity))
                # 2/5
                icon_title = activity["IconTitle"]
                if len(icon_title) > 0:
                    strings.add((icon_title, 'activity: %s, icon-title' % identity))
                # 3/5
                description_full = activity["Desc"]
                if len(description_full) > 0:
                    strings.add((description_full, 'activity: %s, description_full' % identity))
                # 4/5
                description_short = activity["ShortDesc"]
                if len(description_short) > 0:
                    strings.add((description_short, 'activity: %s, description_short' % identity))
                # 5/5
                rule = activity["Rule"]
                if len(rule) > 0:
                    strings.add((rule, 'activity: %s, rule' % identity))
        except Exception as e:
            print(e)
        finally:
            return strings

    def dump_result(self, unused=None):
        if self._database is None:
            messagebox.showerror(MainApp.TITLE, 'Must load data from database or collect data from scratch at first!')
            return
        if unused is None:
            unused = self._database.read_all_unused()
        #
        # 为加快速度，把文本ID拆分成多个集合
        unused_dict = {}
        prefixes = [i + '_' for i in self._xlsx_sheets]
        for i in self._xlsx_sheets:
            unused_dict[i] = set()
        for each_text in unused:
            for prefix in prefixes:
                if each_text.startswith(prefix):
                    tid = each_text[len(prefix):]
                    section = prefix[:-1]
                    unused_dict[section].add(tid)
                    break
        # 最后结果输出到这两个文件里
        writer_used = pd.ExcelWriter("used.xlsx")
        writer_unused = pd.ExcelWriter("unused.xlsx")
        # 遍历源Excel
        workbook = os.path.join(game_root, r'Assets\Text\LOC.xlsx')
        frame_dict = pd.read_excel(workbook, sheet_name=None)
        for sheet, frame in frame_dict.items():
            try:
                frame_used = pd.DataFrame(columns=frame.columns)
                frame_unused = pd.DataFrame(columns=frame.columns)
                for index, row in frame.iterrows():
                    if row['ID'] in unused_dict[sheet]:
                        frame_unused.loc[index] = row
                    else:
                        frame_used.loc[index] = row
                frame_used.to_excel(writer_used, sheet_name=sheet, index=False)
                writer_used.save()
                frame_unused.to_excel(writer_unused, sheet_name=sheet, index=False)
                writer_unused.save()
            except Exception as e:
                print(e)
        # 保存到文件
        writer_used.close()
        writer_unused.close()

    @staticmethod
    def create_stats(used_strings):
        stats = TextStats()
        for tid, location in used_strings:
            stats.add_entry(tid, location)
        return stats


def test():
    target = r"Tutorial_New_Fte_chat_{0}_5"
    sections = ['Tutorial']
    pattern = '|'.join([i + '_' for i in sections])
    pattern = '((%s).+)' % pattern
    regex = re.compile(pattern)
    splitor = StringSplit('|;, ')
    targets = splitor.split(target)
    for each in targets:
        result = regex.match(each)
        if result is not None:
            print(result.group(1))


if __name__ == "__main__":
    MainApp().mainloop()
    #test()
