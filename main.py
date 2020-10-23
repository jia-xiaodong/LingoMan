#!/usr/bin/python3
# -*- coding: utf-8 -*-
import pandas as pd
import os
import re

game_root = r"D:\Projects\B2\UnityExperiment"


def read_string_from_solution():
    return set()


def read_string_from_prefab():
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
                    strings.add(result.group(1))
            ifs.close()
    return strings


def read_string_from_game_data():
    """
    All game data are stored in xls sheet files.
    :return:
    """
    strings = set()
    data_root = os.path.join(game_root, r'config\GameDatasNew')
    regex = re.compile(r'(LC_\w+)')
    folders = ['Client', 'Server', 'Share']
    for each_dir in folders:
        folder = os.path.join(data_root, each_dir)
        wordbooks = [i for i in os.listdir(folder) if i.endswith('.xls')]
        for each_book in wordbooks:
            # read all sheets at once [to a dictionary {sheet : data_frame}]
            frame_dict = pd.read_excel(os.path.join(folder, each_book), sheet_name=None)
            for sheet, frame in frame_dict.items():
                for each_row in frame.index.values:
                    for each_col in range(len(frame.columns.values)):
                        cell = frame.values[each_row, each_col]
                        if not isinstance(cell, str):
                            continue
                        result = regex.search(cell)
                        if result is not None:
                            strings.add(result.group(1))
    return strings


def read_string_from_xlsx():
    strings = set()
    workbook = os.path.join(game_root, r'Assets\Text\LOC.xlsx')
    frame_dict = pd.read_excel(workbook, sheet_name=None)
    for sheet, frame in frame_dict.items():
        for cell in frame['ID']: # 暂时只考虑ID这一列
            strings.add('LC_{0}_{1}'.format(sheet, cell))
    return strings


def main():
    used_strings = set()
    used_strings |= read_string_from_solution()
    used_strings |= read_string_from_prefab()
    used_strings |= read_string_from_game_data()
    print('strings used in Game: %d' % len(used_strings))

    all_strings = read_string_from_xlsx()
    print('strings stored in xlsx: %d' % len(all_strings))

    unrecognized = used_strings - all_strings
    if len(unrecognized) > 0:
        print('----- Unrecognized strings ------')
        for s in unrecognized:
            print(s)
    else:
        redundant = all_strings - used_strings


def test():
    regex = re.compile(r'(LC_\w+)')
    result = regex.search(r'LC_abcdef')
    if result is not None:
        s = result.group(1)
        if len(s) > 0:
            print(s)


if __name__ == "__main__":
    main()
