#!/usr/bin/env python3

#
# test_seek_directories.py
#
# Date    : 2024-05-21
# Auther  : Hirotoshi FUJIBE
# History :
#
# Copyright (c) 2024 Hirotoshi FUJIBE
#

import os
import sys
import glob
from pathlib import Path


# os.walk
def os_walk() -> None:
    file_list = []
    for root, dirs, files in os.walk('.\\input\\src'):
        for filename in files:
            file_list.append(os.path.join(root, filename))
        for dirname in dirs:
            file_list.append(os.path.join(root, dirname))
    for file_name in file_list:
        print(file_name)
    return


# glob.glob()
def glob_glob() -> None:

    file_list = glob.glob('.\\input\\src\\**/*', recursive=True)
    for file_name in file_list:
        print(file_name)
    return


# Path.glob()
def path_glob() -> None:

    file_list = list(Path('.\\input\\src').glob("**/*"))
    for file_name in file_list:
        print('.\\%s' % file_name)
    return


# Seek Directories
def seek_directories(level: int, dir_root: str, dir_relative: str) -> None:
    dirs = []
    files = []
    for path in os.listdir(dir_root):
        if os.path.isfile(os.path.join(dir_root, path)):
            files.append(path)
        else:
            dirs.append(path)
    files.sort(key=str.lower)
    for file in files:
        print(os.path.join(dir_root, file))
    dirs.sort(key=str.lower)
    for dir_nest in dirs:
        print(os.path.join(dir_root, dir_nest))
        seek_directories(level + 1, os.path.join(dir_root, dir_nest), os.path.join(dir_relative, dir_nest))
    return


# Main
def main() -> None:

    print('-- os.walk() -- ルートフォルダーファイル、子フォルダー、子フォルダーファイル、...')
    os_walk()

    print('-- glob.glob() --　ルートフォルダーファイル、子フォルダー、子フォルダーファイル、...（.ignore、.c.c が取り出せない？）')
    glob_glob()

    print('-- Path.glob() -- ルートフォルダーファイル、子フォルダー、子フォルダーファイル、...')
    path_glob()

    print('-- os.listdir() の再帰呼出し -- ルートフォルダーファイル、子フォルダー、子フォルダーのファイル、...')
    seek_directories(0, '.\\input\\src', '.\\src')

    sys.exit(0)


# Goto Main
if __name__ == '__main__':
    main()
