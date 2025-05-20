from tkinter import *
import tkinter as tk
import tkinter.font as tkFont
import time
import re
from apscheduler.schedulers.background import BackgroundScheduler
import os
import shutil
import zipfile
from datetime import datetime
import psutil
import sys

DIR = [  # New_dir,                   local_dir,                         backup_dir
    [r'New file dir', r'Local file dir', r'Backup file dir'],
]

Log_dir = [
    r'Log file dir',

]


def clear_folder(folder_path):
    shutil.rmtree(folder_path)
    os.makedirs(folder_path)


def check(New_dir, local_dir, backup_dir):  # Check whether local dir file match with New dir file

    # Get latest file filename
    global Local_file_filename
    for file_root, dirs, files in os.walk(New_dir):
        for file in files:
            if file[-3:] == "zip":
                target_file = file
                Latest_file_filename = file[:-4]
                print("Latest File:", Latest_file_filename)

    # Local directory for file software
    if not os.path.exists(local_dir):
        os.makedirs(local_dir)

    # Get local file filename
    for file_root, dirs, files in os.walk(
            local_dir):  # Compare whether the  file same with the seed file in sharepoint
        for file in files:
            if file == []:
                print("file empty")

            if file[-3:] == "zip":
                local_file = file
                Local_file_filename = file[:-4]
                print("Local file:", Local_file_filename)

    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)

    if Local_file_filename == Latest_file_filename:
        print("File updated")

        Label(root, text='File 已更新' + '\n' + Latest_file_filename, background='#00FF00', font=1).place(x=20, y=30)
        root.attributes('-alpha', 0.3)

    else:
        print("File not updated")
        Label(root, text='File 未更新' + '\n' + Local_file_filename, background='#FF0000', font=1).place(x=20, y=30)
        root.attributes('-alpha', 0.3)

        local_file_file = local_dir + '\\' + local_file
        check_file_in_backup_dir = backup_dir + '\\' + local_file

        if not os.path.exists(check_file_in_backup_dir):
            shutil.move(local_file_file, backup_dir)  # move old file to backup directory
            print("Old file moved to backup folder")
        else:
            print("Local file file exist in backup folder already")

        clear_folder(local_dir)  # Clear existing File files

        target_file_file = New_dir + '\\' + target_file
        shutil.copy(target_file_file, local_dir)  # move new config to dir
        print("Latest file copied to local folder")

        shutil.unpack_archive(target_file_file, local_dir)  # Auto unpack the ip file
        print("File updated. Zip file unpacked successfully")
        Label(root, text='File updated' + '\n' + Latest_file_filename, background='#00FF00', font=1).place(x=20, y=30)
        root.attributes('-alpha', 0.3)

        write_logs(Log_dir, Local_file_filename, Latest_file_filename)


def path():
    if check_process_status('RuntimeBroker.exe'):  #  software name input
        return

    global New_dir, local_dir, backup_dir
    New_dir = None
    local_dir = None
    backup_dir = None
    for A, B, C in DIR:
        New_dir = A
        local_dir = B
        backup_dir = C
        check(New_dir, local_dir, backup_dir)


def write_logs(Log_dir, filename_Last, filename_new):
    for file_path in Log_dir:
        try:
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            with open(file_path, 'a') as file:
                current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                current = f'Update time: {current_time}   filename_Last: {filename_Last}   filename_new: {filename_new}'
                file.write(current + '\n')
        except Exception as e:
            print(f"写入日志文件 {file_path} 时出错: {e}")


def center_window(root, width, height):
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width), 50)
    root.geometry(size)


def check_process_status(target_name, case_sensitive=False):
    target_name = target_name.strip()  # 移除首尾空格
    if not target_name:
        raise ValueError("进程名不能为空")

    # 统一大小写处理
    if not case_sensitive:
        target_name = target_name.lower()

    for process in psutil.process_iter():
        try:
            # 获取进程名并处理大小写
            process_name = process.name()
            if not case_sensitive:
                process_name = process_name.lower()

            # 匹配进程名
            if process_name == target_name:
                print(f"[检测到] 进程 '{target_name}' 正在运行")
                Label(root, text='进程正在运行' + '\n' + '请先关闭程序', background='#00FF00', font=1).place(x=200,
                                                                                                             y=30)
                root.attributes('-alpha', 0.3)
                return True

        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # 进程已终止/权限不足/僵尸进程，跳过
            continue

    print(f"[未找到] 进程 '{target_name}' 未运行")
    Label(root, text='进程未运行', background='#00FF00', font=1).place(x=200, y=30)
    root.attributes('-alpha', 0.3)
    return False


if __name__ == "__main__":
    global New_dir, local_dir, backup_dir
    # 創件文件路徑對話框
    root = Tk()

    # Font type
    ft1 = tkFont.Font(size=25, weight=tkFont.BOLD)
    ft2 = tkFont.Font(size=15)

    # 對話框格式
    root.title('File_detection_v1.0')
    root.attributes('-alpha', 0.70)  # 背景半透明
    root.attributes('-topmost', True)  # 程序永久置頂
    root.config(bg='#C0C0C0')
    center_window(root, 490, 100)
    root.overrideredirect(True)  # 去除邊框和標題欄

    scheduler = BackgroundScheduler()
    # 每隔2分钟执行一次， */1：每隔1分钟执行一次
    scheduler.add_job(path, 'cron', second='*/10', id='job1')

    scheduler.start()

    Button(root, text='離開', command=root.destroy, font=ft2).place(x=350, y=20)  # 對話框"退出"按鈕

    mainloop()
