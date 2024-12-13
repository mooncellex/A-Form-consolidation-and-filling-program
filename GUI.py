# -*- coding: utf-8 -*-
"""
Created on Fri Nov 22 20:26:51 2024

@author: mooncell
"""

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

def merge_excel_files(input_folder, output_file, sheet_name='Sheet1'):
    merged_df = pd.DataFrame()
    files_found = False  # 添加一个标志来跟踪是否找到了Excel文件
    for filename in os.listdir(input_folder):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            files_found = True  # 设置标志为True，表示找到了Excel文件
            file_path = os.path.join(input_folder, filename)
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=1)
                df = df[df.iloc[:, 0] != '合计']
                if not df.empty:
                    df = df.iloc[:-1, :]
                    merged_df = pd.concat([merged_df, df], ignore_index=True)
            except Exception as e:
                print(f"Error reading file {filename}: {e}")  # 打印读取单个文件的错误，但不影响其他文件
    if not merged_df.empty:
        merged_df.to_excel(output_file, index=False, sheet_name=sheet_name)
        print(f"合并文件已存入 {output_file}")
    else:
        if files_found:  # 如果找到了文件但合并后的DataFrame为空，可能是因为所有文件都是空的或不符合条件
            print("没有读取到任何有效数据，请检查输入文件夹中的Excel文件内容.")
        else:  # 如果没有找到任何Excel文件
            raise ValueError("在指定的文件夹中没有找到任何Excel文件.")

def select_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, folder_selected)
 
def start_merge():
    input_folder = folder_entry.get()
    if not input_folder or not os.path.isdir(input_folder):
        messagebox.showerror("错误", "请选择包含Excel文件的文件夹")
        return
    
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_file:
        return  # 如果用户取消了保存操作
    
    try:
        merge_excel_files(input_folder, output_file)
        messagebox.showinfo("成功", "Excel 文件成功合并!")
    except ValueError as ve:
        messagebox.showerror("错误", ve)
    except Exception as e:
        messagebox.showerror("错误", f"发生了如下的错误: {e}")
        
# 创建主窗口
root = tk.Tk()
root.title("Excel合并器")

# 创建和放置文件夹选择标签和输入框
folder_label = tk.Label(root, text="请选择包含Excel文件的文件夹:")
folder_label.grid(row=0, column=0, padx=10, pady=10)

folder_entry = tk.Entry(root, width=50)
folder_entry.grid(row=0, column=1, padx=10, pady=10)

# 创建和放置浏览按钮
browse_button = tk.Button(root, text="浏览", command=select_folder)
browse_button.grid(row=0, column=2, padx=10, pady=10)

# 创建和放置输出文件标签和保存按钮
#output_label = tk.Label(root, text="保存合并文件为:")
#output_label.grid(row=1, column=0, padx=10, pady=10)

# 合并按钮
merge_button = tk.Button(root, text="开始合并", command=start_merge)
merge_button.grid(row=2, column=0, columnspan=3, pady=20)

# 运行主循环
root.mainloop()
