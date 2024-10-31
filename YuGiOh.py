import os
import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import winreg

def load_options(file_path):
    options = {}
    try:
        exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        external_file = os.path.join(exe_dir, 'YuGiOh.txt')

        if os.path.exists(external_file):
            file_path = external_file

        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                key, value = line.strip().split(' ', 1)
                options[key] = value
    except FileNotFoundError:
        messagebox.showerror("Error", f"Options file not found: {file_path}")
        return None
    except Exception as e:
        messagebox.showerror("Error", f"Error reading options file: {str(e)}")
        return None
    return options

def rename_folder(path, new_name):
    try:
        subdirs = [d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))]
        if len(subdirs) != 1:
            messagebox.showerror("Error", f"Expected 1 subfolder in {path}, found {len(subdirs)}")
            return False
        old_name = subdirs[0]
        old_path = os.path.join(path, old_name)
        new_path = os.path.join(path, new_name)
        os.rename(old_path, new_path)
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Error renaming folder: {str(e)}")
        return False

def on_select(*args):
    selected = option_var.get()
    new_name = options.get(selected)
    if new_name:
        base_path = folder_path.get()
        if base_path:
            if rename_folder(base_path, new_name):
                messagebox.showinfo("Success", f"Folder renamed to {new_name}")
                save_last_path(base_path)
            root.quit()
        else:
            messagebox.showerror("Error", "Please select a folder path")
    else:
        messagebox.showerror("Error", "Invalid selection")

def get_resource_path(relative_path):
    current_dir_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)
    if os.path.exists(current_dir_path):
        return current_dir_path

    # 如果当前目录没有，则使用打包后的路径
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def save_last_path(path):
    try:
        # 创建或打开注册表键
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\YuGiOhFolderRenamer")
        # 保存路径
        winreg.SetValueEx(key, "LastPath", 0, winreg.REG_SZ, path)
        winreg.CloseKey(key)
    except Exception as e:
        print(f"Error saving path: {e}")

def load_last_path():
    try:
        # 打开注册表键
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\YuGiOhFolderRenamer", 0, winreg.KEY_READ)
        # 读取路径
        path = winreg.QueryValueEx(key, "LastPath")[0]
        winreg.CloseKey(key)
        # 验证路径是否存在
        if os.path.exists(path):
            return path
    except Exception as e:
        print(f"Error loading path: {e}")
    return r"D:\Steam\steamapps\common\Yu-Gi-Oh!  Master Duel\LocalData"

# Load options
options_file = get_resource_path('YuGiOh.txt')
options = load_options(options_file)
if not options:
    exit()

root = tk.Tk()
root.title("Folder Renamer")

tk.Label(root, text="Select an option:").pack(pady=10)

option_var = tk.StringVar(root)
option_menu = tk.OptionMenu(root, option_var, *options.keys())
option_menu.pack(pady=10)

option_var.trace('w', on_select)

tk.Label(root, text="Yu-Gi-Oh! Master Duel folder path:").pack(pady=5)
folder_path = tk.StringVar(root, value=load_last_path())
folder_entry = tk.Entry(root, textvariable=folder_path, width=50)
folder_entry.pack(pady=5)

def browse_folder():
    folder = filedialog.askdirectory()
    if folder:
        folder_path.set(folder)
        save_last_path(folder)

browse_button = tk.Button(root, text="Browse", command=browse_folder)
browse_button.pack(pady=10)

# 添加窗口关闭事件处理
def on_closing():
    save_last_path(folder_path.get())
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()
