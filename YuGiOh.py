import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
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

def browse_folder():
    path = filedialog.askdirectory()
    if path:
        folder_path.set(path)

# Load options
options_file = get_resource_path('YuGiOh.txt')
options = load_options(options_file)
if not options:
    exit()

root = tk.Tk()
root.title("Yu-Gi-Oh! Master Duel Folder Renamer")
root.geometry("600x300")  # 设置窗口大小

# 配置样式
style = ttk.Style()
style.configure('TLabel', font=('Arial', 10))
style.configure('TButton', font=('Arial', 10))
style.configure('Header.TLabel', font=('Arial', 12, 'bold'))

# 创建主框架
main_frame = ttk.Frame(root, padding="20")
main_frame.pack(fill=tk.BOTH, expand=True)

# 标题
header = ttk.Label(main_frame, text="Yu-Gi-Oh! Master Duel Folder Renamer", style='Header.TLabel')
header.pack(pady=(0, 20))

# 选项框架
option_frame = ttk.LabelFrame(main_frame, text="Rename Options", padding="10")
option_frame.pack(fill=tk.X, padx=5, pady=5)

option_var = tk.StringVar(root)
option_menu = ttk.Combobox(option_frame, textvariable=option_var, values=list(options.keys()), state='readonly')
option_menu.pack(fill=tk.X, padx=5, pady=5)
option_menu.bind('<<ComboboxSelected>>', on_select)

# 路径框架
path_frame = ttk.LabelFrame(main_frame, text="Folder Path", padding="10")
path_frame.pack(fill=tk.X, padx=5, pady=10)

folder_path = tk.StringVar(root, value=load_last_path())
path_entry = ttk.Entry(path_frame, textvariable=folder_path, width=50)
path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))

browse_button = ttk.Button(path_frame, text="Browse", command=browse_folder)
browse_button.pack(side=tk.RIGHT, padx=5)

# 添加说明文本
info_text = """
Please select a rename option from the dropdown menu and verify the folder path.
The program will rename the selected folder according to your choice.
"""
info_label = ttk.Label(main_frame, text=info_text, wraplength=500, justify=tk.LEFT)
info_label.pack(pady=20)

# 添加底部按钮
button_frame = ttk.Frame(main_frame)
button_frame.pack(fill=tk.X, pady=10)

# Define the function first
def on_closing():
    save_last_path(folder_path.get())
    root.destroy()

# Then create the button
cancel_button = ttk.Button(button_frame, text="Cancel", command=on_closing)
cancel_button.pack(side=tk.RIGHT, padx=5)

# 添加窗口关闭事件处理
root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()
