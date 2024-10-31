import os
import tkinter as tk
from tkinter import filedialog, messagebox
import sys

def load_options(file_path):
    options = {}
    try:
        # 首先尝试读取exe同目录下的文件
        exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        external_file = os.path.join(exe_dir, 'YuGiOh.txt')
        
        # 如果外部文件存在，优先使用外部文件
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
        # 获取LocalData文件夹下的唯一子文件夹
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
            root.quit()
        else:
            messagebox.showerror("Error", "Please select a folder path")
    else:
        messagebox.showerror("Error", "Invalid selection")

def get_resource_path(relative_path):
    # 首先检查当前目录
    current_dir_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)
    if os.path.exists(current_dir_path):
        return current_dir_path
        
    # 如果当前目录没有，则使用打包后的路径
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# Load options
options_file = get_resource_path('YuGiOh.txt')
options = load_options(options_file)
if not options:
    exit()

# Create main window
root = tk.Tk()
root.title("Folder Renamer")

# Create and pack widgets
tk.Label(root, text="Select an option:").pack(pady=10)

option_var = tk.StringVar(root)
option_menu = tk.OptionMenu(root, option_var, *options.keys())
option_menu.pack(pady=10)

option_var.trace('w', on_select)

tk.Label(root, text="Yu-Gi-Oh! Master Duel folder path:").pack(pady=5)
default_path = r"D:\Steam\steamapps\common\Yu-Gi-Oh!  Master Duel\LocalData"
folder_path = tk.StringVar(root, value=default_path)
folder_entry = tk.Entry(root, textvariable=folder_path, width=50)
folder_entry.pack(pady=5)

def browse_folder():
    folder = filedialog.askdirectory()
    if folder:
        folder_path.set(folder)

browse_button = tk.Button(root, text="Browse", command=browse_folder)
browse_button.pack(pady=10)

root.mainloop()
