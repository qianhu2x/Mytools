import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import os
import json

# ====================== 在这里配置你的工具 ======================
# type: normal / log_file / log_folder
# normal: 普通工具不带选择
# log_file: 点了选单个log文件 自动当参数传给脚本
# log_folder: 点了选log整个文件夹 自动当参数传给脚本
TOOLS = {
    "记事本": {"path":"C:/Windows/notepad.exe", "arg":"", "type":"normal"},
    "分析单个Log": {"path":"python", "arg":"D:/tools/analyze.py", "type":"log_file"},
    "分析Log文件夹": {"path":"python", "arg":"D:/tools/analyze.py", "type":"log_folder"}
}
# ==================================================================

CONFIG_FILE = "tools_config.json"

class ToolLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("📌 工具启动器")
        self.root.geometry("500x500")
        self.root.resizable(False, False)
        
        self.tools = self.load_tools()
        self.create_widgets()

    def load_tools(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except:
            pass
        return TOOLS.copy()

    def save_tools(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.tools, f, ensure_ascii=False, indent=2)
        except:
            pass

    def create_widgets(self):
        title_label = ttk.Label(
            self.root, 
            text="我的工具集", 
            font=("微软雅黑", 16, "bold")
        )
        title_label.pack(pady=15)

        self.button_frame = ttk.Frame(self.root)
        self.button_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        self.refresh_buttons()

        control_frame = ttk.Frame(self.root)
        control_frame.pack(pady=15)

        ttk.Button(
            control_frame, 
            text="添加工具", 
            command=self.add_tool_window
        ).grid(row=0, column=0, padx=10)

        ttk.Button(
            control_frame, 
            text="删除工具", 
            command=self.delete_tool_window
        ).grid(row=0, column=1, padx=10)

    def refresh_buttons(self):
        for widget in self.button_frame.winfo_children():
            widget.destroy()

        if not self.tools:
            ttk.Label(
                self.button_frame, 
                text="暂无工具，请点击【添加工具】", 
                font=("微软雅黑", 11)
            ).pack(pady=20)
            return

        for name in self.tools:
            btn = ttk.Button(
                self.button_frame,
                text=name,
                command=lambda n=name: self.launch_tool(n),
                width=30
            )
            btn.pack(pady=6, ipady=5)

    def launch_tool(self, name):
        info = self.tools[name]
        path = info["path"]
        arg = info["arg"]
        ttype = info["type"]

        cmd_str = f"{path} {arg}"

        # 选单个log文件
        if ttype == "log_file":
            p = filedialog.askopenfilename(title="选择Log文件", filetypes=[("log/txt", "*.log;*.txt"), ("所有文件", "*.*")])
            if not p:
                return
            cmd_str = f"{path} {arg} {p}"

        # 选log文件夹
        elif ttype == "log_folder":
            p = filedialog.askdirectory(title="选择Log文件夹")
            if not p:
                return
            cmd_str = f"{path} {arg} {p}"

        # 完全沿用你第一版最稳的启动方式
        try:
            subprocess.Popen(cmd_str, shell=True)
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def add_tool_window(self):
        top = tk.Toplevel(self.root)
        top.title("添加新工具")
        top.geometry("450x280")
        top.resizable(False, False)
        top.grab_set()

        ttk.Label(top, text="工具名称：").pack(pady=5)
        name_entry = ttk.Entry(top, width=40)
        name_entry.pack(pady=2)

        ttk.Label(top, text="路径：python 或 程序exe完整路径").pack(pady=5)
        path_entry = ttk.Entry(top, width=40)
        path_entry.pack(pady=2)

        ttk.Label(top, text="固定参数：如 D:/xxx.py").pack(pady=5)
        arg_entry = ttk.Entry(top, width=40)
        arg_entry.pack(pady=2)

        ttk.Label(top, text="类型：normal / log_file / log_folder").pack(pady=5)
        type_entry = ttk.Entry(top, width=40)
        type_entry.insert(0, "normal")
        type_entry.pack(pady=2)

        def confirm_add():
            name = name_entry.get().strip()
            pth = path_entry.get().strip()
            ag = arg_entry.get().strip()
            tp = type_entry.get().strip()
            if not name or not pth:
                messagebox.showwarning("提示", "名称和路径不能为空")
                return
            self.tools[name] = {"path":pth, "arg":ag, "type":tp}
            self.save_tools()
            self.refresh_buttons()
            top.destroy()

        ttk.Button(top, text="确认添加", command=confirm_add).pack(pady=15)

    def delete_tool_window(self):
        if not self.tools:
            messagebox.showwarning("提示", "没有可删除的工具！")
            return

        top = tk.Toplevel(self.root)
        top.title("删除工具")
        top.geometry("400x150")
        top.resizable(False, False)
        top.grab_set()

        tool_names = list(self.tools.keys())
        selected = tk.StringVar(value=tool_names[0])
        combo = ttk.Combobox(top, textvariable=selected, values=tool_names, state="readonly")
        combo.pack(pady=10)

        def confirm_delete():
            name = selected.get()
            del self.tools[name]
            self.save_tools()
            self.refresh_buttons()
            top.destroy()

        ttk.Button(top, text="确认删除", command=confirm_delete).pack(pady=15)

if __name__ == "__main__":
    root = tk.Tk()
    app = ToolLauncher(root)
    root.mainloop()