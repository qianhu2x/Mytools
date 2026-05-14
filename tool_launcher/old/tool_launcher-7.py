import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import os
import json

# 初始工具
TOOLS = {
    "示例": {
        "exe": "python",
        "param": "hsdes_ticket_creator.py --excel WO_Template-CRB-DDR.xlsx",
        "log_path": "",
        "cwd": "D:/你的脚本目录"
    }
}

CFG_FILE = "tools_cfg.json"

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("工具启动器")
        self.root.geometry("680x480")
        self.root.resizable(False, False)
        self.tools = self.load_cfg()
        self.init_ui()

    def load_cfg(self):
        try:
            if os.path.exists(CFG_FILE):
                with open(CFG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except:
            pass
        return TOOLS.copy()

    def save_cfg(self):
        try:
            with open(CFG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.tools, f, ensure_ascii=False, indent=2)
        except:
            pass

    def init_ui(self):
        ttk.Label(self.root, text="我的工具列表", font=("微软雅黑",15,"bold")).pack(pady=10)
        self.frame = ttk.Frame(self.root)
        self.frame.pack(fill=tk.BOTH, expand=True, padx=30)
        self.refresh_btn()

        f = ttk.Frame(self.root)
        f.pack(pady=10)
        ttk.Button(f, text="添加", command=self.add_win).grid(row=0,column=0,padx=5)
        ttk.Button(f, text="编辑", command=self.edit_win).grid(row=0,column=1,padx=5)
        ttk.Button(f, text="删除", command=self.del_win).grid(row=0,column=2,padx=5)

    def refresh_btn(self):
        for w in self.frame.winfo_children():
            w.destroy()
        if not self.tools:
            ttk.Label(self.frame, text="暂无工具").pack(pady=30)
            return
        for name in self.tools:
            ttk.Button(
                self.frame, text=name, width=50,
                command=lambda n=name:self.run_tool(n)
            ).pack(pady=4, ipady=4)

    def run_tool(self, name):
        try:
            cfg = self.tools[name]
            cmd = cfg["exe"].strip()
            if cfg["param"].strip():
                cmd += " " + cfg["param"].strip()
            if cfg["log_path"].strip():
                cmd += " " + cfg["log_path"].strip()

            cwd = cfg.get("cwd", "").strip()
            if not cwd:
                cwd = None

            # 最稳定启动方式
            subprocess.Popen(cmd, shell=True, cwd=cwd)
        except:
            messagebox.showerror("错误", "启动失败")

    def add_win(self):
        win = tk.Toplevel(self.root)
        win.title("添加工具")
        win.geometry("650x350")
        win.resizable(False, False)
        win.grab_set()

        ttk.Label(win, text="工具名称").pack()
        e_name = ttk.Entry(win, width=80)
        e_name.pack()

        ttk.Label(win, text="1. 工具路径(python/exe)").pack()
        e_exe = ttk.Entry(win, width=80)
        e_exe.pack()

        ttk.Label(win, text="2. 参数").pack()
        e_param = ttk.Entry(win, width=80)
        e_param.pack()

        ttk.Label(win, text="3. Log路径").pack()
        e_log = ttk.Entry(win, width=80)
        e_log.pack()

        ttk.Label(win, text="4. 运行目录（填脚本所在文件夹）").pack()
        e_cwd = ttk.Entry(win, width=80)
        e_cwd.pack()

        def ok():
            n = e_name.get().strip()
            exe = e_exe.get().strip()
            if not n or not exe:
                messagebox.showwarning("提示", "名称和路径不能为空")
                return
            self.tools[n] = {
                "exe": exe,
                "param": e_param.get().strip(),
                "log_path": e_log.get().strip(),
                "cwd": e_cwd.get().strip()
            }
            self.save_cfg()
            self.refresh_btn()
            win.destroy()

        ttk.Button(win, text="确定", command=ok).pack(pady=10)

    def edit_win(self):
        if not self.tools:
            messagebox.showwarning("提示","无工具")
            return
        win = tk.Toplevel(self.root)
        win.title("编辑工具")
        win.geometry("650x350")
        win.resizable(False, False)
        win.grab_set()

        names = list(self.tools.keys())
        var = tk.StringVar(value=names[0])
        ttk.Label(win, text="选择工具").pack()
        cbx = ttk.Combobox(win, textvariable=var, values=names, state="readonly", width=75)
        cbx.pack()

        ttk.Label(win, text="工具路径").pack()
        e_exe = ttk.Entry(win, width=80)
        e_exe.pack()

        ttk.Label(win, text="参数").pack()
        e_param = ttk.Entry(win, width=80)
        e_param.pack()

        ttk.Label(win, text="Log路径").pack()
        e_log = ttk.Entry(win, width=80)
        e_log.pack()

        ttk.Label(win, text="运行目录").pack()
        e_cwd = ttk.Entry(win, width=80)
        e_cwd.pack()

        def load():
            n = var.get()
            d = self.tools[n]
            e_exe.delete(0,tk.END)
            e_exe.insert(0, d["exe"])
            e_param.delete(0,tk.END)
            e_param.insert(0, d["param"])
            e_log.delete(0,tk.END)
            e_log.insert(0, d["log_path"])
            e_cwd.delete(0,tk.END)
            e_cwd.insert(0, d.get("cwd",""))

        def save():
            n = var.get()
            self.tools[n]["exe"] = e_exe.get().strip()
            self.tools[n]["param"] = e_param.get().strip()
            self.tools[n]["log_path"] = e_log.get().strip()
            self.tools[n]["cwd"] = e_cwd.get().strip()
            self.save_cfg()
            self.refresh_btn()
            win.destroy()

        cbx.bind("<<ComboboxSelected>>", lambda x:load())
        load()
        ttk.Button(win, text="保存", command=save).pack(pady=10)

    def del_win(self):
        if not self.tools:
            messagebox.showwarning("提示","无工具")
            return
        win = tk.Toplevel(self.root)
        win.title("删除")
        win.grab_set()
        names = list(self.tools.keys())
        var = tk.StringVar(value=names[0])
        ttk.Combobox(win, textvariable=var, values=names, state="readonly").pack(pady=10)
        def do():
            del self.tools[var.get()]
            self.save_cfg()
            self.refresh_btn()
            win.destroy()
        ttk.Button(win, text="删除", command=do).pack()

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()