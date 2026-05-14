import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import os
import json

CFG_FILE = "tools_cfg.json"

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("工具启动器")
        self.geometry("750x550")
        self.resizable(True, True)

        self.tools = self.load_cfg()
        self.tool_list = list(self.tools.keys())

        # 主布局
        self.list_frame = ttk.Frame(self)
        self.list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # 滚动画布，关闭多余边框减少卡顿
        self.canvas = tk.Canvas(self.list_frame, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.list_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)

        self.inner.bind("<Configure>", self.on_frame_config)
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 绑定鼠标滚轮滚动
        self.canvas.bind_all("<MouseWheel>", self.on_mouse_wheel)

        # 底部固定按钮
        self.btn_frame = ttk.Frame(self)
        self.btn_frame.pack(fill=tk.X, padx=15, pady=10)
        ttk.Button(self.btn_frame, text="添加", command=self.add_win, width=12).pack(side=tk.LEFT, expand=True)
        ttk.Button(self.btn_frame, text="编辑", command=self.edit_win, width=12).pack(side=tk.LEFT, expand=True)
        ttk.Button(self.btn_frame, text="删除", command=self.del_win, width=12).pack(side=tk.LEFT, expand=True)

        self.refresh_list()

    def on_frame_config(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_mouse_wheel(self, event):
        # 适配Windows滚轮
        delta = -1 * (event.delta // 120)
        self.canvas.yview_scroll(delta, "units")

    def load_cfg(self):
        try:
            if os.path.exists(CFG_FILE):
                with open(CFG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except:
            pass
        return {}

    def save_cfg(self):
        try:
            od = {k: self.tools[k] for k in self.tool_list if k in self.tools}
            with open(CFG_FILE, "w", encoding="utf-8") as f:
                json.dump(od, f, ensure_ascii=False, indent=2)
        except:
            pass

    def refresh_list(self):
        # 清空旧控件，不做多余重绘，降低CPU占用
        for w in self.inner.winfo_children():
            w.destroy()

        if not self.tool_list:
            ttk.Label(self.inner, text="暂无工具").pack(pady=40)
            return

        for name in self.tool_list:
            if name not in self.tools:
                continue
            btn = ttk.Button(self.inner, text=name, width=65,
                             command=lambda n=name: self.run_tool(n))
            btn.pack(fill=tk.X, pady=3, ipady=4)

    def run_tool(self, name):
        try:
            d = self.tools[name]
            exe = d.get("exe", "").strip()
            cwd = d.get("cwd", "").strip()
            param = d.get("param", "").strip()
            log = d.get("log_path", "").strip()
            cmd = f"{exe} {param} {log}".strip()
            subprocess.Popen(cmd, shell=True, cwd=cwd if cwd else None)
        except:
            messagebox.showerror("错误", "启动失败")

    def add_win(self):
        win = tk.Toplevel(self)
        win.title("添加工具")
        win.geometry("780x400")
        win.resizable(False, False)
        win.grab_set()

        ttk.Label(win, text="工具名称").pack()
        e_name = ttk.Entry(win, width=90)
        e_name.pack()

        ttk.Label(win, text="python（optional）").pack()
        e_exe = ttk.Entry(win, width=90)
        e_exe.pack()

        ttk.Label(win, text="运行目录").pack()
        e_cwd = ttk.Entry(win, width=90)
        e_cwd.pack()

        ttk.Label(win, text="参数").pack()
        e_param = ttk.Entry(win, width=90)
        e_param.pack()

        ttk.Label(win, text="log路径（optional）").pack()
        e_log = ttk.Entry(win, width=90)
        e_log.pack()

        def ok():
            n = e_name.get().strip()
            if not n:
                messagebox.showwarning("提示", "名称不能为空")
                return
            self.tools[n] = {
                "exe": e_exe.get().strip(),
                "cwd": e_cwd.get().strip(),
                "param": e_param.get().strip(),
                "log_path": e_log.get().strip()
            }
            if n not in self.tool_list:
                self.tool_list.append(n)
            self.save_cfg()
            self.refresh_list()
            win.destroy()

        ttk.Button(win, text="确定", command=ok).pack(pady=15)

    def edit_win(self):
        if not self.tool_list:
            messagebox.showwarning("提示", "无工具可编辑")
            return
        win = tk.Toplevel(self)
        win.title("编辑工具")
        win.geometry("780x400")
        win.resizable(False, False)
        win.grab_set()

        var = tk.StringVar(value=self.tool_list[0])
        ttk.Label(win, text="选择工具").pack()
        cbx = ttk.Combobox(win, textvariable=var, values=self.tool_list, state="readonly", width=85)
        cbx.pack()

        ttk.Label(win, text="python（optional）").pack()
        e_exe = ttk.Entry(win, width=90)
        e_exe.pack()

        ttk.Label(win, text="运行目录").pack()
        e_cwd = ttk.Entry(win, width=90)
        e_cwd.pack()

        ttk.Label(win, text="参数").pack()
        e_param = ttk.Entry(win, width=90)
        e_param.pack()

        ttk.Label(win, text="log路径（optional）").pack()
        e_log = ttk.Entry(win, width=90)
        e_log.pack()

        def load_item():
            d = self.tools[var.get()]
            e_exe.delete(0, tk.END)
            e_exe.insert(0, d.get("exe", ""))
            e_cwd.delete(0, tk.END)
            e_cwd.insert(0, d.get("cwd", ""))
            e_param.delete(0, tk.END)
            e_param.insert(0, d.get("param", ""))
            e_log.delete(0, tk.END)
            e_log.insert(0, d.get("log_path", ""))

        def save_item():
            n = var.get()
            self.tools[n]["exe"] = e_exe.get().strip()
            self.tools[n]["cwd"] = e_cwd.get().strip()
            self.tools[n]["param"] = e_param.get().strip()
            self.tools[n]["log_path"] = e_log.get().strip()
            self.save_cfg()
            self.refresh_list()
            win.destroy()

        cbx.bind("<<ComboboxSelected>>", lambda e: load_item())
        load_item()
        ttk.Button(win, text="保存", command=save_item).pack(pady=15)

    def del_win(self):
        if not self.tool_list:
            messagebox.showwarning("提示", "无工具可删除")
            return
        win = tk.Toplevel(self)
        win.title("删除工具")
        win.geometry("680x180")
        win.grab_set()

        var = tk.StringVar(value=self.tool_list[0])
        ttk.Label(win, text="选择要删除的工具").pack(pady=8)
        cbx = ttk.Combobox(win, textvariable=var, values=self.tool_list, state="readonly", width=75)
        cbx.pack()

        def do_del():
            n = var.get()
            if n in self.tools:
                del self.tools[n]
            if n in self.tool_list:
                self.tool_list.remove(n)
            self.save_cfg()
            self.refresh_list()
            win.destroy()

        ttk.Button(win, text="确认删除", command=do_del).pack(pady=12)

if __name__ == "__main__":
    app = App()
    app.mainloop()