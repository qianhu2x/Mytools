import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import os
import json

TOOLS = {
    "示例工具": {
        "exe": "python",
        "cwd": "D:/脚本目录",
        "param": "hsdes_ticket_creator.py --excel template.xlsx",
        "log_path": ""
    }
}

CFG_FILE = "tools_cfg.json"

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("工具启动器")
        self.root.geometry("720x550")
        self.root.resizable(True, True)

        self.tools = self.load_cfg()
        self.tool_order = list(self.tools.keys())
        self.drag_data = {"index": None, "widget": None}
        self.create_ui()

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
            ordered = {k: self.tools[k] for k in self.tool_order if k in self.tools}
            with open(CFG_FILE, "w", encoding="utf-8") as f:
                json.dump(ordered, f, ensure_ascii=False, indent=2)
        except:
            pass

    def create_ui(self):
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(main_container, text="工具列表（鼠标拖动排序）", font=("微软雅黑",14,"bold")).pack(pady=5)

        # 滚动区域（修复边框遮挡 + 不花屏）
        self.canvas = tk.Canvas(main_container, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(main_container, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scroll_frame = ttk.Frame(self.canvas)

        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=0, pady=0)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 底部固定按钮（永不遮挡）
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame, text="添加", command=self.add_win, width=14).pack(side=tk.LEFT, padx=20, expand=True)
        ttk.Button(btn_frame, text="编辑", command=self.edit_win, width=14).pack(side=tk.LEFT, padx=20, expand=True)
        ttk.Button(btn_frame, text="删除", command=self.del_win, width=14).pack(side=tk.LEFT, padx=20, expand=True)

        self.refresh_buttons()

    def refresh_buttons(self):
        for w in self.scroll_frame.winfo_children():
            w.destroy()

        if not self.tool_order:
            ttk.Label(self.scroll_frame, text="暂无工具").pack(pady=40)
            return

        self.btn_widgets = []
        for idx, name in enumerate(self.tool_order):
            if name not in self.tools:
                continue

            btn = ttk.Button(self.scroll_frame, text=name, width=60)
            btn.pack(pady=4, ipady=5, fill=tk.X, padx=5)
            btn.config(command=lambda n=name: self.run_tool(n))
            self.btn_widgets.append(btn)

            btn.bind("<Button-1>", lambda e, i=idx, w=btn: self._on_press(e, i, w))
            btn.bind("<B1-Motion>", self._on_drag)
            btn.bind("<ButtonRelease-1>", self._on_release)

    def _on_press(self, event, index, widget):
        self.drag_data["index"] = index
        self.drag_data["widget"] = widget
        self._start_y = event.y_root

    def _on_drag(self, event):
        w = self.drag_data["widget"]
        if not w:
            return
        delta = event.y_root - self._start_y
        current_y = w.winfo_y()
        w.place(y=current_y + delta)
        self._start_y = event.y_root

    def _on_release(self, event):
        idx = self.drag_data["index"]
        w = self.drag_data["widget"]
        if not w or idx is None:
            return

        y = w.winfo_rooty()
        target_idx = len(self.btn_widgets)
        for i, btn in enumerate(self.btn_widgets):
            if btn.winfo_rooty() > y:
                target_idx = i
                break

        name = self.tool_order.pop(idx)
        self.tool_order.insert(target_idx, name)
        self.drag_data = {"index": None, "widget": None}
        self.save_cfg()
        self.refresh_buttons()

    def run_tool(self, name):
        try:
            cfg = self.tools[name]
            exe = cfg.get("exe", "").strip()
            cwd = cfg.get("cwd", "").strip()
            param = cfg.get("param", "").strip()
            log = cfg.get("log_path", "").strip()
            cmd = f"{exe} {param} {log}".strip()
            run_dir = cwd if cwd else None
            subprocess.Popen(cmd, shell=True, cwd=run_dir)
        except:
            messagebox.showerror("错误", "启动失败")

    # ==================== 添加窗口（超宽） ====================
    def add_win(self):
        top = tk.Toplevel(self.root)
        top.title("添加工具")
        top.geometry("780x420")
        top.resizable(False, False)
        top.grab_set()

        ttk.Label(top, text="工具名称").pack(pady=4)
        e_name = ttk.Entry(top, width=90)
        e_name.pack()

        ttk.Label(top, text="python（optional）").pack(pady=4)
        e_exe = ttk.Entry(top, width=90)
        e_exe.pack()

        ttk.Label(top, text="运行目录").pack(pady=4)
        e_cwd = ttk.Entry(top, width=90)
        e_cwd.pack()

        ttk.Label(top, text="参数").pack(pady=4)
        e_param = ttk.Entry(top, width=90)
        e_param.pack()

        ttk.Label(top, text="log路径（optional）").pack(pady=4)
        e_log = ttk.Entry(top, width=90)
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
            if n not in self.tool_order:
                self.tool_order.append(n)
            self.save_cfg()
            self.refresh_buttons()
            top.destroy()

        ttk.Button(top, text="确认添加", command=ok).pack(pady=18)

    # ==================== 编辑窗口（超宽） ====================
    def edit_win(self):
        if not self.tools:
            messagebox.showwarning("提示", "无工具")
            return

        top = tk.Toplevel(self.root)
        top.title("编辑工具")
        top.geometry("780x420")
        top.resizable(False, False)
        top.grab_set()

        var = tk.StringVar(value=self.tool_order[0])
        ttk.Label(top, text="选择工具").pack(pady=4)
        cbx = ttk.Combobox(top, textvariable=var, values=self.tool_order, state="readonly", width=85)
        cbx.pack()

        ttk.Label(top, text="python（optional）").pack(pady=4)
        e_exe = ttk.Entry(top, width=90)
        e_exe.pack()

        ttk.Label(top, text="运行目录").pack(pady=4)
        e_cwd = ttk.Entry(top, width=90)
        e_cwd.pack()

        ttk.Label(top, text="参数").pack(pady=4)
        e_param = ttk.Entry(top, width=90)
        e_param.pack()

        ttk.Label(top, text="log路径（optional）").pack(pady=4)
        e_log = ttk.Entry(top, width=90)
        e_log.pack()

        def load():
            d = self.tools[var.get()]
            e_exe.delete(0, tk.END)
            e_exe.insert(0, d.get("exe", ""))
            e_cwd.delete(0, tk.END)
            e_cwd.insert(0, d.get("cwd", ""))
            e_param.delete(0, tk.END)
            e_param.insert(0, d.get("param", ""))
            e_log.delete(0, tk.END)
            e_log.insert(0, d.get("log_path", ""))

        def save():
            n = var.get()
            self.tools[n]["exe"] = e_exe.get().strip()
            self.tools[n]["cwd"] = e_cwd.get().strip()
            self.tools[n]["param"] = e_param.get().strip()
            self.tools[n]["log_path"] = e_log.get().strip()
            self.save_cfg()
            self.refresh_buttons()
            top.destroy()

        cbx.bind("<<ComboboxSelected>>", lambda x: load())
        load()
        ttk.Button(top, text="保存修改", command=save).pack(pady=18)

    # ==================== 删除窗口（加宽） ====================
    def del_win(self):
        if not self.tools:
            messagebox.showwarning("提示", "无工具")
            return

        top = tk.Toplevel(self.root)
        top.title("删除工具")
        top.geometry("680x180")
        top.grab_set()

        var = tk.StringVar(value=self.tool_order[0])
        ttk.Label(top, text="选择要删除的工具").pack(pady=8)
        cbx = ttk.Combobox(top, textvariable=var, values=self.tool_order, state="readonly", width=75)
        cbx.pack()

        def do():
            n = var.get()
            if n in self.tools:
                del self.tools[n]
            if n in self.tool_order:
                self.tool_order.remove(n)
            self.save_cfg()
            self.refresh_buttons()
            top.destroy()

        ttk.Button(top, text="确认删除", command=do).pack(pady=12)

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()