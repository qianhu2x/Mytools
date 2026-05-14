import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import os
import json

# ====================== 【在这里配置你的工具】 ======================
# 格式："工具名称": {"path":"路径", "args":"参数", "need_log":"True/False"}
TOOLS = {
    "示例 - 记事本": {"path": "C:/Windows/notepad.exe", "args": "", "need_log": "False"},
    "Log分析工具": {"path": "python", "args": "D:/tools/analyze_log.py", "need_log": "True"},
}
# ==================================================================

CONFIG_FILE = "tools_config.json"

class ToolLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("📌 通用工具启动器（支持Log分析）")
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

        for name, info in self.tools.items():
            btn = ttk.Button(
                self.button_frame,
                text=name,
                command=lambda n=name: self.launch_tool(n),
                width=30
            )
            btn.pack(pady=6, ipady=5)

    def launch_tool(self, name):
        try:
            info = self.tools[name]
            path = info["path"]
            args = info.get("args", "")
            need_log = info.get("need_log", "False")

            cmd = [path] + args.split()

            # ====================== 核心功能：选择Log文件 ======================
            if need_log == "True":
                log_path = filedialog.askopenfilename(
                    title="请选择要分析的Log文件",
                    filetypes=[("Log文件", "*.log *.txt *.json"), ("所有文件", "*.*")]
                )
                if not log_path:
                    messagebox.showwarning("取消", "未选择Log文件，已取消启动")
                    return
                cmd.append(log_path)  # 把log路径加到参数最后

            # 启动（完全沿用第一版逻辑，不闪退）
            subprocess.Popen(cmd, shell=True)
            messagebox.showinfo("成功", f"已启动：{name}\nLog文件：{log_path if need_log=='True' else '无'}")
        except Exception as e:
            messagebox.showerror("启动失败", f"错误：{str(e)}")

    def add_tool_window(self):
        top = tk.Toplevel(self.root)
        top.title("添加新工具")
        top.geometry("480x320")
        top.resizable(False, False)
        top.grab_set()

        ttk.Label(top, text="工具名称：", font=("微软雅黑", 10)).pack(pady=3)
        name_entry = ttk.Entry(top, width=45)
        name_entry.pack(pady=2)

        ttk.Label(top, text="工具路径：", font=("微软雅黑", 10)).pack(pady=3)
        path_entry = ttk.Entry(top, width=45)
        path_entry.pack(pady=2)

        ttk.Label(top, text="启动参数：", font=("微软雅黑", 10)).pack(pady=3)
        arg_entry = ttk.Entry(top, width=45)
        arg_entry.pack(pady=2)

        ttk.Label(top, text="需要选择Log文件？(True/False)：", font=("微软雅黑", 10)).pack(pady=3)
        log_entry = ttk.Entry(top, width=45)
        log_entry.insert(0, "False")
        log_entry.pack(pady=2)

        def confirm_add():
            name = name_entry.get().strip()
            path = path_entry.get().strip()
            arg = arg_entry.get().strip()
            need_log = log_entry.get().strip()
            if not name or not path:
                messagebox.showwarning("提示", "名称和路径不能为空！")
                return
            self.tools[name] = {
                "path": path,
                "args": arg,
                "need_log": need_log
            }
            self.save_tools()
            self.refresh_buttons()
            top.destroy()
            messagebox.showinfo("成功", "工具添加成功！")

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

        ttk.Label(top, text="选择要删除的工具：", font=("微软雅黑", 10)).pack(pady=10)
        tool_names = list(self.tools.keys())
        selected = tk.StringVar(value=tool_names[0])
        combo = ttk.Combobox(top, textvariable=selected, values=tool_names, state="readonly")
        combo.pack(pady=5)

        def confirm_delete():
            name = selected.get()
            del self.tools[name]
            self.save_tools()
            self.refresh_buttons()
            top.destroy()
            messagebox.showinfo("成功", f"已删除：{name}")

        ttk.Button(top, text="确认删除", command=confirm_delete).pack(pady=15)

if __name__ == "__main__":
    root = tk.Tk()
    app = ToolLauncher(root)
    root.mainloop()