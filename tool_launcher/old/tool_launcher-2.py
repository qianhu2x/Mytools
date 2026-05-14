import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import os
import json

# ====================== 【在这里配置你的工具】 ======================
# 格式："名称": {"path":"程序路径", "args":"参数"}
TOOLS = {
    "示例 - 记事本": {"path":"C:/Windows/notepad.exe", "args":""},
    "示例 - 计算器": {"path":"calc.exe", "args":""},
}
# ==================================================================

CONFIG_FILE = "tools_config.json"

class ToolLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("📌 通用工具启动器（支持参数）")
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

        # 新加：编辑参数按钮
        ttk.Button(
            control_frame,
            text="编辑参数",
            command=self.edit_arg_window
        ).grid(row=0, column=2, padx=10)

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
        """沿用第一版启动逻辑，只加参数拼接，不改动底层防止闪退"""
        try:
            info = self.tools[name]
            path = info["path"]
            args = info.get("args", "").split()
            cmd = [path] + args

            # 用最兼容第一版的启动方式，不搞新标志，防闪退
            subprocess.Popen(cmd, shell=True)
            messagebox.showinfo("成功", f"已启动：{name}")
        except Exception as e:
            messagebox.showerror("启动失败", f"错误：{str(e)}")

    def add_tool_window(self):
        top = tk.Toplevel(self.root)
        top.title("添加新工具")
        top.geometry("450x260")
        top.resizable(False, False)
        top.grab_set()

        ttk.Label(top, text="工具名称：", font=("微软雅黑", 10)).pack(pady=5)
        name_entry = ttk.Entry(top, width=40, font=("微软雅黑", 10))
        name_entry.pack(pady=2)

        ttk.Label(top, text="工具完整路径：", font=("微软雅黑", 10)).pack(pady=5)
        path_entry = ttk.Entry(top, width=40, font=("微软雅黑", 10))
        path_entry.pack(pady=2)

        # 新增参数输入框
        ttk.Label(top, text="启动参数（空格分隔）：", font=("微软雅黑", 10)).pack(pady=5)
        arg_entry = ttk.Entry(top, width=40, font=("微软雅黑", 10))
        arg_entry.pack(pady=2)

        def confirm_add():
            name = name_entry.get().strip()
            path = path_entry.get().strip()
            arg = arg_entry.get().strip()
            if not name or not path:
                messagebox.showwarning("提示", "名称和路径不能为空！")
                return
            self.tools[name] = {"path":path, "args":arg}
            self.save_tools()
            self.refresh_buttons()
            top.destroy()
            messagebox.showinfo("成功", "工具添加成功！")

        ttk.Button(top, text="确认添加", command=confirm_add).pack(pady=15)

    def edit_arg_window(self):
        if not self.tools:
            messagebox.showwarning("提示", "没有工具可编辑！")
            return

        top = tk.Toplevel(self.root)
        top.title="编辑工具参数"
        top.geometry("450x220")
        top.resizable(False, False)
        top.grab_set()

        ttk.Label(top, text="选择工具：", font=("微软雅黑", 10)).pack(pady=10)
        tool_names = list(self.tools.keys())
        selected = tk.StringVar(value=tool_names[0])
        combo = ttk.Combobox(top, textvariable=selected, values=tool_names, state="readonly", width=30)
        combo.pack(pady=5)

        ttk.Label(top, text="修改参数：", font=("微软雅黑", 10)).pack(pady=5)
        arg_entry = ttk.Entry(top, width=40)
        arg_entry.pack(pady=2)

        def load_arg(event):
            name = selected.get()
            arg_entry.delete(0, tk.END)
            arg_entry.insert(0, self.tools[name].get("args",""))

        def save_arg():
            name = selected.get()
            new_arg = arg_entry.get().strip()
            self.tools[name]["args"] = new_arg
            self.save_tools()
            top.destroy()
            messagebox.showinfo("成功", "参数已保存！")

        combo.bind("<<ComboboxSelected>>", load_arg)
        load_arg(None)
        ttk.Button(top, text="保存参数", command=save_arg).pack(pady=10)

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