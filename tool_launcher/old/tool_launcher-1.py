import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import os
import json

# ====================== 【在这里配置你的工具】 ======================
# 格式：{"工具名称": "工具完整路径"}
# 支持：exe、bat、py、sh、cmd 等所有可执行文件
TOOLS = {
    "示例 - 记事本": "C:/Windows/notepad.exe",
    "示例 - 计算器": "calc.exe",  # 系统命令可直接写名称
    # "我的数据处理工具": "D:/tools/data_tool.exe",
    # "测试脚本": "E:/scripts/test.py",
    # "批量重命名": "F:/utils/rename.bat",
}
# ==================================================================

# 工具配置文件（自动保存）
CONFIG_FILE = "tools_config.json"

class ToolLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("📌 通用工具启动器")
        self.root.geometry("500x500")
        self.root.resizable(False, False)
        
        # 加载配置
        self.tools = self.load_tools()
        
        # 创建界面
        self.create_widgets()

    def load_tools(self):
        """加载工具配置"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except:
            pass
        return TOOLS.copy()

    def save_tools(self):
        """保存工具配置"""
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.tools, f, ensure_ascii=False, indent=2)
        except:
            pass

    def create_widgets(self):
        """创建界面组件"""
        # 标题
        title_label = ttk.Label(
            self.root, 
            text="我的工具集", 
            font=("微软雅黑", 16, "bold")
        )
        title_label.pack(pady=15)

        # 工具按钮容器
        self.button_frame = ttk.Frame(self.root)
        self.button_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        # 刷新按钮列表
        self.refresh_buttons()

        # 底部操作栏
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
        """刷新工具按钮"""
        # 清空旧按钮
        for widget in self.button_frame.winfo_children():
            widget.destroy()

        if not self.tools:
            ttk.Label(
                self.button_frame, 
                text="暂无工具，请点击【添加工具】", 
                font=("微软雅黑", 11)
            ).pack(pady=20)
            return

        # 创建按钮
        for name, path in self.tools.items():
            btn = ttk.Button(
                self.button_frame,
                text=name,
                command=lambda p=path: self.launch_tool(p),
                width=30
            )
            btn.pack(pady=6, ipady=5)

    def launch_tool(self, tool_path):
        """启动工具"""
        try:
            # 跨平台启动，不阻塞窗口
            if os.name == "nt":  # Windows
                subprocess.Popen(tool_path, shell=True)
            else:  # Mac/Linux
                subprocess.Popen(["open", tool_path] if os.name == "posix" else ["xdg-open", tool_path])
            
            messagebox.showinfo("成功", f"已启动：\n{tool_path}")
        except Exception as e:
            messagebox.showerror("启动失败", f"错误信息：{str(e)}")

    def add_tool_window(self):
        """添加工具窗口"""
        top = tk.Toplevel(self.root)
        top.title("添加新工具")
        top.geometry("450x200")
        top.resizable(False, False)
        top.grab_set()  # 模态窗口

        ttk.Label(top, text="工具名称：", font=("微软雅黑", 10)).pack(pady=5)
        name_entry = ttk.Entry(top, width=40, font=("微软雅黑", 10))
        name_entry.pack(pady=2)

        ttk.Label(top, text="工具完整路径：", font=("微软雅黑", 10)).pack(pady=5)
        path_entry = ttk.Entry(top, width=40, font=("微软雅黑", 10))
        path_entry.pack(pady=2)

        def confirm_add():
            name = name_entry.get().strip()
            path = path_entry.get().strip()
            if not name or not path:
                messagebox.showwarning("提示", "名称和路径不能为空！")
                return
            self.tools[name] = path
            self.save_tools()
            self.refresh_buttons()
            top.destroy()
            messagebox.showinfo("成功", "工具添加成功！")

        ttk.Button(top, text="确认添加", command=confirm_add).pack(pady=15)

    def delete_tool_window(self):
        """删除工具窗口"""
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
        combo = ttk.Combobox(top, textvariable=selected, values=tool_names, width=30, state="readonly")
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