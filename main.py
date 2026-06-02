# -*- coding: utf-8 -*-
"""
利润统计工具 GUI

使用 Tkinter 构建的本地桌面界面。
启动方式: python main.py
"""

import os
import sys
import subprocess
import threading
from datetime import datetime
from tkinter import Tk, Frame, Label, Button, Entry, Text, Scrollbar, Listbox, MULTIPLE
from tkinter import filedialog, messagebox, END, VERTICAL, HORIZONTAL, BOTH, LEFT, RIGHT, Y, X, W, E, N, S
from tkinter.ttk import Progressbar, Separator, Combobox

# 将当前目录加入 path，确保能导入 profit_core
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from profit_core import run_calculation, parse_date, get_unique_statuses

# 脚本所在目录，作为文件选择对话框的默认起始目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


class ProfitApp:
    def __init__(self, root):
        self.root = root
        self.root.title("利润统计工具 v2.0")
        self.root.geometry("780x680")
        self.root.minsize(650, 550)

        # 状态变量
        self.order_paths = []  # 订单表路径列表（支持多个）
        self.cost_path = None
        self.output_dir = None
        self.last_output_path = None  # 计算成功后记录输出文件路径
        self.last_dir = BASE_DIR  # 记住上次选择的目录

        self._build_ui()
        self._auto_detect_files()

    # ============================================================
    # UI 构建
    # ============================================================

    def _build_ui(self):
        main_frame = Frame(self.root, padx=18, pady=12)
        main_frame.pack(fill=BOTH, expand=True)

        # ---- 标题 ----
        title = Label(main_frame, text="利润统计工具",
                      font=("Microsoft YaHei", 16, "bold"), fg="#333333")
        title.pack(anchor=W, pady=(0, 15))

        # ---- 文件选择区域 ----
        file_frame = Frame(main_frame)
        file_frame.pack(fill=X, pady=(0, 5))

        # 订单表（支持多文件）
        row0_label = Frame(file_frame)
        row0_label.pack(fill=X, pady=(3, 0))
        Label(row0_label, text="订单表：", width=10, anchor=W,
              font=("Microsoft YaHei", 10)).pack(side=LEFT)
        Label(row0_label, text="（支持多选，共用同一成本表）",
              fg="#999999", font=("Microsoft YaHei", 8)).pack(side=LEFT)

        row0 = Frame(file_frame)
        row0.pack(fill=X, pady=3)

        # 左侧：文件列表
        list_frame = Frame(row0)
        list_frame.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))

        self.order_listbox = Listbox(list_frame, height=4, font=("Microsoft YaHei", 9))
        self.order_listbox.pack(side=LEFT, fill=X, expand=True)

        order_scroll = Scrollbar(list_frame, orient=VERTICAL, command=self.order_listbox.yview)
        order_scroll.pack(side=RIGHT, fill=Y)
        self.order_listbox.config(yscrollcommand=order_scroll.set)

        # 右侧：操作按钮
        order_btn_frame = Frame(row0)
        order_btn_frame.pack(side=LEFT)
        Button(order_btn_frame, text="添加文件", command=self._add_order_files,
               width=9, font=("Microsoft YaHei", 9)).pack(pady=(0, 2))
        Button(order_btn_frame, text="移除选中", command=self._remove_order_files,
               width=9, font=("Microsoft YaHei", 9)).pack()

        # 成本表
        row1 = Frame(file_frame)
        row1.pack(fill=X, pady=3)
        Label(row1, text="成本表：", width=10, anchor=W,
              font=("Microsoft YaHei", 10)).pack(side=LEFT)
        self.cost_entry = Entry(row1, font=("Microsoft YaHei", 9))
        self.cost_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
        Button(row1, text="选择文件", command=self._select_cost,
               width=9, font=("Microsoft YaHei", 9)).pack(side=LEFT)

        # 输出目录
        row2 = Frame(file_frame)
        row2.pack(fill=X, pady=3)
        Label(row2, text="输出目录：", width=10, anchor=W,
              font=("Microsoft YaHei", 10)).pack(side=LEFT)
        self.output_entry = Entry(row2, font=("Microsoft YaHei", 9))
        self.output_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
        Button(row2, text="选择目录", command=self._select_output_dir,
               width=9, font=("Microsoft YaHei", 9)).pack(side=LEFT)

        # 分隔线
        Separator(main_frame, orient=HORIZONTAL).pack(fill=X, pady=(10, 5))

        # ---- 日期区域 ----
        date_frame = Frame(main_frame)
        date_frame.pack(fill=X, pady=(5, 5))

        Label(date_frame, text="日期筛选（可选）：", anchor=W,
              font=("Microsoft YaHei", 10, "bold"), fg="#555555").pack(anchor=W)

        date_inner = Frame(date_frame)
        date_inner.pack(fill=X, pady=(5, 0))

        Label(date_inner, text="开始日期：", width=10, anchor=W,
              font=("Microsoft YaHei", 10)).grid(row=0, column=0, sticky=W, pady=3)
        self.start_date_entry = Entry(date_inner, width=18, font=("Microsoft YaHei", 9))
        self.start_date_entry.grid(row=0, column=1, sticky=W, padx=(0, 15))
        Label(date_inner, text="YYYY-MM-DD", fg="#999999",
              font=("Microsoft YaHei", 9)).grid(row=0, column=2, sticky=W)

        Label(date_inner, text="结束日期：", width=10, anchor=W,
              font=("Microsoft YaHei", 10)).grid(row=1, column=0, sticky=W, pady=3)
        self.end_date_entry = Entry(date_inner, width=18, font=("Microsoft YaHei", 9))
        self.end_date_entry.grid(row=1, column=1, sticky=W, padx=(0, 15))
        Label(date_inner, text="YYYY-MM-DD", fg="#999999",
              font=("Microsoft YaHei", 9)).grid(row=1, column=2, sticky=W)

        # 分隔线
        Separator(main_frame, orient=HORIZONTAL).pack(fill=X, pady=(10, 5))

        # ---- 订单状态区域 ----
        status_frame = Frame(main_frame)
        status_frame.pack(fill=X, pady=(5, 5))

        Label(status_frame, text="订单状态筛选（可选）：", anchor=W,
              font=("Microsoft YaHei", 10, "bold"), fg="#555555").pack(anchor=W)

        status_inner = Frame(status_frame)
        status_inner.pack(fill=X, pady=(5, 0))

        Label(status_inner, text="订单状态：", width=10, anchor=W,
              font=("Microsoft YaHei", 10)).grid(row=0, column=0, sticky=W, pady=3)

        # 状态列表（多选）
        status_list_frame = Frame(status_inner)
        status_list_frame.grid(row=0, column=1, sticky=W, padx=(0, 5))

        self.status_listbox = Listbox(status_list_frame, width=16, height=5,
                                       font=("Microsoft YaHei", 9),
                                       selectmode=MULTIPLE, exportselection=False)
        self.status_listbox.pack(side=LEFT, fill=Y)

        status_scroll = Scrollbar(status_list_frame, orient=VERTICAL, command=self.status_listbox.yview)
        status_scroll.pack(side=RIGHT, fill=Y)
        self.status_listbox.config(yscrollcommand=status_scroll.set)

        # 全选/取消按钮
        status_btn_frame = Frame(status_inner)
        status_btn_frame.grid(row=0, column=2, sticky=W)
        Button(status_btn_frame, text="全选", command=self._select_all_statuses,
               width=5, font=("Microsoft YaHei", 8)).pack(pady=(0, 2))
        Button(status_btn_frame, text="取消", command=self._deselect_all_statuses,
               width=5, font=("Microsoft YaHei", 8)).pack()

        Label(status_inner, text="可多选（Ctrl+点击），不选=全部",
              fg="#999999", font=("Microsoft YaHei", 9)).grid(row=0, column=3, sticky=W)

        # 分隔线
        Separator(main_frame, orient=HORIZONTAL).pack(fill=X, pady=(10, 5))

        # ---- 按钮区域 ----
        btn_frame = Frame(main_frame)
        btn_frame.pack(fill=X, pady=(5, 8))

        self.calc_btn = Button(btn_frame, text="▶  开始计算", command=self._start_calculation,
                               bg="#4472C4", fg="white",
                               font=("Microsoft YaHei", 11, "bold"),
                               height=1, width=14, cursor="hand2")
        self.calc_btn.pack(side=LEFT, padx=(0, 8))

        self.open_btn = Button(btn_frame, text="📂 打开输出文件", command=self._open_output,
                               bg="#548235", fg="white",
                               font=("Microsoft YaHei", 10),
                               height=1, width=14, cursor="hand2", state='disabled')
        self.open_btn.pack(side=LEFT, padx=(0, 8))

        self.clear_btn = Button(btn_frame, text="清空日志", command=self._clear_log,
                                font=("Microsoft YaHei", 10),
                                height=1, width=10)
        self.clear_btn.pack(side=LEFT)

        self.progress = Progressbar(btn_frame, mode='indeterminate', length=180)
        # 默认隐藏，计算时显示

        # ---- 统计摘要区域 ----
        self.stats_frame = Frame(main_frame, bg="#F2F2F2", padx=10, pady=5)
        # 默认隐藏，计算完成后显示

        # ---- 日志输出 ----
        log_label_frame = Frame(main_frame)
        log_label_frame.pack(fill=X, pady=(5, 2))
        Label(log_label_frame, text="运行日志：", anchor=W,
              font=("Microsoft YaHei", 10, "bold"), fg="#555555").pack(side=LEFT)

        log_container = Frame(main_frame)
        log_container.pack(fill=BOTH, expand=True)

        self.log_text = Text(log_container, wrap='word',
                             font=("Consolas", 9),
                             bg="#1E1E1E", fg="#D4D4D4",
                             insertbackground="white",
                             padx=8, pady=5)
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)

        scrollbar = Scrollbar(log_container, orient=VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        # 底部状态栏
        status_frame = Frame(main_frame, bg="#E8E8E8", height=22)
        status_frame.pack(fill=X, pady=(5, 0))
        self.status_label = Label(status_frame, text="就绪",
                                  bg="#E8E8E8", fg="#666666",
                                  font=("Microsoft YaHei", 8), anchor=W)
        self.status_label.pack(fill=X, padx=8, pady=2)

        # 欢迎日志
        self._log("欢迎使用利润统计工具 v2.1")
        self._log("请添加订单表（支持多个）、选择成本表及输出目录，然后点击「开始计算」。")
        self._log("多个订单表将使用同一成本表合并计算。")
        self._log("日期筛选为可选项，不填则统计全部数据。")
        self._log("订单状态支持多选（Ctrl+点击），不选=全部状态。")

    # ============================================================
    # 自动检测文件
    # ============================================================

    def _auto_detect_files(self):
        """在当前目录下自动检测可能的订单表和成本表"""
        detected_order = None
        detected_cost = None

        for f in os.listdir(BASE_DIR):
            if not f.endswith(('.xlsx', '.xls')):
                continue
            full = os.path.join(BASE_DIR, f)
            if not os.path.isfile(full):
                continue
            lower = f.lower()
            # 排除 output 目录下的文件
            if 'output' in lower or '利润统计结果' in lower:
                continue
            # 成本表关键词
            if any(kw in lower for kw in ['成本', 'cost']):
                detected_cost = full
            else:
                detected_order = full

        if detected_order:
            self.order_paths.append(detected_order)
            self.order_listbox.insert(END, detected_order)
            self._log(f"[自动检测] 订单表: {os.path.basename(detected_order)}")

        if detected_cost:
            self.cost_path = detected_cost
            self.cost_entry.delete(0, END)
            self.cost_entry.insert(0, detected_cost)
            self._log(f"[自动检测] 成本表: {os.path.basename(detected_cost)}")

        # 自动填充输出目录
        default_output = os.path.join(BASE_DIR, 'output')
        self.output_dir = default_output
        self.output_entry.delete(0, END)
        self.output_entry.insert(0, default_output)

        # 自动加载订单状态
        if self.order_paths:
            self._load_order_statuses()

    # ============================================================
    # 文件选择
    # ============================================================

    def _load_order_statuses(self):
        """从所有已选订单文件中读取可用的订单状态列表，合并后填充到多选列表"""
        all_statuses = set()
        for order_path in self.order_paths:
            if not order_path or not os.path.exists(order_path):
                continue
            try:
                statuses = get_unique_statuses(order_path)
                if statuses:
                    all_statuses.update(statuses)
            except Exception as e:
                self._log(f"[状态加载] 读取 {os.path.basename(order_path)} 状态失败: {e}")

        # 更新列表
        self.status_listbox.delete(0, END)
        if all_statuses:
            for s in sorted(all_statuses):
                self.status_listbox.insert(END, s)
            self._log(f"[状态加载] 可用订单状态: {', '.join(sorted(all_statuses))}")
        else:
            self._log("[状态加载] 订单表中未找到「订单状态」列")

    def _add_order_files(self):
        """添加订单表文件（支持多选）"""
        paths = filedialog.askopenfilenames(
            title="选择订单表（可多选）",
            initialdir=self.last_dir,
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        for path in paths:
            if path and path not in self.order_paths:
                self.order_paths.append(path)
                self.order_listbox.insert(END, path)
                self._set_status(f"已添加订单表: {os.path.basename(path)}")
        if paths:
            self.last_dir = os.path.dirname(paths[0])
            self._load_order_statuses()

    def _remove_order_files(self):
        """移除选中的订单表文件"""
        selected = self.order_listbox.curselection()
        if not selected:
            messagebox.showinfo("提示", "请先在列表中选择要移除的文件")
            return
        # 从后往前删除，避免索引错乱
        for idx in reversed(selected):
            path = self.order_listbox.get(idx)
            self.order_listbox.delete(idx)
            if path in self.order_paths:
                self.order_paths.remove(path)
            self._set_status(f"已移除: {os.path.basename(path)}")
        self._load_order_statuses()

    def _select_all_statuses(self):
        """全选所有订单状态"""
        self.status_listbox.select_set(0, END)

    def _deselect_all_statuses(self):
        """取消全选订单状态"""
        self.status_listbox.selection_clear(0, END)

    def _select_cost(self):
        path = filedialog.askopenfilename(
            title="选择成本表",
            initialdir=self.last_dir,
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.cost_path = path
            self.last_dir = os.path.dirname(path)
            self.cost_entry.delete(0, END)
            self.cost_entry.insert(0, path)
            self._set_status(f"已选择成本表: {os.path.basename(path)}")

    def _select_output_dir(self):
        path = filedialog.askdirectory(
            title="选择输出目录",
            initialdir=self.output_dir if self.output_dir else self.last_dir
        )
        if path:
            self.output_dir = path
            self.output_entry.delete(0, END)
            self.output_entry.insert(0, path)
            self._set_status(f"输出目录: {path}")

    # ============================================================
    # 日志与状态
    # ============================================================

    def _log(self, msg):
        self.log_text.insert(END, msg + "\n")
        self.log_text.see(END)
        self.root.update_idletasks()

    def _clear_log(self):
        self.log_text.delete(1.0, END)
        self._set_status("日志已清空")

    def _set_status(self, text):
        self.status_label.config(text=text)

    # ============================================================
    # 打开输出文件
    # ============================================================

    def _open_output(self):
        if not self.last_output_path or not os.path.exists(self.last_output_path):
            messagebox.showwarning("提示", "输出文件不存在，请先执行计算。")
            return
        try:
            if sys.platform == 'win32':
                os.startfile(self.last_output_path)
            else:
                subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open',
                                  self.last_output_path])
            self._set_status(f"已打开: {os.path.basename(self.last_output_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件:\n{e}")

    # ============================================================
    # 计算流程
    # ============================================================

    def _start_calculation(self):
        # 收集订单表路径
        order_paths = list(self.order_paths)  # 复制一份，避免计算中途被修改
        cost_path = self.cost_path or self.cost_entry.get().strip()
        output_dir = self.output_dir or self.output_entry.get().strip()

        # 验证输入
        if not order_paths:
            messagebox.showwarning("提示", "请先添加至少一个订单表文件")
            return
        if not cost_path:
            messagebox.showwarning("提示", "请先选择成本表文件")
            return
        if not output_dir:
            messagebox.showwarning("提示", "请先选择输出目录")
            return
        for op in order_paths:
            if not os.path.exists(op):
                messagebox.showerror("错误", f"订单表文件不存在:\n{op}")
                return
        if not os.path.exists(cost_path):
            messagebox.showerror("错误", f"成本表文件不存在:\n{cost_path}")
            return

        # 解析日期
        start_date = None
        end_date = None
        start_str = self.start_date_entry.get().strip()
        end_str = self.end_date_entry.get().strip()

        if start_str:
            try:
                start_date = parse_date(start_str)
            except ValueError as e:
                messagebox.showerror("日期格式错误", str(e))
                return

        if end_str:
            try:
                end_date = parse_date(end_str)
                end_date = end_date.replace(hour=23, minute=59, second=59)
            except ValueError as e:
                messagebox.showerror("日期格式错误", str(e))
                return

        # 获取订单状态筛选（多选列表）
        selected_indices = self.status_listbox.curselection()
        if selected_indices:
            order_statuses = [self.status_listbox.get(i) for i in selected_indices]
        else:
            order_statuses = None  # 未选择任何状态 = 全部

        # 清空日志和上次结果
        self.log_text.delete(1.0, END)
        self._hide_stats()
        self.last_output_path = None
        self.open_btn.config(state='disabled')

        # 禁用按钮，显示进度
        self.calc_btn.config(state='disabled', text="计算中...", bg="#8DB4E2")
        self.clear_btn.config(state='disabled')
        self.progress.pack(side=LEFT, padx=(12, 0))
        self.progress.start()
        self._set_status("正在计算...")

        # 后台线程执行计算
        def run():
            try:
                output_path, stats = run_calculation(
                    order_paths, cost_path, output_dir,
                    start_date=start_date, end_date=end_date,
                    order_statuses=order_statuses,
                    log_func=self._log
                )
                self.root.after(0, lambda: self._on_success(output_path, stats))
            except Exception as e:
                import traceback
                tb = traceback.format_exc()
                self.root.after(0, lambda: self._on_error(str(e), tb))

        threading.Thread(target=run, daemon=True).start()

    def _on_success(self, output_path, stats):
        self.progress.stop()
        self.progress.pack_forget()
        self.calc_btn.config(state='normal', text="▶  开始计算", bg="#4472C4")
        self.clear_btn.config(state='normal')

        self.last_output_path = output_path
        self.open_btn.config(state='normal')

        self._log(f"\n{'='*50}")
        self._log(f"✓ 计算成功完成!")
        self._log(f"输出文件: {output_path}")
        self._set_status(f"计算完成 — {os.path.basename(output_path)}")

        # 显示统计摘要
        self._show_stats(stats)

        messagebox.showinfo(
            "计算完成",
            f"利润统计已完成!\n\n"
            f"输出文件:\n{output_path}\n\n"
            f"订单表行数: {stats['订单表行数']}\n"
            f"成本表行数: {stats['成本表行数']}\n"
            f"正常参与计算行数: {stats['正常参与计算行数']}\n"
            f"异常行数: {stats['异常行数']}"
        )

    def _on_error(self, error_msg, traceback_str):
        self.progress.stop()
        self.progress.pack_forget()
        self.calc_btn.config(state='normal', text="▶  开始计算", bg="#4472C4")
        self.clear_btn.config(state='normal')

        self._log(f"\n{'='*50}")
        self._log(f"✗ 计算失败: {error_msg}")
        self._log(traceback_str)
        self._set_status("计算失败")
        messagebox.showerror("计算失败", f"计算过程中发生错误:\n\n{error_msg}")

    # ============================================================
    # 统计摘要面板
    # ============================================================

    def _show_stats(self, stats):
        self.stats_frame.pack(fill=X, pady=(0, 8), before=self.log_text.master)

        # 清除旧内容
        for w in self.stats_frame.winfo_children():
            w.destroy()

        Label(self.stats_frame, text="📊 统计摘要",
              font=("Microsoft YaHei", 10, "bold"),
              bg="#F2F2F2", fg="#333333").pack(anchor=W)

        items = [
            ("订单表行数", stats['订单表行数']),
            ("成本表行数", stats['成本表行数']),
            ("正常参与计算行数", stats['正常参与计算行数']),
            ("异常行数", stats['异常行数']),
        ]
        row = Frame(self.stats_frame, bg="#F2F2F2")
        row.pack(fill=X, pady=(4, 0))
        for i, (label, value) in enumerate(items):
            item = Frame(row, bg="#FFFFFF", padx=12, pady=6, relief="groove", bd=1)
            item.pack(side=LEFT, padx=(0, 10))
            Label(item, text=label, font=("Microsoft YaHei", 8),
                  bg="#FFFFFF", fg="#888888").pack()
            Label(item, text=str(value), font=("Microsoft YaHei", 14, "bold"),
                  bg="#FFFFFF", fg="#4472C4").pack()

    def _hide_stats(self):
        self.stats_frame.pack_forget()


# ============================================================
# 入口
# ============================================================

def main():
    root = Tk()
    # 尝试设置窗口图标（如无可忽略）
    try:
        root.iconbitmap(default='')
    except Exception:
        pass
    app = ProfitApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
