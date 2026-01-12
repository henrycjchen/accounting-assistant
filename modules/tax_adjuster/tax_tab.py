#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调整税负率 Tab
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES

from .adjust_tax import TaxAdjuster


class TaxAdjustTab(ttk.Frame):
    """调整税负率 Tab"""

    def __init__(self, parent):
        super().__init__(parent, padding="10")

        self.adjuster = None
        self.result = None

        # 文件路径（完整路径）
        self.excel_file_path = ""
        # 显示用的文件名
        self.file_display = tk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        """设置界面"""
        # 标题
        title_label = ttk.Label(
            self,
            text="调整税负率",
            font=("Microsoft YaHei", 16, "bold")
        )
        title_label.pack(pady=(0, 20))

        # === 文件选择区域 ===
        file_frame = ttk.LabelFrame(self, text="选择文件", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))

        # 文件选择行
        row_frame = ttk.Frame(file_frame)
        row_frame.pack(fill=tk.X, pady=(0, 10))

        label = ttk.Label(row_frame, text="测算表：", width=15)
        label.pack(side=tk.LEFT)

        self.file_entry = ttk.Entry(row_frame, textvariable=self.file_display, width=40, state="readonly")
        self.file_entry.pack(side=tk.LEFT, padx=(0, 10))

        # 拖放文件支持
        self.file_entry.drop_target_register(DND_FILES)
        self.file_entry.dnd_bind('<<Drop>>', self.on_drop)

        btn = ttk.Button(row_frame, text="选择", command=self.browse_file, width=8)
        btn.pack(side=tk.LEFT)

        # === 参数设置区域 ===
        param_frame = ttk.LabelFrame(self, text="参数设置", padding="10")
        param_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(param_frame, text="目标税负率:").pack(side=tk.LEFT)
        self.rate_var = tk.StringVar(value="0.00414")
        rate_entry = ttk.Entry(param_frame, textvariable=self.rate_var, width=15)
        rate_entry.pack(side=tk.LEFT, padx=(5, 10))
        ttk.Label(param_frame, text="(例: 0.00414 = 0.414%)").pack(side=tk.LEFT)

        # === 操作按钮 ===
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(btn_frame, text="计算调整方案", command=self.calculate).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="应用修改", command=self.apply_changes).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(btn_frame, text="另存为...", command=self.save_as).pack(side=tk.LEFT, padx=(10, 0))

        # === 结果显示区域 ===
        result_frame = ttk.LabelFrame(self, text="计算结果", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True)

        # 文本框 + 滚动条
        text_frame = ttk.Frame(result_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.result_text = tk.Text(text_frame, wrap=tk.NONE, font=("Courier", 12), yscrollcommand=scrollbar.set)
        self.result_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.result_text.yview)

        # 横向滚动条
        h_scrollbar = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.result_text.xview)
        h_scrollbar.pack(fill=tk.X)
        self.result_text.config(xscrollcommand=h_scrollbar.set)

    def on_drop(self, event):
        """处理拖放文件"""
        file_path = event.data
        # 处理路径中可能包含的花括号
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
        # 只取第一个文件
        if ' ' in file_path and not os.path.exists(file_path):
            file_path = file_path.split()[0]
        # 保存完整路径，显示文件名
        self.excel_file_path = file_path
        self.file_display.set(os.path.basename(file_path))

    def browse_file(self):
        """选择文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.excel_file_path = file_path
            self.file_display.set(os.path.basename(file_path))

    def calculate(self):
        file_path = self.excel_file_path
        if not file_path:
            messagebox.showerror("错误", "请先选择Excel文件")
            return

        try:
            rate = float(self.rate_var.get().strip())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的税负率")
            return

        try:
            self.adjuster = TaxAdjuster(file_path)
            self.result = self.adjuster.calculate_adjustment(rate)
            self.display_result()
        except Exception as e:
            messagebox.showerror("错误", f"计算失败: {e}")

    def display_result(self):
        if not self.result:
            return

        current = self.result['current']
        target = self.result['target']
        verify = self.result['verify']

        self.result_text.delete(1.0, tk.END)

        lines = []
        lines.append("=" * 60)
        lines.append(" 调整目标")
        lines.append("=" * 60)
        lines.append(f"  税负率:  {target['rate']*100:.4f}%")
        lines.append(f"  E31:     -10 ~ 10 之间")
        lines.append("")

        lines.append("=" * 60)
        lines.append(" 需要调整的数据")
        lines.append("=" * 60)
        lines.append("")
        lines.append(f"  G25 (成本系数):   {current['G25']}  ->  {target['G25']:.9f}")
        lines.append(f"  E18 (年利润总额): {current['E18']:.2f}  ->  {target['E18']:.2f}")
        lines.append("")

        lines.append("=" * 60)
        lines.append(" 调整前后对比")
        lines.append("=" * 60)
        lines.append(f"  {'项目':<16} {'调整前':>14} {'调整后':>14} {'变化':>12}")
        lines.append(f"  {'-'*56}")
        lines.append(f"  G25 成本系数     {current['G25']:>14.9f} {target['G25']:>14.9f} {(target['G25']-current['G25'])/current['G25']*100:>+11.2f}%")
        lines.append(f"  E18 年利润总额   {current['E18']:>14,.2f} {target['E18']:>14,.2f} {target['E18']-current['E18']:>+12,.2f}")
        lines.append(f"  J12 销售成本     {current['J12']:>14,.2f} {verify['J12']:>14,.2f} {verify['J12']-current['J12']:>+12,.2f}")
        lines.append(f"  B46 当月利润     {current['B46']:>14,.2f} {verify['B46']:>14,.2f} {verify['B46']-current['B46']:>+12,.2f}")
        lines.append(f"  E21 年应纳税额   {current['E21']:>14,.2f} {verify['E21']:>14,.2f} {verify['E21']-current['E21']:>+12,.2f}")
        lines.append(f"  E31 差异         {current['E31']:>14,.2f} {verify['E31']:>14,.2f} {verify['E31']-current['E31']:>+12,.2f}")
        cur_rate = current['E21']/current['E17']*100
        new_rate = verify['rate']*100
        lines.append(f"  税负率           {cur_rate:>13.4f}% {new_rate:>13.4f}% {new_rate-cur_rate:>+11.4f}%")
        lines.append("")

        lines.append("=" * 60)
        lines.append(" 验证结果")
        lines.append("=" * 60)
        rate_ok = abs(verify['rate'] - target['rate']) < 0.00001
        e31_ok = -10 <= verify['E31'] <= 10
        lines.append(f"  税负率: {verify['rate']*100:.4f}%  {'OK' if rate_ok else 'FAIL'}")
        lines.append(f"  E31:    {verify['E31']:.2f}  {'OK' if e31_ok else 'FAIL'}")

        self.result_text.insert(tk.END, "\n".join(lines))

    def apply_changes(self):
        if not self.adjuster or not self.result:
            messagebox.showerror("错误", "请先计算调整方案")
            return

        if messagebox.askyesno("确认", "确定要将修改应用到原文件吗？"):
            try:
                target = self.result['target']
                save_path = self.adjuster.apply_adjustment(target['G25'], target['E18'])
                messagebox.showinfo("成功", f"已保存至:\n{save_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {e}")

    def save_as(self):
        if not self.adjuster or not self.result:
            messagebox.showerror("错误", "请先计算调整方案")
            return

        filename = filedialog.asksaveasfilename(
            title="另存为",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if filename:
            try:
                target = self.result['target']
                save_path = self.adjuster.apply_adjustment(target['G25'], target['E18'], filename)
                messagebox.showinfo("成功", f"已保存至:\n{save_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {e}")
