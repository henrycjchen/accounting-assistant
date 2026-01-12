#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成凭证 Tab
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from openpyxl import Workbook
from tkinterdnd2 import DND_FILES

from .create_outbound import create_outbound
from .create_inbound import create_inbound
from .create_issuing import create_issuing
from .create_receiving import create_receiving


class VoucherTab(ttk.Frame):
    """生成凭证 Tab"""

    def __init__(self, parent):
        super().__init__(parent, padding="10")

        # 文件路径（完整路径）
        self.outbound_invoices_path = ""
        self.calculate_path = ""
        self.inbound_invoices_path = ""

        # 显示用的文件名
        self.outbound_display = tk.StringVar()
        self.calculate_display = tk.StringVar()
        self.inbound_display = tk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        """设置界面"""
        # 标题
        title_label = ttk.Label(
            self,
            text="生成凭证",
            font=("Microsoft YaHei", 16, "bold")
        )
        title_label.pack(pady=(0, 20))

        # 文件选择区域
        files_frame = ttk.LabelFrame(self, text="选择文件", padding="10")
        files_frame.pack(fill=tk.X, pady=(0, 20))

        # 出库发票文件
        self.create_file_row(
            files_frame,
            "出库发票文件：",
            self.outbound_display,
            "outbound_invoices_path"
        )

        # 测算表文件
        self.create_file_row(
            files_frame,
            "测算表：",
            self.calculate_display,
            "calculate_path"
        )

        # 入库发票文件
        self.create_file_row(
            files_frame,
            "入库发票文件：",
            self.inbound_display,
            "inbound_invoices_path"
        )

        # 按钮区域
        button_frame = ttk.Frame(self)
        button_frame.pack(fill=tk.X, pady=(0, 20))

        generate_btn = ttk.Button(
            button_frame,
            text="生成凭证",
            command=self.generate_files,
            width=20
        )
        generate_btn.pack(side=tk.LEFT, padx=(0, 10))

        clear_btn = ttk.Button(
            button_frame,
            text="清空",
            command=self.clear_files,
            width=10
        )
        clear_btn.pack(side=tk.LEFT)

        # 状态区域
        self.status_frame = ttk.LabelFrame(self, text="状态", padding="10")
        self.status_frame.pack(fill=tk.BOTH, expand=True)

        self.status_text = tk.Text(
            self.status_frame,
            height=12,
            state=tk.DISABLED,
            wrap=tk.WORD
        )
        self.status_text.pack(fill=tk.BOTH, expand=True)

    def create_file_row(self, parent, label_text, display_var, path_attr):
        """创建文件选择行"""
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=(0, 10))

        label = ttk.Label(frame, text=label_text, width=15)
        label.pack(side=tk.LEFT)

        entry = ttk.Entry(frame, textvariable=display_var, width=40, state="readonly")
        entry.pack(side=tk.LEFT, padx=(0, 10))

        # 注册拖放支持
        entry.drop_target_register(DND_FILES)
        entry.dnd_bind('<<Drop>>', lambda e: self.on_drop(e, display_var, path_attr))

        btn = ttk.Button(
            frame,
            text="选择",
            command=lambda: self.select_file(display_var, path_attr),
            width=8
        )
        btn.pack(side=tk.LEFT)

    def on_drop(self, event, display_var, path_attr):
        """处理文件拖放"""
        file_path = event.data
        # 处理路径中可能包含的花括号
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
        # 只取第一个文件
        if ' ' in file_path and not os.path.exists(file_path):
            file_path = file_path.split()[0]
        # 保存完整路径，显示文件名
        setattr(self, path_attr, file_path)
        display_var.set(os.path.basename(file_path))
        self.log_status(f"已拖入文件：{os.path.basename(file_path)}")

    def select_file(self, display_var, path_attr):
        """选择文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            setattr(self, path_attr, file_path)
            display_var.set(os.path.basename(file_path))

    def clear_files(self):
        """清空文件选择"""
        self.outbound_invoices_path = ""
        self.calculate_path = ""
        self.inbound_invoices_path = ""
        self.outbound_display.set("")
        self.calculate_display.set("")
        self.inbound_display.set("")
        self.log_status("已清空文件选择")

    def log_status(self, message):
        """记录状态"""
        self.status_text.config(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
        self.update()

    def generate_files(self):
        """生成凭证文件"""
        outbound_path = self.outbound_invoices_path
        calculate_path = self.calculate_path
        inbound_path = self.inbound_invoices_path

        if not outbound_path:
            messagebox.showerror("错误", "请选择出库发票文件")
            return

        try:
            self.log_status("开始生成凭证...")

            # 创建工作簿
            workbook = Workbook()
            if 'Sheet' in workbook.sheetnames:
                del workbook['Sheet']

            # 确定输出路径
            output_dir = os.path.dirname(outbound_path)
            output_filename = f"会计助手-{datetime.now().strftime('%Y%m')}.xlsx"

            # 生成出库凭证
            self.log_status("正在生成出库凭证...")
            outbound = create_outbound(workbook, outbound_path)

            if calculate_path:
                output_dir = os.path.dirname(calculate_path)

                # 生成入库凭证
                self.log_status("正在生成入库凭证...")
                inbound = create_inbound(workbook, calculate_path, outbound)

                # 生成领料单
                self.log_status("正在生成领料单...")
                issuing = create_issuing(workbook, calculate_path, inbound)

                if inbound_path:
                    # 生成收料单
                    self.log_status("正在生成收料单...")
                    create_receiving(workbook, inbound_path, issuing)

            # 保存文件
            output_path = os.path.join(output_dir, output_filename)
            # 如果文件已存在，先删除
            if os.path.exists(output_path):
                os.remove(output_path)
            workbook.save(output_path)

            self.log_status(f"生成完成！文件已保存到：{output_path}")
            messagebox.showinfo("成功", f"凭证文件已生成：\n{output_path}")

        except Exception as e:
            self.log_status(f"生成失败：{str(e)}")
            messagebox.showerror("错误", f"生成失败：{str(e)}")
