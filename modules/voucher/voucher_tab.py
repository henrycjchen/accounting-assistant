#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成凭证 Tab
"""

import os
import wx
from datetime import datetime
from openpyxl import Workbook

from .create_outbound import create_outbound
from .create_inbound import create_inbound
from .create_issuing import create_issuing
from .create_receiving import create_receiving


class FileDropTarget(wx.FileDropTarget):
    """文件拖放目标"""

    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def OnDropFiles(self, x, y, filenames):
        if filenames:
            self.callback(filenames[0])
        return True


class VoucherTab(wx.Panel):
    """生成凭证 Tab"""

    def __init__(self, parent):
        super().__init__(parent)

        # 文件路径（完整路径）
        self.outbound_invoices_path = ""
        self.calculate_path = ""
        self.inbound_invoices_path = ""

        # 显示用的输入框
        self.outbound_entry = None
        self.calculate_entry = None
        self.inbound_entry = None

        self.setup_ui()

    def setup_ui(self):
        """设置界面"""
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # 标题
        title_label = wx.StaticText(self, label="生成凭证")
        title_font = title_label.GetFont()
        title_font.SetPointSize(16)
        title_font.SetWeight(wx.FONTWEIGHT_BOLD)
        title_label.SetFont(title_font)
        main_sizer.Add(title_label, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        # 文件选择区域
        files_box = wx.StaticBox(self, label="选择文件")
        files_sizer = wx.StaticBoxSizer(files_box, wx.VERTICAL)

        # 出库发票文件
        self.outbound_entry = self.create_file_row(
            files_box,
            files_sizer,
            "出库发票文件：",
            self.on_outbound_drop,
            lambda: self.select_file("outbound_invoices_path", self.outbound_entry)
        )

        # 测算表文件
        self.calculate_entry = self.create_file_row(
            files_box,
            files_sizer,
            "测算表：",
            self.on_calculate_drop,
            lambda: self.select_file("calculate_path", self.calculate_entry)
        )

        # 入库发票文件
        self.inbound_entry = self.create_file_row(
            files_box,
            files_sizer,
            "入库发票文件：",
            self.on_inbound_drop,
            lambda: self.select_file("inbound_invoices_path", self.inbound_entry)
        )

        main_sizer.Add(files_sizer, 0, wx.EXPAND | wx.ALL, 10)

        # 按钮区域
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)

        generate_btn = wx.Button(self, label="生成凭证", size=(120, -1))
        generate_btn.Bind(wx.EVT_BUTTON, self.generate_files)
        button_sizer.Add(generate_btn, 0, wx.RIGHT, 10)

        clear_btn = wx.Button(self, label="清空", size=(80, -1))
        clear_btn.Bind(wx.EVT_BUTTON, self.clear_files)
        button_sizer.Add(clear_btn, 0)

        main_sizer.Add(button_sizer, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # 状态区域
        status_box = wx.StaticBox(self, label="状态")
        status_sizer = wx.StaticBoxSizer(status_box, wx.VERTICAL)

        self.status_text = wx.TextCtrl(
            status_box,
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL
        )
        status_sizer.Add(self.status_text, 1, wx.EXPAND | wx.ALL, 5)

        main_sizer.Add(status_sizer, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        self.SetSizer(main_sizer)

    def create_file_row(self, parent, sizer, label_text, drop_callback, select_callback):
        """创建文件选择行"""
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)

        label = wx.StaticText(parent, label=label_text, size=(100, -1))
        row_sizer.Add(label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)

        entry = wx.TextCtrl(parent, style=wx.TE_READONLY, size=(300, -1))
        entry.SetDropTarget(FileDropTarget(drop_callback))
        row_sizer.Add(entry, 1, wx.EXPAND | wx.RIGHT, 10)

        btn = wx.Button(parent, label="选择", size=(60, -1))
        btn.Bind(wx.EVT_BUTTON, lambda e: select_callback())
        row_sizer.Add(btn, 0)

        sizer.Add(row_sizer, 0, wx.EXPAND | wx.ALL, 5)

        return entry

    def on_outbound_drop(self, file_path):
        """处理出库发票文件拖放"""
        self.outbound_invoices_path = file_path
        self.outbound_entry.SetValue(os.path.basename(file_path))
        self.log_status(f"已拖入文件：{os.path.basename(file_path)}")

    def on_calculate_drop(self, file_path):
        """处理测算表文件拖放"""
        self.calculate_path = file_path
        self.calculate_entry.SetValue(os.path.basename(file_path))
        self.log_status(f"已拖入文件：{os.path.basename(file_path)}")

    def on_inbound_drop(self, file_path):
        """处理入库发票文件拖放"""
        self.inbound_invoices_path = file_path
        self.inbound_entry.SetValue(os.path.basename(file_path))
        self.log_status(f"已拖入文件：{os.path.basename(file_path)}")

    def select_file(self, path_attr, entry):
        """选择文件"""
        with wx.FileDialog(
            self,
            "选择Excel文件",
            wildcard="Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
        ) as dialog:
            if dialog.ShowModal() == wx.ID_OK:
                file_path = dialog.GetPath()
                setattr(self, path_attr, file_path)
                entry.SetValue(os.path.basename(file_path))

    def clear_files(self, event=None):
        """清空文件选择"""
        self.outbound_invoices_path = ""
        self.calculate_path = ""
        self.inbound_invoices_path = ""
        self.outbound_entry.SetValue("")
        self.calculate_entry.SetValue("")
        self.inbound_entry.SetValue("")
        self.log_status("已清空文件选择")

    def log_status(self, message):
        """记录状态"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.AppendText(f"[{timestamp}] {message}\n")

    def generate_files(self, event=None):
        """生成凭证文件"""
        outbound_path = self.outbound_invoices_path
        calculate_path = self.calculate_path
        inbound_path = self.inbound_invoices_path

        if not outbound_path:
            wx.MessageBox("请选择出库发票文件", "错误", wx.OK | wx.ICON_ERROR)
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
            wx.MessageBox(f"凭证文件已生成：\n{output_path}", "成功", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            self.log_status(f"生成失败：{str(e)}")
            wx.MessageBox(f"生成失败：{str(e)}", "错误", wx.OK | wx.ICON_ERROR)
