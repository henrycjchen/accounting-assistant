#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成凭证 Tab
"""

import os
import wx
import wx.grid as gridlib
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

        # 结果表格区域
        result_box = wx.StaticBox(self, label="生成结果")
        result_sizer = wx.StaticBoxSizer(result_box, wx.VERTICAL)

        # 创建 Grid
        self.result_grid = gridlib.Grid(result_box)
        self.result_grid.CreateGrid(4, 3)
        self.result_grid.EnableEditing(False)
        self.result_grid.SetRowLabelSize(0)
        self.result_grid.SetDefaultCellAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)

        # 设置列标题
        self.result_grid.SetColLabelValue(0, "单据类型")
        self.result_grid.SetColLabelValue(1, "数量")
        self.result_grid.SetColLabelValue(2, "状态")

        # 设置列宽
        self.result_grid.SetColSize(0, 120)
        self.result_grid.SetColSize(1, 80)
        self.result_grid.SetColSize(2, 100)

        # 设置行高
        for i in range(4):
            self.result_grid.SetRowSize(i, 28)

        # 初始化数据
        self.voucher_types = ["出库凭证", "入库凭证", "领料单", "收料单"]
        for i, vtype in enumerate(self.voucher_types):
            self.result_grid.SetCellValue(i, 0, vtype)
            self.result_grid.SetCellValue(i, 1, "--")
            self.result_grid.SetCellValue(i, 2, "等待")
            self.result_grid.SetCellTextColour(i, 2, wx.Colour(150, 150, 150))

        result_sizer.Add(self.result_grid, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)

        # 状态信息
        self.status_label = wx.StaticText(result_box, label="等待生成...")
        self.status_label.SetForegroundColour(wx.Colour(128, 128, 128))
        result_sizer.Add(self.status_label, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.BOTTOM, 10)

        main_sizer.Add(result_sizer, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

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
        self.set_status(f"已选择出库发票文件：{os.path.basename(file_path)}")

    def on_calculate_drop(self, file_path):
        """处理测算表文件拖放"""
        self.calculate_path = file_path
        self.calculate_entry.SetValue(os.path.basename(file_path))
        self.set_status(f"已选择测算表：{os.path.basename(file_path)}")

    def on_inbound_drop(self, file_path):
        """处理入库发票文件拖放"""
        self.inbound_invoices_path = file_path
        self.inbound_entry.SetValue(os.path.basename(file_path))
        self.set_status(f"已选择入库发票文件：{os.path.basename(file_path)}")

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
        self.reset_grid()
        self.set_status("已清空文件选择")

    def set_status(self, message, is_error=False):
        """设置状态信息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_label.SetLabel(f"[{timestamp}] {message}")
        if is_error:
            self.status_label.SetForegroundColour(wx.Colour(204, 0, 0))
        else:
            self.status_label.SetForegroundColour(wx.Colour(128, 128, 128))
        self.status_label.Refresh()

    def reset_grid(self):
        """重置表格"""
        for i in range(4):
            self.result_grid.SetCellValue(i, 1, "--")
            self.result_grid.SetCellValue(i, 2, "等待")
            self.result_grid.SetCellTextColour(i, 2, wx.Colour(150, 150, 150))
        self.result_grid.ForceRefresh()
        self.set_status("等待生成...")

    def update_grid_row(self, row, count=None, status="等待"):
        """更新表格行"""
        if count is not None:
            self.result_grid.SetCellValue(row, 1, str(count))
        elif status == "处理中":
            self.result_grid.SetCellValue(row, 1, "...")

        self.result_grid.SetCellValue(row, 2, status)

        # 设置状态颜色
        if status == "等待":
            self.result_grid.SetCellTextColour(row, 2, wx.Colour(150, 150, 150))
        elif status == "处理中":
            self.result_grid.SetCellTextColour(row, 2, wx.Colour(66, 133, 244))
        elif status == "完成":
            self.result_grid.SetCellTextColour(row, 2, wx.Colour(52, 168, 83))
        elif status == "错误":
            self.result_grid.SetCellTextColour(row, 2, wx.Colour(234, 67, 53))

        self.result_grid.ForceRefresh()

    def generate_files(self, event=None):
        """生成凭证文件"""
        outbound_path = self.outbound_invoices_path
        calculate_path = self.calculate_path
        inbound_path = self.inbound_invoices_path

        if not outbound_path:
            wx.MessageBox("请选择出库发票文件", "错误", wx.OK | wx.ICON_ERROR)
            return

        # 统计数量
        counts = [0, 0, 0, 0]
        current_row = 0

        try:
            # 重置表格状态
            self.reset_grid()
            self.set_status("开始生成凭证...")
            wx.GetApp().Yield()

            # 创建工作簿
            workbook = Workbook()
            if 'Sheet' in workbook.sheetnames:
                del workbook['Sheet']

            # 确定输出路径
            output_dir = os.path.dirname(outbound_path)
            output_filename = f"会计助手-{datetime.now().strftime('%Y%m')}.xlsx"

            # 生成出库凭证
            current_row = 0
            self.update_grid_row(0, status="处理中")
            self.set_status("正在生成出库凭证...")
            wx.GetApp().Yield()

            outbound = create_outbound(workbook, outbound_path)
            counts[0] = len(outbound) if outbound else 0
            self.update_grid_row(0, counts[0], "完成")

            if calculate_path:
                output_dir = os.path.dirname(calculate_path)

                # 生成入库凭证
                current_row = 1
                self.update_grid_row(1, status="处理中")
                self.set_status("正在生成入库凭证...")
                wx.GetApp().Yield()

                inbound = create_inbound(workbook, calculate_path, outbound)
                counts[1] = len(inbound) if inbound else 0
                self.update_grid_row(1, counts[1], "完成")

                # 生成领料单
                current_row = 2
                self.update_grid_row(2, status="处理中")
                self.set_status("正在生成领料单...")
                wx.GetApp().Yield()

                issuing = create_issuing(workbook, calculate_path, inbound)
                counts[2] = len(issuing) if issuing else 0
                self.update_grid_row(2, counts[2], "完成")

                if inbound_path:
                    # 生成收料单
                    current_row = 3
                    self.update_grid_row(3, status="处理中")
                    self.set_status("正在生成收料单...")
                    wx.GetApp().Yield()

                    receiving = create_receiving(workbook, inbound_path, issuing)
                    counts[3] = len(receiving) if receiving else 0
                    self.update_grid_row(3, counts[3], "完成")

            # 保存文件
            output_path = os.path.join(output_dir, output_filename)
            if os.path.exists(output_path):
                os.remove(output_path)
            workbook.save(output_path)

            # 计算总数
            total = sum(counts)
            self.set_status(f"生成完成！共 {total} 张单据，已保存到 {os.path.basename(output_path)}")

            wx.MessageBox(f"凭证文件已生成：\n{output_path}", "成功", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            # 将当前处理中的行设为错误状态
            self.update_grid_row(current_row, status="错误")
            self.set_status(f"生成失败：{str(e)}", is_error=True)
            wx.MessageBox(f"生成失败：{str(e)}", "错误", wx.OK | wx.ICON_ERROR)
