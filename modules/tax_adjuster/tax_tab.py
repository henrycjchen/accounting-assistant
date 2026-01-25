#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调整测算表 Tab
"""

import os
import wx

from .adjust_tax import TaxAdjuster


class FileDropTarget(wx.FileDropTarget):
    """文件拖放目标"""

    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def OnDropFiles(self, x, y, filenames):
        if filenames:
            self.callback(filenames[0])
        return True


class TaxAdjustTab(wx.Panel):
    """调整测算表 Tab"""

    def __init__(self, parent):
        super().__init__(parent)

        self.adjuster = None

        # 文件路径（完整路径）
        self.excel_file_path = ""

        self.setup_ui()

    def setup_ui(self):
        """设置界面"""
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # 标题
        title_label = wx.StaticText(self, label="调整测算表")
        title_font = title_label.GetFont()
        title_font.SetPointSize(16)
        title_font.SetWeight(wx.FONTWEIGHT_BOLD)
        title_label.SetFont(title_font)
        main_sizer.Add(title_label, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        # === 文件选择区域 ===
        file_box = wx.StaticBox(self, label="选择文件")
        file_sizer = wx.StaticBoxSizer(file_box, wx.VERTICAL)

        # 文件选择行
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)

        label = wx.StaticText(file_box, label="测算表：", size=(100, -1))
        row_sizer.Add(label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)

        self.file_entry = wx.TextCtrl(file_box, style=wx.TE_READONLY, size=(300, -1))
        self.file_entry.SetDropTarget(FileDropTarget(self.on_drop))
        row_sizer.Add(self.file_entry, 1, wx.EXPAND | wx.RIGHT, 10)

        browse_btn = wx.Button(file_box, label="选择", size=(60, -1))
        browse_btn.Bind(wx.EVT_BUTTON, self.browse_file)
        row_sizer.Add(browse_btn, 0)

        file_sizer.Add(row_sizer, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(file_sizer, 0, wx.EXPAND | wx.ALL, 10)

        # === 操作按钮 ===
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        btn1 = wx.Button(self, label="调整年利润")
        btn1.Bind(wx.EVT_BUTTON, self.adjust_annual_profit)
        btn_sizer.Add(btn1, 0, wx.RIGHT, 10)

        btn2 = wx.Button(self, label="调整月毛利")
        btn2.Bind(wx.EVT_BUTTON, self.adjust_monthly_profit)
        btn_sizer.Add(btn2, 0, wx.RIGHT, 10)

        btn3 = wx.Button(self, label="调整库存毛利率")
        btn3.Bind(wx.EVT_BUTTON, self.adjust_inventory_margin)
        btn_sizer.Add(btn3, 0)

        main_sizer.Add(btn_sizer, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # === 结果显示区域 ===
        result_box = wx.StaticBox(self, label="计算结果")
        result_sizer = wx.StaticBoxSizer(result_box, wx.VERTICAL)

        self.result_text = wx.TextCtrl(
            result_box,
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL | wx.VSCROLL
        )
        # 使用等宽字体
        font = wx.Font(12, wx.FONTFAMILY_TELETYPE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        self.result_text.SetFont(font)
        result_sizer.Add(self.result_text, 1, wx.EXPAND | wx.ALL, 5)

        main_sizer.Add(result_sizer, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        self.SetSizer(main_sizer)

    def on_drop(self, file_path):
        """处理拖放文件"""
        self.excel_file_path = file_path
        self.file_entry.SetValue(os.path.basename(file_path))

    def browse_file(self, event=None):
        """选择文件"""
        with wx.FileDialog(
            self,
            "选择Excel文件",
            wildcard="Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
        ) as dialog:
            if dialog.ShowModal() == wx.ID_OK:
                file_path = dialog.GetPath()
                self.excel_file_path = file_path
                self.file_entry.SetValue(os.path.basename(file_path))

    def _ensure_file_selected(self):
        """确保已选择文件"""
        if not self.excel_file_path:
            wx.MessageBox("请先选择Excel文件", "错误", wx.OK | wx.ICON_ERROR)
            return False
        return True

    def _load_adjuster(self):
        """加载调整器"""
        try:
            self.adjuster = TaxAdjuster(self.excel_file_path)
            return True
        except Exception as e:
            wx.MessageBox(f"加载文件失败: {e}", "错误", wx.OK | wx.ICON_ERROR)
            return False

    def adjust_annual_profit(self, event=None):
        """处理"调整年利润"按钮点击"""
        if not self._ensure_file_selected():
            return

        if not self._load_adjuster():
            return

        try:
            result = self.adjuster.calculate_annual_profit_adjustment()
            self.display_annual_profit_result(result)
        except Exception as e:
            wx.MessageBox(f"计算失败: {e}", "错误", wx.OK | wx.ICON_ERROR)

    def adjust_monthly_profit(self, event=None):
        """处理"调整月毛利"按钮点击"""
        if not self._ensure_file_selected():
            return

        if not self._load_adjuster():
            return

        try:
            result = self.adjuster.calculate_monthly_profit_adjustment()
            self.display_monthly_profit_result(result)
        except Exception as e:
            wx.MessageBox(f"计算失败: {e}", "错误", wx.OK | wx.ICON_ERROR)

    def adjust_inventory_margin(self, event=None):
        """处理"调整库存毛利率"按钮点击"""
        if not self._ensure_file_selected():
            return

        if not self._load_adjuster():
            return

        try:
            result = self.adjuster.calculate_inventory_margin_adjustment()
            self.display_inventory_margin_result(result)
        except Exception as e:
            wx.MessageBox(f"计算失败: {e}", "错误", wx.OK | wx.ICON_ERROR)

    def display_annual_profit_result(self, result):
        """显示年利润调整结果"""
        current = result['current']
        target = result['target']
        verify = result['verify']
        constant = result['constant']

        lines = []
        lines.append("=" * 60)
        lines.append(" 调整年利润 - 目标: G22 = 0.00")
        lines.append("=" * 60)
        lines.append("")

        lines.append("【当前值】")
        lines.append(f"  E17 (年收入):     {current['E17']:,.2f}")
        lines.append(f"  E18 (年利润总额): {current['E18']:,.2f}")
        lines.append(f"  E21 (年应纳税额): {current['E21']:,.2f}")
        lines.append(f"  E22 (E21/2):      {current['E22']:,.2f}")
        lines.append(f"  G22 (目标单元格): {current['G22']:,.2f}")
        lines.append(f"  税负率常量:       {constant:.5f}")
        lines.append("")

        lines.append("【建议调整】")
        lines.append(f"  E18 应改为: {target['E18']:,.2f}")
        lines.append("")

        lines.append("【验证结果】")
        lines.append(f"  调整后 E21: {verify['E21']:,.2f}")
        lines.append(f"  调整后 E22: {verify['E22']:,.2f}")
        lines.append(f"  调整后 G22: {verify['G22']:,.4f}")

        g22_ok = abs(verify['G22']) < 0.009
        lines.append(f"  状态: {'OK' if g22_ok else 'FAIL'}")

        self.result_text.SetValue("\n".join(lines))

    def display_monthly_profit_result(self, result):
        """显示月毛利调整结果"""
        current = result['current']
        target = result['target']
        verify = result['verify']
        prev_profit = result['prev_profit']

        lines = []
        lines.append("=" * 60)
        lines.append(" 调整月毛利 - 目标: E31 = 0.00")
        lines.append("=" * 60)
        lines.append("")

        lines.append("【当前值】")
        lines.append(f"  G25 (成本系数):   {current['G25']:.9f}")
        lines.append(f"  B47 (利润总额):   {current['B47']:,.2f}")
        lines.append(f"  E29 (年利润总额): {current['E29']:,.2f}")
        lines.append(f"  E30 (累计利润):   {current['E30']:,.2f}")
        lines.append(f"  E31 (目标单元格): {current['E31']:,.2f}")
        lines.append(f"  J12 (销售成本):   {current['J12']:,.2f}")
        lines.append(f"  上期累计利润:     {prev_profit:,.2f}")
        lines.append("")

        lines.append("【建议调整】")
        lines.append(f"  G25 应改为: {target['G25']:.9f}")
        lines.append(f"  (目标B47:   {target['B47']:,.2f})")
        lines.append("")

        lines.append("【验证结果】")
        lines.append(f"  调整后 B47: {verify['B47']:,.2f}")
        lines.append(f"  调整后 E30: {verify['E30']:,.2f}")
        lines.append(f"  调整后 E31: {verify['E31']:,.4f}")
        lines.append(f"  调整后 J12: {verify['J12']:,.2f}")

        e31_ok = abs(verify['E31']) < 0.009
        lines.append(f"  状态: {'OK' if e31_ok else 'FAIL'}")

        self.result_text.SetValue("\n".join(lines))

    def display_inventory_margin_result(self, result):
        """显示库存毛利率调整结果"""
        if 'error' in result:
            lines = []
            lines.append("=" * 60)
            lines.append(" 调整库存毛利率 - 错误")
            lines.append("=" * 60)
            lines.append("")
            lines.append(f"  {result['error']}")
            self.result_text.SetValue("\n".join(lines))
            return

        current = result['current']
        target = result['target']
        verify = result['verify']

        lines = []
        lines.append("=" * 60)
        lines.append(" 调整库存毛利率 - 目标: H11 = 0, F20 = 0")
        lines.append("=" * 60)
        lines.append("")

        lines.append("【当前值】")
        lines.append(f"  毛利率 (E14除数): {current['margin']:.4f}")
        lines.append(f"  H11 (目标1):      {current['H11']:,.2f}")
        lines.append(f"  B11 (加工费):     {current['B11']:,.2f}")
        lines.append(f"  F20 (目标2):      {current['F20']:,.2f}")
        lines.append("")

        lines.append("【建议调整】")
        lines.append(f"  毛利率应改为: {target['margin']:.4f}")
        lines.append(f"  B11 应改为:   {target['B11']:,.2f}")
        lines.append("")

        lines.append("【验证结果】")
        lines.append(f"  调整后 H11: {verify['H11']:,.2f}")
        lines.append(f"  调整后 F20: {verify['F20']:,.2f}")

        h11_ok = abs(verify['H11']) < 5.0
        f20_ok = abs(verify['F20']) < 20000.0
        lines.append(f"  H11 状态: {'OK' if h11_ok else 'FAIL'} (容差 ±5.0)")
        lines.append(f"  F20 状态: {'OK' if f20_ok else 'FAIL'} (容差 ±20000.0)")

        self.result_text.SetValue("\n".join(lines))
