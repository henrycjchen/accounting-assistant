#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调整税负率 Tab
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
    """调整税负率 Tab"""

    def __init__(self, parent):
        super().__init__(parent)

        self.adjuster = None
        self.result = None

        # 文件路径（完整路径）
        self.excel_file_path = ""

        self.setup_ui()

    def setup_ui(self):
        """设置界面"""
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # 标题
        title_label = wx.StaticText(self, label="调整税负率")
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

        # === 参数设置区域 ===
        param_box = wx.StaticBox(self, label="参数设置")
        param_sizer = wx.StaticBoxSizer(param_box, wx.HORIZONTAL)

        param_label = wx.StaticText(param_box, label="目标税负率:")
        param_sizer.Add(param_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.rate_entry = wx.TextCtrl(param_box, value="0.00414", size=(100, -1))
        param_sizer.Add(self.rate_entry, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        hint_label = wx.StaticText(param_box, label="(例: 0.00414 = 0.414%)")
        param_sizer.Add(hint_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        main_sizer.Add(param_sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # === 操作按钮 ===
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        calc_btn = wx.Button(self, label="计算调整方案")
        calc_btn.Bind(wx.EVT_BUTTON, self.calculate)
        btn_sizer.Add(calc_btn, 0, wx.RIGHT, 10)

        apply_btn = wx.Button(self, label="应用修改")
        apply_btn.Bind(wx.EVT_BUTTON, self.apply_changes)
        btn_sizer.Add(apply_btn, 0, wx.RIGHT, 10)

        save_btn = wx.Button(self, label="另存为...")
        save_btn.Bind(wx.EVT_BUTTON, self.save_as)
        btn_sizer.Add(save_btn, 0)

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

    def calculate(self, event=None):
        file_path = self.excel_file_path
        if not file_path:
            wx.MessageBox("请先选择Excel文件", "错误", wx.OK | wx.ICON_ERROR)
            return

        try:
            rate = float(self.rate_entry.GetValue().strip())
        except ValueError:
            wx.MessageBox("请输入有效的税负率", "错误", wx.OK | wx.ICON_ERROR)
            return

        try:
            self.adjuster = TaxAdjuster(file_path)
            self.result = self.adjuster.calculate_adjustment(rate)
            self.display_result()
        except Exception as e:
            wx.MessageBox(f"计算失败: {e}", "错误", wx.OK | wx.ICON_ERROR)

    def display_result(self):
        if not self.result:
            return

        current = self.result['current']
        target = self.result['target']
        verify = self.result['verify']

        self.result_text.SetValue("")

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

        self.result_text.SetValue("\n".join(lines))

    def apply_changes(self, event=None):
        if not self.adjuster or not self.result:
            wx.MessageBox("请先计算调整方案", "错误", wx.OK | wx.ICON_ERROR)
            return

        dialog = wx.MessageDialog(
            self,
            "确定要将修改应用到原文件吗？",
            "确认",
            wx.YES_NO | wx.ICON_QUESTION
        )
        if dialog.ShowModal() == wx.ID_YES:
            try:
                target = self.result['target']
                save_path = self.adjuster.apply_adjustment(target['G25'], target['E18'])
                wx.MessageBox(f"已保存至:\n{save_path}", "成功", wx.OK | wx.ICON_INFORMATION)
            except Exception as e:
                wx.MessageBox(f"保存失败: {e}", "错误", wx.OK | wx.ICON_ERROR)

    def save_as(self, event=None):
        if not self.adjuster or not self.result:
            wx.MessageBox("请先计算调整方案", "错误", wx.OK | wx.ICON_ERROR)
            return

        with wx.FileDialog(
            self,
            "另存为",
            wildcard="Excel文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*",
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
        ) as dialog:
            if dialog.ShowModal() == wx.ID_OK:
                filename = dialog.GetPath()
                try:
                    target = self.result['target']
                    save_path = self.adjuster.apply_adjustment(target['G25'], target['E18'], filename)
                    wx.MessageBox(f"已保存至:\n{save_path}", "成功", wx.OK | wx.ICON_INFORMATION)
                except Exception as e:
                    wx.MessageBox(f"保存失败: {e}", "错误", wx.OK | wx.ICON_ERROR)
