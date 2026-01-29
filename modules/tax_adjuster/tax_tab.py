#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调整测算表 Tab
"""

import os
import threading
import wx
import wx.grid

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


class MarginParamsDialog(wx.Dialog):
    """库存毛利率参数设置对话框（扫描加工费对照表）"""

    def __init__(self, parent):
        super().__init__(
            parent,
            title="调整库存毛利率 - 参数设置",
            style=wx.DEFAULT_DIALOG_STYLE
        )

        # 默认值
        self.defaults = {
            'b11_start': 20000,
            'b11_end': 400000,
            'b11_step': 20000,
            'h11_min': TaxAdjuster.H11_MIN,
            'h11_max': TaxAdjuster.H11_MAX,
            'margin_min': TaxAdjuster.MARGIN_MIN,
            'margin_max': TaxAdjuster.MARGIN_MAX,
        }

        self.setup_ui()
        self.Centre()

    def setup_ui(self):
        """设置界面"""
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # 参数输入区域
        grid_sizer = wx.FlexGridSizer(rows=4, cols=4, hgap=10, vgap=10)

        # B11（加工费）范围
        grid_sizer.Add(wx.StaticText(self, label="加工费起始:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.b11_start_ctrl = wx.TextCtrl(self, value=str(self.defaults['b11_start']), size=(80, -1))
        grid_sizer.Add(self.b11_start_ctrl, 0)
        grid_sizer.Add(wx.StaticText(self, label="结束:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_RIGHT)
        self.b11_end_ctrl = wx.TextCtrl(self, value=str(self.defaults['b11_end']), size=(80, -1))
        grid_sizer.Add(self.b11_end_ctrl, 0)

        # B11 步进
        grid_sizer.Add(wx.StaticText(self, label="步进:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.b11_step_ctrl = wx.TextCtrl(self, value=str(self.defaults['b11_step']), size=(80, -1))
        grid_sizer.Add(self.b11_step_ctrl, 0)
        grid_sizer.Add(wx.StaticText(self, label=""), 0)  # 占位
        grid_sizer.Add(wx.StaticText(self, label=""), 0)  # 占位

        # H11 范围
        grid_sizer.Add(wx.StaticText(self, label="H11 范围:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.h11_min_ctrl = wx.TextCtrl(self, value=str(self.defaults['h11_min']), size=(80, -1))
        grid_sizer.Add(self.h11_min_ctrl, 0)
        grid_sizer.Add(wx.StaticText(self, label="~"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_CENTER_HORIZONTAL)
        self.h11_max_ctrl = wx.TextCtrl(self, value=str(self.defaults['h11_max']), size=(80, -1))
        grid_sizer.Add(self.h11_max_ctrl, 0)

        # 毛利率范围
        grid_sizer.Add(wx.StaticText(self, label="毛利率范围:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.margin_min_ctrl = wx.TextCtrl(self, value=str(self.defaults['margin_min']), size=(80, -1))
        grid_sizer.Add(self.margin_min_ctrl, 0)
        grid_sizer.Add(wx.StaticText(self, label="~"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_CENTER_HORIZONTAL)
        self.margin_max_ctrl = wx.TextCtrl(self, value=str(self.defaults['margin_max']), size=(80, -1))
        grid_sizer.Add(self.margin_max_ctrl, 0)

        main_sizer.Add(grid_sizer, 0, wx.ALL | wx.EXPAND, 20)

        # 按钮
        btn_sizer = wx.StdDialogButtonSizer()
        cancel_btn = wx.Button(self, wx.ID_CANCEL, "取消")
        ok_btn = wx.Button(self, wx.ID_OK, "开始扫描")
        ok_btn.SetDefault()
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.AddButton(ok_btn)
        btn_sizer.Realize()

        main_sizer.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        self.SetSizer(main_sizer)
        main_sizer.Fit(self)

    def get_params(self):
        """返回用户输入的参数"""
        try:
            b11_start = int(self.b11_start_ctrl.GetValue())
            b11_end = int(self.b11_end_ctrl.GetValue())
            b11_step = int(self.b11_step_ctrl.GetValue())
            h11_min = float(self.h11_min_ctrl.GetValue())
            h11_max = float(self.h11_max_ctrl.GetValue())
            margin_min = float(self.margin_min_ctrl.GetValue())
            margin_max = float(self.margin_max_ctrl.GetValue())
        except ValueError:
            # 输入无效，返回默认值
            return {
                'b11_start': self.defaults['b11_start'],
                'b11_end': self.defaults['b11_end'],
                'b11_step': self.defaults['b11_step'],
                'h11_target_range': (self.defaults['h11_min'], self.defaults['h11_max']),
                'margin_range': (self.defaults['margin_min'], self.defaults['margin_max']),
            }

        return {
            'b11_start': b11_start,
            'b11_end': b11_end,
            'b11_step': b11_step,
            'h11_target_range': (h11_min, h11_max),
            'margin_range': (margin_min, margin_max),
        }


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

        btn1 = wx.Button(self, label="调整年利润与月毛利")
        btn1.Bind(wx.EVT_BUTTON, self.adjust_combined)
        btn_sizer.Add(btn1, 0, wx.RIGHT, 10)

        btn2 = wx.Button(self, label="调整库存毛利率")
        btn2.Bind(wx.EVT_BUTTON, self.adjust_inventory_margin)
        btn_sizer.Add(btn2, 0)

        main_sizer.Add(btn_sizer, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # === 进度显示区域 ===
        progress_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.progress_gauge = wx.Gauge(self, range=100, size=(300, 20))
        progress_sizer.Add(self.progress_gauge, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)

        self.progress_label = wx.StaticText(self, label="")
        progress_sizer.Add(self.progress_label, 1, wx.ALIGN_CENTER_VERTICAL)

        main_sizer.Add(progress_sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # 初始隐藏进度条
        self.progress_gauge.Hide()
        self.progress_label.Hide()

        # === 结果显示区域 ===
        result_panel = wx.Panel(self)
        result_sizer = wx.BoxSizer(wx.VERTICAL)

        # 建议调整卡片
        self.suggest_box = wx.StaticBox(result_panel, label="建议调整")
        suggest_sizer = wx.StaticBoxSizer(self.suggest_box, wx.VERTICAL)
        self.suggest_grid = self._create_grid(self.suggest_box, rows=2, cols=2)
        suggest_sizer.Add(self.suggest_grid, 0, wx.ALL, 5)
        result_sizer.Add(suggest_sizer, 0, wx.LEFT | wx.BOTTOM, 0)

        # 验证结果卡片
        self.verify_box = wx.StaticBox(result_panel, label="验证结果")
        verify_sizer = wx.StaticBoxSizer(self.verify_box, wx.VERTICAL)
        self.verify_grid = self._create_grid(self.verify_box, rows=5, cols=4)
        verify_sizer.Add(self.verify_grid, 0, wx.ALL, 5)
        result_sizer.Add(verify_sizer, 0, wx.LEFT | wx.BOTTOM, 0)

        # 状态卡片
        self.status_box = wx.StaticBox(result_panel, label="状态")
        status_sizer = wx.StaticBoxSizer(self.status_box, wx.VERTICAL)
        self.status_grid = self._create_grid(self.status_box, rows=1, cols=4)
        status_sizer.Add(self.status_grid, 0, wx.ALL, 5)
        result_sizer.Add(status_sizer, 0, wx.LEFT, 0)

        result_panel.SetSizer(result_sizer)
        main_sizer.Add(result_panel, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # 初始化时隐藏结果区域
        self.suggest_box.Hide()
        self.verify_box.Hide()
        self.status_box.Hide()

        self.SetSizer(main_sizer)

    def _create_grid(self, parent, rows, cols):
        """创建 Grid 控件"""
        grid = wx.grid.Grid(parent)
        grid.CreateGrid(rows, cols)
        grid.EnableEditing(False)
        grid.EnableGridLines(True)
        grid.SetRowLabelSize(0)
        grid.SetColLabelSize(0)
        grid.SetDefaultCellAlignment(wx.ALIGN_LEFT, wx.ALIGN_CENTER)
        grid.SetDefaultRenderer(wx.grid.GridCellStringRenderer())
        # 禁用滚动条
        grid.ShowScrollbars(wx.SHOW_SB_NEVER, wx.SHOW_SB_NEVER)
        return grid

    def _auto_size_grid(self, grid):
        """自动调整 Grid 大小"""
        grid.AutoSizeColumns()
        grid.AutoSizeRows()
        # 计算总高度和宽度
        total_height = sum(grid.GetRowSize(i) for i in range(grid.GetNumberRows()))
        total_width = sum(grid.GetColSize(i) for i in range(grid.GetNumberCols()))
        # 设置精确尺寸，无多余空白
        grid.SetSize((total_width + 2, total_height + 2))
        grid.SetMinSize((total_width + 2, total_height + 2))
        grid.SetMaxSize((total_width + 2, total_height + 2))

    def _set_cell(self, grid, row, col, value, bold=False, color=None, align_right=False):
        """设置单元格值和样式"""
        grid.SetCellValue(row, col, value)
        if bold:
            font = grid.GetCellFont(row, col)
            font.SetWeight(wx.FONTWEIGHT_BOLD)
            grid.SetCellFont(row, col, font)
        if color:
            grid.SetCellTextColour(row, col, color)
        if align_right:
            grid.SetCellAlignment(row, col, wx.ALIGN_RIGHT, wx.ALIGN_CENTER)

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
            self.adjuster = TaxAdjuster(self.excel_file_path, progress_callback=self._on_progress)
            return True
        except Exception as e:
            wx.MessageBox(f"加载文件失败: {e}", "错误", wx.OK | wx.ICON_ERROR)
            return False

    def _on_progress(self, progress, message):
        """进度回调（从工作线程调用）"""
        wx.CallAfter(self._update_progress, progress, message)

    def _update_progress(self, progress, message):
        """更新进度显示（在主线程中）"""
        self.progress_gauge.SetValue(progress)
        self.progress_label.SetLabel(message)

    def _show_progress(self):
        """显示进度条"""
        self.progress_gauge.SetValue(0)
        self.progress_label.SetLabel("准备中...")
        self.progress_gauge.Show()
        self.progress_label.Show()
        self.Layout()

    def _hide_progress(self):
        """隐藏进度条"""
        self.progress_gauge.Hide()
        self.progress_label.Hide()
        self.Layout()

    def _set_buttons_enabled(self, enabled):
        """启用/禁用按钮"""
        for child in self.GetChildren():
            if isinstance(child, wx.Button):
                child.Enable(enabled)

    def adjust_combined(self, event=None):
        """处理"调整年利润与月毛利"按钮点击"""
        if not self._ensure_file_selected():
            return

        if not self._load_adjuster():
            return

        # 显示进度条，禁用按钮
        self._show_progress()
        self._set_buttons_enabled(False)

        def do_calculate():
            try:
                result = self.adjuster.calculate_combined_adjustment()
                wx.CallAfter(self._on_combined_complete, result, None)
            except Exception as e:
                wx.CallAfter(self._on_combined_complete, None, e)

        thread = threading.Thread(target=do_calculate, daemon=True)
        thread.start()

    def _on_combined_complete(self, result, error):
        """处理计算完成（在主线程中）"""
        self._hide_progress()
        self._set_buttons_enabled(True)

        if error:
            wx.MessageBox(f"计算失败: {error}", "错误", wx.OK | wx.ICON_ERROR)
        else:
            self.display_combined_result(result)

    def adjust_inventory_margin(self, event=None):
        """处理"调整库存毛利率"按钮点击 - 扫描加工费对照表"""
        if not self._ensure_file_selected():
            return

        # 弹出参数设置对话框
        dialog = MarginParamsDialog(self)
        if dialog.ShowModal() != wx.ID_OK:
            dialog.Destroy()
            return

        params = dialog.get_params()
        dialog.Destroy()

        if not self._load_adjuster():
            return

        # 显示进度条，禁用按钮
        self._show_progress()
        self._set_buttons_enabled(False)

        def do_calculate():
            try:
                result = self.adjuster.scan_b11_margin_table(
                    b11_start=params['b11_start'],
                    b11_end=params['b11_end'],
                    b11_step=params['b11_step'],
                    h11_target_range=params['h11_target_range'],
                    margin_range=params['margin_range']
                )
                wx.CallAfter(self._on_inventory_margin_complete, result, None)
            except Exception as e:
                wx.CallAfter(self._on_inventory_margin_complete, None, e)

        thread = threading.Thread(target=do_calculate, daemon=True)
        thread.start()

    def _on_inventory_margin_complete(self, result, error):
        """处理计算完成（在主线程中）"""
        self._hide_progress()
        self._set_buttons_enabled(True)

        if error:
            wx.MessageBox(f"计算失败: {error}", "错误", wx.OK | wx.ICON_ERROR)
        else:
            self.display_inventory_margin_result(result)

    def _clear_grid(self, grid):
        """清空 Grid 内容"""
        grid.ClearGrid()

    def display_combined_result(self, result):
        """显示整合调整结果（年利润 + 月毛利）"""
        current = result['current']
        target = result['target']
        verify = result['verify']

        # 显示卡片
        self.suggest_box.Show()
        self.verify_box.Show()
        self.status_box.Show()

        # === 建议调整卡片 ===
        self._clear_grid(self.suggest_grid)
        self._set_cell(self.suggest_grid, 0, 0, "E18 (年利润总额):", bold=True)
        self._set_cell(self.suggest_grid, 0, 1, f"{target['E18']:,.2f}", color=wx.Colour(0, 100, 180), align_right=True)
        self._set_cell(self.suggest_grid, 1, 0, "G25 (成本系数):", bold=True)
        self._set_cell(self.suggest_grid, 1, 1, f"{target['G25']:.9f}", color=wx.Colour(0, 100, 180), align_right=True)
        self._auto_size_grid(self.suggest_grid)

        # === 验证结果卡片 ===
        self._clear_grid(self.verify_grid)
        # 表头
        self._set_cell(self.verify_grid, 0, 0, "项目", bold=True)
        self._set_cell(self.verify_grid, 0, 1, "当前值", bold=True)
        self._set_cell(self.verify_grid, 0, 2, "调整范围", bold=True)
        self._set_cell(self.verify_grid, 0, 3, "调整值", bold=True)

        # E18
        self._set_cell(self.verify_grid, 1, 0, "E18 (年利润总额)")
        self._set_cell(self.verify_grid, 1, 1, f"{current['E18']:,.2f}", align_right=True)
        self._set_cell(self.verify_grid, 1, 2, f"{TaxAdjuster.E18_MIN:,} ~ {TaxAdjuster.E18_MAX:,}", align_right=True)
        self._set_cell(self.verify_grid, 1, 3, f"{target['E18']:,.2f}", align_right=True)

        # G25
        self._set_cell(self.verify_grid, 2, 0, "G25 (成本系数)")
        self._set_cell(self.verify_grid, 2, 1, f"{current['G25']:.6f}", align_right=True)
        self._set_cell(self.verify_grid, 2, 2, f"{TaxAdjuster.G25_MIN:.2f} ~ {TaxAdjuster.G25_MAX:.2f}", align_right=True)
        self._set_cell(self.verify_grid, 2, 3, f"{target['G25']:.6f}", align_right=True)

        # G22
        g22_ok = abs(verify['G22']) < 0.009
        self._set_cell(self.verify_grid, 3, 0, "G22 (税负差异)")
        self._set_cell(self.verify_grid, 3, 1, f"{current['G22']:,.2f}", align_right=True)
        self._set_cell(self.verify_grid, 3, 2, "-0.009 ~ 0.009", align_right=True)
        self._set_cell(self.verify_grid, 3, 3, f"{verify['G22']:.4f}",
                       color=wx.Colour(0, 128, 0) if g22_ok else wx.Colour(200, 0, 0), align_right=True)

        # E31
        e31_ok = abs(verify['E31']) < 0.009
        self._set_cell(self.verify_grid, 4, 0, "E31 (利润差异)")
        self._set_cell(self.verify_grid, 4, 1, f"{current['E31']:,.2f}", align_right=True)
        self._set_cell(self.verify_grid, 4, 2, "-0.009 ~ 0.009", align_right=True)
        self._set_cell(self.verify_grid, 4, 3, f"{verify['E31']:.4f}",
                       color=wx.Colour(0, 128, 0) if e31_ok else wx.Colour(200, 0, 0), align_right=True)
        self._auto_size_grid(self.verify_grid)

        # === 状态卡片 ===
        self._clear_grid(self.status_grid)
        self._set_cell(self.status_grid, 0, 0, "G22:", bold=True)
        self._set_cell(self.status_grid, 0, 1, "OK" if g22_ok else "FAIL",
                       color=wx.Colour(0, 128, 0) if g22_ok else wx.Colour(200, 0, 0))
        self._set_cell(self.status_grid, 0, 2, "E31:", bold=True)
        self._set_cell(self.status_grid, 0, 3, "OK" if e31_ok else "FAIL",
                       color=wx.Colour(0, 128, 0) if e31_ok else wx.Colour(200, 0, 0))
        self._auto_size_grid(self.status_grid)

        self.Layout()

    def display_inventory_margin_result(self, result):
        """显示库存毛利率调整结果（扫描加工费对照表）"""
        if 'error' in result:
            self.suggest_box.Show()
            self.verify_box.Hide()
            self.status_box.Hide()
            self._clear_grid(self.suggest_grid)
            self._set_cell(self.suggest_grid, 0, 0, "错误:", bold=True, color=wx.Colour(200, 0, 0))
            self._set_cell(self.suggest_grid, 0, 1, result['error'], color=wx.Colour(200, 0, 0))
            self._auto_size_grid(self.suggest_grid)
            self.Layout()
            return

        table = result.get('table', [])
        stats = result.get('stats', {})

        if not table:
            self.suggest_box.Show()
            self.verify_box.Hide()
            self.status_box.Hide()
            self._clear_grid(self.suggest_grid)
            self._set_cell(self.suggest_grid, 0, 0, "提示:", bold=True)
            self._set_cell(self.suggest_grid, 0, 1, "未找到可行方案")
            self._auto_size_grid(self.suggest_grid)
            self.Layout()
            return

        # 显示卡片
        self.suggest_box.Show()
        self.verify_box.Show()
        self.status_box.Show()

        # === 建议调整卡片（显示统计信息）===
        self._clear_grid(self.suggest_grid)
        self._set_cell(self.suggest_grid, 0, 0, "扫描结果:", bold=True)
        self._set_cell(self.suggest_grid, 0, 1, f"共 {stats.get('total_rows', 0)} 行")
        self._set_cell(self.suggest_grid, 1, 0, "H11 达标:", bold=True)
        converged = stats.get('converged_count', 0)
        total = stats.get('total_rows', 0)
        self._set_cell(self.suggest_grid, 1, 1, f"{converged}/{total} 行",
                       color=wx.Colour(0, 128, 0) if converged == total else wx.Colour(200, 150, 0))
        self._auto_size_grid(self.suggest_grid)

        # === 验证结果卡片（显示对照表）===
        num_rows = len(table)
        num_cols = 5  # B11, 毛利率, H11, F20, 状态

        # 重新调整 Grid 大小
        current_rows = self.verify_grid.GetNumberRows()
        current_cols = self.verify_grid.GetNumberCols()

        if current_rows < num_rows + 1:
            self.verify_grid.AppendRows(num_rows + 1 - current_rows)
        elif current_rows > num_rows + 1:
            self.verify_grid.DeleteRows(num_rows + 1, current_rows - num_rows - 1)

        if current_cols < num_cols:
            self.verify_grid.AppendCols(num_cols - current_cols)
        elif current_cols > num_cols:
            self.verify_grid.DeleteCols(num_cols, current_cols - num_cols)

        self._clear_grid(self.verify_grid)

        # 表头
        headers = ["加工费(B11)", "毛利率", "H11", "F20", "状态"]
        for col, header in enumerate(headers):
            self._set_cell(self.verify_grid, 0, col, header, bold=True)

        # 填充每行数据
        h11_range = stats.get('h11_range', (-10, 10))
        h11_min, h11_max = h11_range

        for row, item in enumerate(table, start=1):
            b11 = item.get('B11', 0)
            margin = item.get('margin', 0)
            h11 = item.get('H11', 0)
            f20 = item.get('F20', 0)
            converged = item.get('converged', False)

            # B11（加工费）
            self._set_cell(self.verify_grid, row, 0, f"{b11:,}", align_right=True)

            # 毛利率
            self._set_cell(self.verify_grid, row, 1, f"{margin:.4f}", align_right=True)

            # H11
            h11_ok = h11_min <= h11 <= h11_max
            h11_color = wx.Colour(0, 128, 0) if h11_ok else wx.Colour(200, 0, 0)
            self._set_cell(self.verify_grid, row, 2, f"{h11:.2f}", color=h11_color, align_right=True)

            # F20
            self._set_cell(self.verify_grid, row, 3, f"{f20:,.0f}", align_right=True)

            # 状态
            status = "✓" if converged else "×"
            status_color = wx.Colour(0, 128, 0) if converged else wx.Colour(200, 0, 0)
            self._set_cell(self.verify_grid, row, 4, status, color=status_color)

        self._auto_size_grid(self.verify_grid)

        # === 状态卡片 ===
        self._clear_grid(self.status_grid)
        self._set_cell(self.status_grid, 0, 0, "H11范围:", bold=True)
        self._set_cell(self.status_grid, 0, 1, f"{h11_min} ~ {h11_max}")
        margin_range = stats.get('margin_range', (0.70, 0.90))
        self._set_cell(self.status_grid, 0, 2, "毛利率范围:", bold=True)
        self._set_cell(self.status_grid, 0, 3, f"{margin_range[0]:.2f} ~ {margin_range[1]:.2f}")
        self._auto_size_grid(self.status_grid)

        self.Layout()
