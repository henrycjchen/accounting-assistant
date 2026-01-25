#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
会计助手 - 整合版
包含：生成凭证、调整测算表
"""

import wx

# 导入模块
from modules.voucher import VoucherTab
from modules.tax_adjuster import TaxAdjustTab


class AccountingAssistantApp(wx.Frame):
    """会计助手主应用"""

    def __init__(self):
        super().__init__(None, title="会计助手", size=(800, 700))
        self.SetMinSize((600, 500))

        self.setup_ui()
        self.Centre()

    def setup_ui(self):
        """设置主界面"""
        # 创建主面板
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # 创建 Notebook (tabs 容器)
        self.notebook = wx.Notebook(panel)

        # 创建 Tab 1: 生成凭证
        self.voucher_tab = VoucherTab(self.notebook)
        self.notebook.AddPage(self.voucher_tab, "生成凭证")

        # 创建 Tab 2: 调整测算表
        self.tax_tab = TaxAdjustTab(self.notebook)
        self.notebook.AddPage(self.tax_tab, "调整测算表")

        main_sizer.Add(self.notebook, 1, wx.EXPAND | wx.ALL, 5)
        panel.SetSizer(main_sizer)


def main():
    app = wx.App()
    frame = AccountingAssistantApp()
    frame.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()
