#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
会计助手 - 整合版
包含：生成凭证、调整税负率
"""

import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import TkinterDnD

# 导入模块
from modules.voucher import VoucherTab
from modules.tax_adjuster import TaxAdjustTab


class AccountingAssistantApp:
    """会计助手主应用"""

    def __init__(self, root):
        self.root = root
        self.root.title("会计助手")
        self.root.geometry("800x700")
        self.root.resizable(True, True)

        self.setup_ui()

    def setup_ui(self):
        """设置主界面"""
        # 创建 Notebook (tabs 容器)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 创建 Tab 1: 生成凭证
        self.voucher_tab = VoucherTab(self.notebook)
        self.notebook.add(self.voucher_tab, text="生成凭证")

        # 创建 Tab 2: 调整税负率
        self.tax_tab = TaxAdjustTab(self.notebook)
        self.notebook.add(self.tax_tab, text="调整税负率")


def main():
    root = TkinterDnD.Tk()
    app = AccountingAssistantApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
