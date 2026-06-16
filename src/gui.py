"""
配箱工具 - tkinter GUI 主界面
低资源占用设计，支持Windows XP+
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import Optional, Callable, Dict, List
from datetime import datetime

# 添加项目路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.engine import PeiXiangEngine, OrderItem, dino_to_hub
from src.excel_sync import ExcelSync


class PeiXiangApp:
    """配箱工具主应用"""

    APP_TITLE = "配箱工具 v1.0"
    WINDOW_SIZE = "1100x720"
    MIN_SIZE = (900, 600)

    # 颜色主题
    COLORS = {
        'bg': '#F5F5F5',
        'primary': '#2196F3',
        'primary_dark': '#1976D2',
        'success': '#4CAF50',
        'warning': '#FF9800',
        'error': '#F44336',
        'text': '#212121',
        'text_secondary': '#757575',
        'border': '#E0E0E0',
        'white': '#FFFFFF',
        'row_alt': '#F8F9FA',
    }

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(self.APP_TITLE)
        self.root.geometry(self.WINDOW_SIZE)
        self.root.minsize(*self.MIN_SIZE)

        # 核心组件
        self.engine = PeiXiangEngine()
        self.excel_sync: Optional[ExcelSync] = None
        self.excel_path: Optional[str] = None
        self.is_loading = False

        # 设置样式
        self._setup_styles()

        # 构建UI
        self._build_ui()

        # 状态栏
        self.status_var = tk.StringVar(value="就绪 - 请打开配箱表文件")
        self._build_statusbar()

        # 绑定事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _setup_styles(self):
        """设置ttk样式"""
        style = ttk.Style()
        style.theme_use('clam')

        style.configure('Title.TLabel', font=('Microsoft YaHei', 11, 'bold'))
        style.configure('Header.TLabel', font=('Microsoft YaHei', 9, 'bold'))
        style.configure('Status.TLabel', font=('Microsoft YaHei', 9), background=self.COLORS['border'])

        style.configure('Primary.TButton', font=('Microsoft YaHei', 9))
        style.configure('Success.TButton', font=('Microsoft YaHei', 9))
        style.configure('Danger.TButton', font=('Microsoft YaHei', 9))

        style.configure('Treeview', font=('Microsoft YaHei', 9), rowheight=26)
        style.configure('Treeview.Heading', font=('Microsoft YaHei', 9, 'bold'))

    def _build_ui(self):
        """构建主界面"""
        # 主容器
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        # 顶部工具栏
        self._build_toolbar(main_frame)

        # 内容区 - 使用Notebook分页
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=4)

        # 5个功能页
        self._build_order_tab()    # 订单导入
        self._build_peixiang_tab() # 配箱计算
        self._build_result_tab()   # 结果导出
        self._build_asn_tab()      # ASN管理
        self._build_settings_tab() # 基础配置

    def _build_toolbar(self, parent):
        """构建顶部工具栏"""
        toolbar = ttk.Frame(parent)
        toolbar.pack(fill=tk.X, pady=(0, 4))

        # 文件操作
        ttk.Button(toolbar, text="📂 打开配箱表", command=self._open_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="💾 保存", command=self._save_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🔄 同步", command=self._sync_data).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)

        # 快捷操作
        ttk.Button(toolbar, text="📋 粘贴订单", command=self._paste_orders).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="⚙️ 一键配箱", command=self._auto_peixiang).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="📊 生成发货通知", command=self._gen_shipment_notice).pack(
            side=tk.LEFT, padx=2)

        # 右侧状态
        self.file_label = ttk.Label(toolbar, text="未打开文件", foreground=self.COLORS['text_secondary'])
        self.file_label.pack(side=tk.RIGHT, padx=8)

    def _build_statusbar(self):
        """构建底部状态栏"""
        statusbar = ttk.Frame(self.root, relief=tk.SUNKEN)
        statusbar.pack(fill=tk.X, side=tk.BOTTOM)

        ttk.Label(statusbar, textvariable=self.status_var, style='Status.TLabel').pack(
            side=tk.LEFT, padx=8, pady=2)

        self.time_var = tk.StringVar(value="")
        ttk.Label(statusbar, textvariable=self.time_var, style='Status.TLabel').pack(
            side=tk.RIGHT, padx=8, pady=2)
        self._update_time()

    def _update_time(self):
        """更新时间显示"""
        self.time_var.set(datetime.now().strftime("%H:%M:%S"))
        self.root.after(1000, self._update_time)

    # ===== 订单导入页 =====
    def _build_order_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=" 订单导入 ")

        # 上部：操作区
        top_frame = ttk.LabelFrame(tab, text="导入方式")
        top_frame.pack(fill=tk.X, padx=8, pady=4)

        ttk.Button(top_frame, text="📋 从剪贴板粘贴 (Ctrl+V)", command=self._paste_orders).pack(
            side=tk.LEFT, padx=8, pady=6)
        ttk.Button(top_frame, text="📄 从CSV文件导入", command=self._import_csv).pack(
            side=tk.LEFT, padx=8, pady=6)
        ttk.Button(top_frame, text="🗑️ 清空订单", command=self._clear_orders).pack(
            side=tk.LEFT, padx=8, pady=6)

        self.order_count_var = tk.StringVar(value="当前订单: 0 条")
        ttk.Label(top_frame, textvariable=self.order_count_var).pack(side=tk.RIGHT, padx=8)

        # 下部：订单列表
        list_frame = ttk.Frame(tab)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        columns = ('po', 'product', 'quantity', 'org', 'brand', 'batch', 'req', 'di')
        self.order_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=20)

        self.order_tree.heading('po', text='客户订单号')
        self.order_tree.heading('product', text='物料名称')
        self.order_tree.heading('quantity', text='数量')
        self.order_tree.heading('org', text='销售组织')
        self.order_tree.heading('brand', text='牌号')
        self.order_tree.heading('batch', text='指定批次')
        self.order_tree.heading('req', text='其他要求')
        self.order_tree.heading('di', text='指定DI')

        self.order_tree.column('po', width=100)
        self.order_tree.column('product', width=250)
        self.order_tree.column('quantity', width=60)
        self.order_tree.column('org', width=80)
        self.order_tree.column('brand', width=80)
        self.order_tree.column('batch', width=80)
        self.order_tree.column('req', width=80)
        self.order_tree.column('di', width=80)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.order_tree.yview)
        self.order_tree.configure(yscrollcommand=scrollbar.set)
        self.order_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # ===== 配箱计算页 =====
    def _build_peixiang_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=" 配箱计算 ")

        # 操作区
        top_frame = ttk.LabelFrame(tab, text="配箱操作")
        top_frame.pack(fill=tk.X, padx=8, pady=4)

        ttk.Button(top_frame, text="🚀 一键配箱", command=self._auto_peixiang).pack(
            side=tk.LEFT, padx=8, pady=6)
        ttk.Button(top_frame, text="🔄 重新配箱", command=self._re_peixiang).pack(
            side=tk.LEFT, padx=8, pady=6)
        ttk.Button(top_frame, text="✅ 写回Excel", command=self._write_back).pack(
            side=tk.LEFT, padx=8, pady=6)

        self.peixiang_status_var = tk.StringVar(value="等待配箱...")
        ttk.Label(top_frame, textvariable=self.peixiang_status_var,
                  foreground=self.COLORS['primary']).pack(side=tk.RIGHT, padx=8)

        # 配箱结果列表
        list_frame = ttk.Frame(tab)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        columns = ('po', 'product', 'quantity', 'org', 'brand_mapped', 'box', 'device', 'di', 'remaining', 'status')
        self.peixiang_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=20)

        self.peixiang_tree.heading('po', text='客户订单号')
        self.peixiang_tree.heading('product', text='物料名称')
        self.peixiang_tree.heading('quantity', text='数量')
        self.peixiang_tree.heading('org', text='销售组织')
        self.peixiang_tree.heading('brand_mapped', text='牌号(映射)')
        self.peixiang_tree.heading('box', text='箱号')
        self.peixiang_tree.heading('device', text='装置')
        self.peixiang_tree.heading('di', text='DI')
        self.peixiang_tree.heading('remaining', text='剩余数量')
        self.peixiang_tree.heading('status', text='状态')

        self.peixiang_tree.column('po', width=100)
        self.peixiang_tree.column('product', width=200)
        self.peixiang_tree.column('quantity', width=50)
        self.peixiang_tree.column('org', width=70)
        self.peixiang_tree.column('brand_mapped', width=140)
        self.peixiang_tree.column('box', width=120)
        self.peixiang_tree.column('device', width=60)
        self.peixiang_tree.column('di', width=100)
        self.peixiang_tree.column('remaining', width=60)
        self.peixiang_tree.column('status', width=70)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.peixiang_tree.yview)
        self.peixiang_tree.configure(yscrollcommand=scrollbar.set)
        self.peixiang_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 标签颜色
        self.peixiang_tree.tag_configure('success', foreground=self.COLORS['success'])
        self.peixiang_tree.tag_configure('error', foreground=self.COLORS['error'])
        self.peixiang_tree.tag_configure('warning', foreground=self.COLORS['warning'])

    # ===== 结果导出页 =====
    def _build_result_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=" 结果导出 ")

        top_frame = ttk.LabelFrame(tab, text="导出操作")
        top_frame.pack(fill=tk.X, padx=8, pady=4)

        ttk.Button(top_frame, text="📊 生成配箱表", command=self._gen_peixiang_table).pack(
            side=tk.LEFT, padx=8, pady=6)
        ttk.Button(top_frame, text="📤 生成发货通知单", command=self._gen_shipment_notice).pack(
            side=tk.LEFT, padx=8, pady=6)
        ttk.Button(top_frame, text="📋 复制到剪贴板", command=self._copy_results).pack(
            side=tk.LEFT, padx=8, pady=6)

        # 结果预览
        self.result_text = scrolledtext.ScrolledText(tab, font=('Consolas', 10), wrap=tk.NONE)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

    # ===== ASN管理页 =====
    def _build_asn_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=" ASN管理 ")

        top_frame = ttk.LabelFrame(tab, text="ASN操作")
        top_frame.pack(fill=tk.X, padx=8, pady=4)

        ttk.Button(top_frame, text="📋 粘贴ASN数据", command=self._paste_asn).pack(
            side=tk.LEFT, padx=8, pady=6)
        ttk.Button(top_frame, text="📥 导入ASN", command=self._import_asn).pack(
            side=tk.LEFT, padx=8, pady=6)

        self.asn_count_var = tk.StringVar(value="ASN记录: 0 条")
        ttk.Label(top_frame, textvariable=self.asn_count_var).pack(side=tk.RIGHT, padx=8)

        # ASN列表
        list_frame = ttk.Frame(tab)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        columns = ('po', 'doc_no', 'status', 'org', 'material', 'quantity', 'container', 'seal')
        self.asn_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=20)

        self.asn_tree.heading('po', text='采购单号')
        self.asn_tree.heading('doc_no', text='单据编号')
        self.asn_tree.heading('status', text='状态')
        self.asn_tree.heading('org', text='销售组织')
        self.asn_tree.heading('material', text='物料名称')
        self.asn_tree.heading('quantity', text='数量')
        self.asn_tree.heading('container', text='集装箱号')
        self.asn_tree.heading('seal', text='铅封号')

        self.asn_tree.column('po', width=100)
        self.asn_tree.column('doc_no', width=110)
        self.asn_tree.column('status', width=60)
        self.asn_tree.column('org', width=80)
        self.asn_tree.column('material', width=200)
        self.asn_tree.column('quantity', width=60)
        self.asn_tree.column('container', width=120)
        self.asn_tree.column('seal', width=100)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.asn_tree.yview)
        self.asn_tree.configure(yscrollcommand=scrollbar.set)
        self.asn_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # ===== 基础配置页 =====
    def _build_settings_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=" 基础配置 ")

        # 文件路径配置
        path_frame = ttk.LabelFrame(tab, text="文件配置")
        path_frame.pack(fill=tk.X, padx=8, pady=4)

        ttk.Label(path_frame, text="配箱表路径:").grid(row=0, column=0, padx=8, pady=4, sticky=tk.W)
        self.path_var = tk.StringVar()
        ttk.Entry(path_frame, textvariable=self.path_var, width=60).grid(row=0, column=1, padx=4, pady=4)
        ttk.Button(path_frame, text="浏览...", command=self._browse_file).grid(row=0, column=2, padx=4, pady=4)

        # 同步配置
        sync_frame = ttk.LabelFrame(tab, text="同步配置")
        sync_frame.pack(fill=tk.X, padx=8, pady=4)

        self.auto_sync_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(sync_frame, text="配箱前自动同步Excel数据", variable=self.auto_sync_var).pack(
            anchor=tk.W, padx=8, pady=2)

        self.auto_backup_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(sync_frame, text="操作前自动备份", variable=self.auto_backup_var).pack(
            anchor=tk.W, padx=8, pady=2)

        self.conflict_alert_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(sync_frame, text="检测到外部修改时提醒", variable=self.conflict_alert_var).pack(
            anchor=tk.W, padx=8, pady=2)

        # 引擎信息
        info_frame = ttk.LabelFrame(tab, text="引擎状态")
        info_frame.pack(fill=tk.X, padx=8, pady=4)

        self.engine_info_var = tk.StringVar(value="引擎未初始化")
        ttk.Label(info_frame, textvariable=self.engine_info_var, wraplength=600).pack(
            anchor=tk.W, padx=8, pady=4)

        ttk.Button(info_frame, text="🔄 刷新引擎状态", command=self._refresh_engine).pack(
            padx=8, pady=4)

        # 牌号映射表
        map_frame = ttk.LabelFrame(tab, text="牌号映射表 (ERP物料 → 物流表物料)")
        map_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        map_list_frame = ttk.Frame(map_frame)
        map_list_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        columns = ('erp_name', 'logistics_name')
        self.map_tree = ttk.Treeview(map_list_frame, columns=columns, show='headings', height=10)
        self.map_tree.heading('erp_name', text='ERP物料名称')
        self.map_tree.heading('logistics_name', text='物流表物料名称')
        self.map_tree.column('erp_name', width=350)
        self.map_tree.column('logistics_name', width=350)

        map_scrollbar = ttk.Scrollbar(map_list_frame, orient=tk.VERTICAL, command=self.map_tree.yview)
        self.map_tree.configure(yscrollcommand=map_scrollbar.set)
        self.map_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        map_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # ===== 业务逻辑方法 =====

    def _open_file(self):
        """打开配箱表文件"""
        filepath = filedialog.askopenfilename(
            title="选择配箱表文件",
            filetypes=[("Excel文件", "*.xlsm *.xlsx"), ("所有文件", "*.*")]
        )
        if not filepath:
            return

        self._load_file(filepath)

    def _load_file(self, filepath: str):
        """加载文件并初始化引擎"""
        self.status_var.set("正在加载文件...")
        self.root.update_idletasks()

        try:
            self.excel_path = filepath
            self.path_var.set(filepath)

            # 备份
            if self.auto_backup_var.get():
                self.excel_sync = ExcelSync(filepath)
                self.excel_sync.open()
                backup_path = self.excel_sync.backup()
                self.status_var.set(f"已备份至: {os.path.basename(backup_path)}")

            # 加载数据到引擎
            self.excel_sync.reload_to_engine(self.engine)

            # 更新UI
            self._refresh_order_list()
            self._refresh_asn_list()
            self._refresh_mapping_list()
            self._refresh_engine()

            self.file_label.config(text=os.path.basename(filepath), foreground=self.COLORS['success'])
            self.status_var.set(f"加载完成 - {self.engine.get_statistics()}")

        except Exception as e:
            messagebox.showerror("加载失败", str(e))
            self.status_var.set(f"加载失败: {e}")

    def _save_file(self):
        """保存文件"""
        if not self.excel_sync:
            messagebox.showwarning("提示", "请先打开配箱表文件")
            return

        try:
            if self.auto_backup_var.get():
                self.excel_sync.backup()
            self.excel_sync.save()
            self.status_var.set("保存成功")
            messagebox.showinfo("成功", "文件已保存")
        except Exception as e:
            messagebox.showerror("保存失败", str(e))

    def _sync_data(self):
        """同步Excel数据"""
        if not self.excel_sync:
            messagebox.showwarning("提示", "请先打开配箱表文件")
            return

        # 检查外部修改
        if self.conflict_alert_var.get() and self.excel_sync.check_external_modify():
            result = messagebox.askyesnocancel(
                "文件已被修改",
                "检测到配箱表已被外部程序修改！\n\n"
                "是 - 重新加载（丢弃当前未保存的修改）\n"
                "否 - 保留当前数据继续\n"
                "取消 - 取消操作"
            )
            if result is None:  # 取消
                return
            elif result:  # 是 - 重新加载
                self.excel_sync.close()
                self._load_file(self.excel_path)
                return

        # 同步
        self.status_var.set("正在同步数据...")
        try:
            self.excel_sync.reload_to_engine(self.engine)
            self._refresh_order_list()
            self._refresh_engine()
            self.status_var.set("同步完成")
        except Exception as e:
            messagebox.showerror("同步失败", str(e))
            self.status_var.set(f"同步失败: {e}")

    def _paste_orders(self):
        """从剪贴板粘贴订单数据"""
        if not self.excel_sync:
            messagebox.showwarning("提示", "请先打开配箱表文件")
            return

        try:
            clipboard = self.root.clipboard_get()
            if not clipboard.strip():
                messagebox.showwarning("提示", "剪贴板为空")
                return

            # 解析Tab分隔的数据
            lines = clipboard.strip().split('\n')
            data = []
            for line in lines:
                row = line.split('\t')
                data.append(row)

            if not data:
                messagebox.showwarning("提示", "未检测到有效数据")
                return

            # 确认导入
            result = messagebox.askyesno(
                "确认导入",
                f"检测到 {len(data)} 行数据，是否导入到配箱公式sheet？\n\n"
                "数据格式应与ERP导出格式一致（Tab分隔）"
            )
            if not result:
                return

            # 导入
            start_row = self.excel_sync.import_orders_from_clipboard(data)
            self.status_var.set(f"已导入 {len(data)} 条订单，起始行: {start_row}")

            # 刷新
            self.excel_sync.reload_to_engine(self.engine)
            self._refresh_order_list()

        except Exception as e:
            messagebox.showerror("导入失败", str(e))

    def _import_csv(self):
        """从CSV文件导入订单"""
        filepath = filedialog.askopenfilename(
            title="选择CSV文件",
            filetypes=[("CSV文件", "*.csv"), ("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if not filepath:
            return

        try:
            import csv
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                data = [row for row in reader]

            if data:
                # 去掉可能的表头行
                first_row = data[0]
                if any(h in str(first_row) for h in ['客户订单号', '物料名称', '销售组织']):
                    data = data[1:]

                result = messagebox.askyesno(
                    "确认导入",
                    f"检测到 {len(data)} 行数据，是否导入？"
                )
                if result:
                    start_row = self.excel_sync.import_orders_from_clipboard(data)
                    self.excel_sync.reload_to_engine(self.engine)
                    self._refresh_order_list()
                    self.status_var.set(f"CSV导入完成: {len(data)} 条")

        except Exception as e:
            messagebox.showerror("CSV导入失败", str(e))

    def _clear_orders(self):
        """清空订单"""
        if not messagebox.askyesno("确认", "确定要清空当前订单数据吗？\n此操作不可撤销！"):
            return
        # 清空配箱公式sheet的订单区域
        if self.excel_sync:
            try:
                ws = self.excel_sync.wb['配箱公式']
                for r in range(2, ws.max_row + 1):
                    if ws.cell(r, 26).value:
                        # 清空Z列开始的订单数据
                        for col in range(1, 90):
                            if col not in [1, 2]:  # 保留A:B映射表
                                ws.cell(r, col).value = None
                self._refresh_order_list()
                self.status_var.set("订单已清空")
            except Exception as e:
                messagebox.showerror("清空失败", str(e))

    def _auto_peixiang(self):
        """一键配箱"""
        if not self.excel_sync:
            messagebox.showwarning("提示", "请先打开配箱表文件")
            return

        # 自动同步
        if self.auto_sync_var.get():
            self._sync_data()

        self.status_var.set("正在执行配箱计算...")
        self.root.update_idletasks()

        try:
            # 读取订单
            orders_data = self.excel_sync.read_orders()
            if not orders_data:
                messagebox.showwarning("提示", "未找到订单数据")
                return

            # 构建OrderItem列表
            orders = []
            for od in orders_data:
                order = OrderItem(
                    row=od['row'],
                    po=od['po'],
                    product_name=od['product_name'],
                    quantity=float(od['quantity']),
                    sales_org=od['sales_org'],
                    brand=od['brand'],
                    spec_batch=od['spec_batch'],
                    other_req=od['other_req'],
                    spec_di=od['spec_di'],
                )
                orders.append(order)

            # 执行配箱
            results = self.engine.compute_all_orders(orders)

            # 更新结果列表
            self._refresh_peixiang_results(results)

            # 统计
            success = sum(1 for r in results if r.box_number and r.box_number != "#N/A")
            failed = sum(1 for r in results if r.box_number == "#N/A")
            self.peixiang_status_var.set(f"配箱完成: 成功 {success}, 未匹配 {failed}")
            self.status_var.set(f"配箱计算完成")

        except Exception as e:
            messagebox.showerror("配箱失败", str(e))
            self.status_var.set(f"配箱失败: {e}")

    def _re_peixiang(self):
        """重新配箱"""
        self._auto_peixiang()

    def _write_back(self):
        """将配箱结果写回Excel"""
        if not self.excel_sync:
            return

        result = messagebox.askyesno(
            "确认写回",
            "将配箱结果写回配箱公式sheet？\n\n"
            "注意：这将覆盖公式列的计算值。"
        )
        if not result:
            return

        try:
            # 获取当前配箱结果
            orders_data = self.excel_sync.read_orders()
            orders = []
            for od in orders_data:
                order = OrderItem(
                    row=od['row'],
                    po=od['po'],
                    product_name=od['product_name'],
                    quantity=float(od['quantity']),
                    sales_org=od['sales_org'],
                    brand=od['brand'],
                    spec_batch=od['spec_batch'],
                    other_req=od['other_req'],
                    spec_di=od['spec_di'],
                )
                orders.append(order)

            results = self.engine.compute_all_orders(orders)
            self.excel_sync.write_peixiang_results(results)

            if self.auto_backup_var.get():
                self.excel_sync.backup()
            self.excel_sync.save()

            self.status_var.set("配箱结果已写回并保存")
            messagebox.showinfo("成功", "配箱结果已写回Excel并保存")

        except Exception as e:
            messagebox.showerror("写回失败", str(e))

    def _gen_peixiang_table(self):
        """生成配箱表"""
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "配箱表生成功能 - 待实现\n")
        self.result_text.insert(tk.END, "将根据配箱结果生成配箱表sheet数据\n")

    def _gen_shipment_notice(self):
        """生成发货通知单"""
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "发货通知单数据生成功能\n")
        self.result_text.insert(tk.END, "=" * 60 + "\n\n")

        if not self.excel_sync:
            self.result_text.insert(tk.END, "请先打开配箱表文件\n")
            return

        try:
            orders_data = self.excel_sync.read_orders()
            orders = []
            for od in orders_data:
                order = OrderItem(
                    row=od['row'],
                    po=od['po'],
                    product_name=od['product_name'],
                    quantity=float(od['quantity']),
                    sales_org=od['sales_org'],
                    brand=od['brand'],
                    spec_batch=od['spec_batch'],
                    other_req=od['other_req'],
                    spec_di=od['spec_di'],
                )
                orders.append(order)

            results = self.engine.compute_all_orders(orders)
            notices = self.excel_sync.generate_shipment_notice(results)

            for notice in notices:
                self.result_text.insert(tk.END, f"PO: {notice['客户采购单号']}\n")
                self.result_text.insert(tk.END, f"  物料: {notice['物料名称']}\n")
                self.result_text.insert(tk.END, f"  数量: {notice['数量']}\n")
                self.result_text.insert(tk.END, f"  箱号: {notice['集装箱号']}\n")
                self.result_text.insert(tk.END, f"  批次: {notice['批次']}\n\n")

        except Exception as e:
            self.result_text.insert(tk.END, f"生成失败: {e}\n")

    def _copy_results(self):
        """复制结果到剪贴板"""
        self.root.clipboard_clear()
        self.root.clipboard_append(self.result_text.get(1.0, tk.END))
        self.status_var.set("已复制到剪贴板")

    def _paste_asn(self):
        """粘贴ASN数据"""
        if not self.excel_sync:
            messagebox.showwarning("提示", "请先打开配箱表文件")
            return

        try:
            clipboard = self.root.clipboard_get()
            lines = clipboard.strip().split('\n')
            data = []
            for line in lines:
                data.append(line.split('\t'))

            if not data:
                return

            result = messagebox.askyesno(
                "确认导入ASN",
                f"检测到 {len(data)} 行ASN数据，是否插入到ASN sheet顶部？"
            )
            if result:
                # 转换为ASN格式
                asn_rows = []
                for row in data:
                    asn_rows.append({
                        'po': row[0] if len(row) > 0 else '',
                        'doc_no': row[1] if len(row) > 1 else '',
                        'status': row[2] if len(row) > 2 else '',
                        'delivery_method': row[3] if len(row) > 3 else '',
                        'sales_org': row[4] if len(row) > 4 else '',
                        'batch': row[5] if len(row) > 5 else '',
                    })

                self.excel_sync.insert_asn_data(asn_rows)
                if self.auto_backup_var.get():
                    self.excel_sync.backup()
                self.excel_sync.save()
                self._refresh_asn_list()
                self.status_var.set(f"ASN数据已插入: {len(data)} 条")

        except Exception as e:
            messagebox.showerror("ASN导入失败", str(e))

    def _import_asn(self):
        """导入ASN"""
        self._paste_asn()

    def _browse_file(self):
        """浏览选择文件"""
        filepath = filedialog.askopenfilename(
            title="选择配箱表文件",
            filetypes=[("Excel文件", "*.xlsm *.xlsx"), ("所有文件", "*.*")]
        )
        if filepath:
            self.path_var.set(filepath)
            self._load_file(filepath)

    def _refresh_engine(self):
        """刷新引擎状态"""
        stats = self.engine.get_statistics()
        info = (
            f"牌号映射: {stats['material_mapping_count']} 条 | "
            f"箱号数据: {stats['box_items_count']} 条 | "
            f"查询索引: {stats['final_key_index_size']} 组 | "
            f"物流数据: {stats['logistics_data_count']} 条 | "
            f"ASN数据: {stats['asn_data_count']} 条"
        )
        self.engine_info_var.set(info)

    def _refresh_order_list(self):
        """刷新订单列表"""
        for item in self.order_tree.get_children():
            self.order_tree.delete(item)

        if not self.excel_sync:
            return

        try:
            orders = self.excel_sync.read_orders()
            for od in orders:
                self.order_tree.insert('', tk.END, values=(
                    od['po'],
                    od['product_name'][:40],
                    od['quantity'],
                    od['sales_org'],
                    od['brand'],
                    od['spec_batch'],
                    od['other_req'],
                    od['spec_di'],
                ))
            self.order_count_var.set(f"当前订单: {len(orders)} 条")
        except Exception:
            pass

    def _refresh_peixiang_results(self, results: List[OrderItem]):
        """刷新配箱结果列表"""
        for item in self.peixiang_tree.get_children():
            self.peixiang_tree.delete(item)

        for r in results:
            status = "✅ 已匹配" if r.box_number and r.box_number != "#N/A" else "❌ 未匹配"
            tag = 'success' if r.box_number and r.box_number != "#N/A" else 'error'

            self.peixiang_tree.insert('', tk.END, values=(
                r.po,
                r.product_name[:30],
                r.quantity,
                r.sales_org,
                r.brand_mapped[:20],
                r.box_number or '',
                r.device,
                r.di,
                r.remaining,
                status,
            ), tags=(tag,))

    def _refresh_asn_list(self):
        """刷新ASN列表"""
        for item in self.asn_tree.get_children():
            self.asn_tree.delete(item)

        if not self.excel_sync:
            return

        try:
            asn_data = self.excel_sync.read_asn_data()
            # 只显示最近100条
            for ad in asn_data[:100]:
                self.asn_tree.insert('', tk.END, values=(
                    ad['po'],
                    ad['doc_no'],
                    ad['status'],
                    ad['sales_org'],
                    ad['material'][:30],
                    ad['quantity'],
                    ad['container_no'],
                    ad['seal_no'],
                ))
            self.asn_count_var.set(f"ASN记录: {len(asn_data)} 条 (显示前100条)")
        except Exception:
            pass

    def _refresh_mapping_list(self):
        """刷新牌号映射列表"""
        for item in self.map_tree.get_children():
            self.map_tree.delete(item)

        for erp_name, logistics_name in self.engine.material_mapping.items():
            self.map_tree.insert('', tk.END, values=(erp_name, logistics_name))

    def _on_close(self):
        """关闭窗口"""
        if self.excel_sync and self.excel_sync.wb:
            try:
                self.excel_sync.close()
            except Exception:
                pass
        self.root.destroy()

    def run(self):
        """启动应用"""
        # 如果有命令行参数，尝试打开文件
        if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
            self.root.after(100, lambda: self._load_file(sys.argv[1]))
        self.root.mainloop()


if __name__ == '__main__':
    app = PeiXiangApp()
    app.run()
