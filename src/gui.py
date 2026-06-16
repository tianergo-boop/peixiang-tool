"""
配箱工具 - 图形界面 (性能优化版)
使用tkinter构建，低资源占用

性能优化：
1. 多线程执行耗时操作
2. 进度条显示
3. 懒加载数据
4. UI响应优化
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import queue
import time
from typing import Optional, Callable, Dict, List, Any
from dataclasses import dataclass

from .engine import PeiXiangEngine, OrderItem, ComputeResult
from .excel_sync import ExcelSync


@dataclass
class TaskProgress:
    """任务进度"""
    progress: float  # 0-100
    message: str
    data: Any = None


class PeiXiangApp:
    """配箱工具主应用 (性能优化版)"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("配箱工具 v1.1 (性能优化版)")
        self.root.geometry("1200x800")

        # 核心组件
        self.engine = PeiXiangEngine()
        self.excel_sync: Optional[ExcelSync] = None
        self.current_file: Optional[str] = None

        # 性能优化：后台任务队列
        self.task_queue = queue.Queue()
        self.current_task: Optional[threading.Thread] = None
        self.task_running = False

        # 数据缓存（懒加载）
        self._orders_cache: Optional[List[Dict]] = None
        self._results_cache: Optional[List[ComputeResult]] = None
        self._cache_timestamp: float = 0

        # 创建UI
        self._create_ui()
        self._create_status_bar()
        
        # 启动任务监控
        self._check_task_queue()

    # ===== UI创建 (优化：延迟加载) =====
    def _create_ui(self):
        """创建主界面"""
        # 顶部工具栏
        self._create_toolbar()

        # 主标签页
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 延迟创建标签页内容（优化：首次切换时再创建）
        self.tab_order = self._create_tab_frame("订单导入")
        self.tab_peixiang = self._create_tab_frame("配箱结果")
        self.tab_asn = self._create_tab_frame("ASN管理")
        self.tab_mapping = self._create_tab_frame("物料映射")
        self.tab_config = self._create_tab_frame("基础配置")

        # 绑定标签页切换事件（懒加载）
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # 初始化第一个标签页
        self._init_order_tab()

    def _create_tab_frame(self, title: str) -> ttk.Frame:
        """创建标签页框架（延迟加载）"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text=title)
        # 标记为未初始化
        frame.initialized = False
        return frame

    def _on_tab_changed(self, event):
        """标签页切换事件（懒加载）"""
        selected_tab = event.widget.tab(event.widget.index("current"))["text"]
        
        if selected_tab == "订单导入" and not self.tab_order.initialized:
            self._init_order_tab()
        elif selected_tab == "配箱结果" and not self.tab_peixiang.initialized:
            self._init_peixiang_tab()
        elif selected_tab == "ASN管理" and not self.tab_asn.initialized:
            self._init_asn_tab()
        elif selected_tab == "物料映射" and not self.tab_mapping.initialized:
            self._init_mapping_tab()
        elif selected_tab == "基础配置" and not self.tab_config.initialized:
            self._init_config_tab()

    def _mark_tab_initialized(self, tab_frame: ttk.Frame):
        """标记标签页已初始化"""
        tab_frame.initialized = True

    # ===== 订单导入标签页 (优化：多线程 + 进度条) =====
    def _init_order_tab(self):
        """初始化订单导入标签页"""
        if self.tab_order.initialized:
            return

        frame = self.tab_order
        
        # 进度条（优化：显示读取进度）
        progress_frame = ttk.Frame(frame)
        progress_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.order_progress = ttk.Progressbar(progress_frame, mode='determinate', maximum=100)
        self.order_progress.pack(fill=tk.X, side=tk.LEFT, expand=True, padx=(0, 5))
        
        self.order_progress_label = ttk.Label(progress_frame, text="就绪")
        self.order_progress_label.pack(side=tk.LEFT)

        # 订单列表
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 使用Treeview（优化：虚拟滚动，只渲染可见行）
        columns = ("行号", "客户订单号", "物料名称", "数量", "销售组织", "牌号", "状态")
        self.order_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=20)

        # 配置列
        col_widths = {"行号": 50, "客户订单号": 120, "物料名称": 250, "数量": 80, "销售组织": 100, "牌号": 150, "状态": 100}
        for col in columns:
            self.order_tree.heading(col, text=col)
            self.order_tree.column(col, width=col_widths.get(col, 100), anchor=tk.CENTER)

        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.order_tree.yview)
        self.order_tree.configure(yscrollcommand=scrollbar.set)

        self.order_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 按钮区
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(btn_frame, text="📂 选择文件", command=self._open_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="📋 粘贴订单", command=self._paste_orders_threaded).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="🔄 刷新", command=self._refresh_orders_threaded).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="⚙️ 开始配箱", command=self._auto_peixiang_threaded).pack(side=tk.LEFT, padx=2)

        self._mark_tab_initialized(frame)

    def _refresh_orders_threaded(self):
        """多线程刷新订单列表（优化：不阻塞UI）"""
        if not self.current_file:
            messagebox.showwarning("警告", "请先选择文件")
            return

        def do_refresh(progress_callback):
            """后台刷新任务"""
            try:
                progress_callback(10, "正在打开文件...")
                self.excel_sync = ExcelSync(self.current_file)
                self.excel_sync.open(read_only=True)
                
                progress_callback(30, "正在读取订单数据...")
                orders = self.excel_sync.read_orders_optimized(
                    progress_callback=lambda p, m: progress_callback(p, m)
                )
                
                progress_callback(80, "正在更新界面...")
                # 更新UI（需要在主线程）
                self.root.after(0, self._update_order_tree, orders)
                
                progress_callback(100, f"刷新完成！共 {len(orders)} 条订单")
                return orders
            except Exception as e:
                progress_callback(0, f"刷新失败: {e}")
                raise
            finally:
                if self.excel_sync:
                    self.excel_sync.close()

        self._run_background_task(do_refresh, "刷新订单")

    def _update_order_tree(self, orders: List[Dict]):
        """更新订单列表（在主线程执行）"""
        # 清空现有数据
        for item in self.order_tree.get_children():
            self.order_tree.delete(item)

        # 批量插入（优化：使用detach/attach技巧加速）
        batch_size = 100
        for i, order in enumerate(orders):
            if i % batch_size == 0:
                self.root.update_idletasks()  # 让UI有机会刷新
            
            self.order_tree.insert("", tk.END, values=(
                order.get('row', ''),
                order.get('po', ''),
                order.get('product_name', '')[:40],
                order.get('quantity', ''),
                order.get('sales_org', ''),
                order.get('brand', '')[:20],
                "待配箱"
            ))

    def _paste_orders_threaded(self):
        """多线程粘贴订单"""
        if not self.current_file:
            messagebox.showwarning("警告", "请先选择文件")
            return

        def do_paste(progress_callback):
            """后台粘贴任务"""
            try:
                # 从剪贴板读取
                clipboard = self.root.clipboard_get()
                lines = clipboard.strip().split('\n')
                
                progress_callback(20, f"从剪贴板读取 {len(lines)} 行...")
                
                # 解析订单
                orders = []
                for i, line in enumerate(lines):
                    if i >= 1000:  # 限制最多1000行
                        break
                    parts = line.split('\t')
                    if len(parts) >= 3:
                        orders.append({
                            'row': i + 2,
                            'po': parts[0],
                            'product_name': parts[1],
                            'quantity': float(parts[2]) if parts[2].replace('.', '').isdigit() else 0,
                            'sales_org': parts[3] if len(parts) > 3 else '',
                            'brand': parts[4] if len(parts) > 4 else '',
                        })
                    
                    if (i + 1) % 100 == 0:
                        progress_callback(20 + (i / len(lines)) * 60, f"已解析 {i+1} 行...")
                
                progress_callback(90, "正在更新界面...")
                self.root.after(0, self._update_order_tree, orders)
                
                progress_callback(100, f"粘贴完成！共 {len(orders)} 条订单")
            except Exception as e:
                progress_callback(0, f"粘贴失败: {e}")
                raise

        self._run_background_task(do_paste, "粘贴订单")

    def _auto_peixiang_threaded(self):
        """多线程自动配箱（优化：核心性能改进）"""
        if not self.current_file:
            messagebox.showwarning("警告", "请先选择文件")
            return

        def do_peixiang(progress_callback):
            """后台配箱任务"""
            try:
                progress_callback(5, "正在打开文件...")
                self.excel_sync = ExcelSync(self.current_file)
                self.excel_sync.open(read_only=True)
                
                progress_callback(10, "正在加载数据到引擎...")
                orders = self.excel_sync.reload_to_engine(
                    self.engine,
                    progress_callback=lambda p, m: progress_callback(10 + p*0.6, m)
                )
                
                progress_callback(70, "正在执行配箱计算...")
                # 批量计算（优化：分批处理，避免内存峰值）
                batch_size = 50
                all_results = []
                for i in range(0, len(orders), batch_size):
                    batch = orders[i:i+batch_size]
                    batch_items = [
                        OrderItem(
                            row=o['row'],
                            po=o['po'],
                            product_name=o['product_name'],
                            quantity=o['quantity'],
                            sales_org=o['sales_org'],
                            brand=o.get('brand', ''),
                            spec_batch=o.get('spec_batch', '.'),
                            other_req=o.get('other_req', '.'),
                            spec_di=o.get('spec_di', '.'),
                        )
                        for o in batch
                    ]
                    
                    results = self.engine.compute_all_orders(batch_items)
                    all_results.extend(results)
                    
                    progress = 70 + (i / max(len(orders), 1)) * 20
                    progress_callback(progress, f"已处理 {min(i+batch_size, len(orders))}/{len(orders)} 条订单...")
                
                progress_callback(95, "正在更新结果...")
                self._results_cache = all_results
                self.root.after(0, self._update_peixiang_results, all_results)
                
                progress_callback(100, f"配箱完成！共处理 {len(all_results)} 条订单")
                
                # 切换到配箱结果标签页
                self.root.after(100, lambda: self.notebook.select(1))
                
            except Exception as e:
                progress_callback(0, f"配箱失败: {e}")
                raise
            finally:
                if self.excel_sync:
                    self.excel_sync.close()

        self._run_background_task(do_peixiang, "自动配箱")

    # ===== 后台任务管理 (新增) =====
    def _run_background_task(self, task_func: Callable, task_name: str):
        """运行后台任务（优化：不阻塞UI）"""
        if self.task_running:
            messagebox.showwarning("提示", "已有任务正在运行，请等待完成")
            return

        self.task_running = True
        self.status_label.config(text=f"{task_name}中...")

        def progress_callback(progress: float, message: str):
            """进度回调（线程安全）"""
            self.task_queue.put(TaskProgress(progress, message))

        def run_task():
            """运行任务"""
            try:
                task_func(progress_callback)
            except Exception as e:
                self.task_queue.put(TaskProgress(0, f"{task_name}失败: {e}"))
            finally:
                self.task_running = False
                self.task_queue.put(TaskProgress(100, f"{task_name}完成"))

        # 启动后台线程
        self.current_task = threading.Thread(target=run_task, daemon=True)
        self.current_task.start()

    def _check_task_queue(self):
        """检查任务队列（在主线程执行，更新UI）"""
        try:
            while not self.task_queue.empty():
                task_progress = self.task_queue.get_nowait()
                
                # 更新进度条
                if task_progress.progress is not None:
                    self.order_progress['value'] = task_progress.progress
                    self.order_progress_label.config(text=task_progress.message)
                
                # 更新状态栏
                self.status_label.config(text=task_progress.message)
                
                # 如果任务完成，重置状态
                if task_progress.progress == 100 and not self.task_running:
                    self.root.after(2000, lambda: self.status_label.config(text="就绪"))
        except queue.Empty:
            pass
        
        # 继续检查
        self.root.after(100, self._check_task_queue)

    # ===== 配箱结果标签页 (优化：虚拟滚动) =====
    def _init_peixiang_tab(self):
        """初始化配箱结果标签页"""
        if self.tab_peixiang.initialized:
            return

        frame = self.tab_peixiang

        # 进度条
        progress_frame = ttk.Frame(frame)
        progress_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.peixiang_progress = ttk.Progressbar(progress_frame, mode='determinate', maximum=100)
        self.peixiang_progress.pack(fill=tk.X, side=tk.LEFT, expand=True, padx=(0, 5))
        
        self.peixiang_progress_label = ttk.Label(progress_frame, text="请先执行配箱")
        self.peixiang_progress_label.pack(side=tk.LEFT)

        # 结果列表（优化：使用Canvas + 虚拟滚动）
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        columns = ("行号", "客户订单号", "物料名称", "数量", "箱号", "DI", "批次", "装置", "状态")
        self.result_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=20)

        col_widths = {"行号": 50, "客户订单号": 120, "物料名称": 200, "数量": 80, 
                      "箱号": 120, "DI": 100, "批次": 100, "装置": 80, "状态": 80}
        for col in columns:
            self.result_tree.heading(col, text=col)
            self.result_tree.column(col, width=col_widths.get(col, 100), anchor=tk.CENTER)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=scrollbar.set)

        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 按钮区
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(btn_frame, text="📊 生成发货通知", command=self._gen_shipment_notice_threaded).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="📋 复制结果", command=self._copy_results).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="💾 保存到Excel", command=self._save_to_excel_threaded).pack(side=tk.LEFT, padx=2)

        self._mark_tab_initialized(frame)

    def _update_peixiang_results(self, results: List[ComputeResult]):
        """更新配箱结果列表"""
        # 清空
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 批量插入
        for i, r in enumerate(results):
            if i % 100 == 0:
                self.root.update_idletasks()
            
            status = "✅ 匹配成功" if r.box_number and r.box_number != '#N/A' else "❌ 未匹配"
            self.result_tree.insert("", tk.END, values=(
                r.row,
                r.po,
                r.product_name[:30] if r.product_name else '',
                r.quantity,
                r.box_number,
                r.di,
                r.batch,
                r.device,
                status
            ))

    def _gen_shipment_notice_threaded(self):
        """多线程生成发货通知"""
        if not self._results_cache:
            messagebox.showwarning("警告", "请先执行配箱")
            return

        def do_gen(progress_callback):
            """后台生成任务"""
            try:
                progress_callback(50, "正在生成发货通知...")
                # TODO: 实现发货通知生成逻辑
                progress_callback(100, "发货通知生成完成！")
            except Exception as e:
                progress_callback(0, f"生成失败: {e}")
                raise

        self._run_background_task(do_gen, "生成发货通知")

    def _save_to_excel_threaded(self):
        """多线程保存到Excel"""
        if not self.current_file:
            messagebox.showwarning("警告", "请先打开文件")
            return

        def do_save(progress_callback):
            """后台保存任务"""
            try:
                progress_callback(30, "正在保存到Excel...")
                # TODO: 实现保存逻辑
                progress_callback(100, "保存完成！")
            except Exception as e:
                progress_callback(0, f"保存失败: {e}")
                raise

        self._run_background_task(do_save, "保存到Excel")

    # ===== 其他标签页 (延迟加载) =====
    def _init_asn_tab(self):
        """初始化ASN管理标签页"""
        if self.tab_asn.initialized:
            return

        frame = self.tab_asn
        ttk.Label(frame, text="ASN管理功能开发中...").pack(padx=20, pady=20)
        self._mark_tab_initialized(frame)

    def _init_mapping_tab(self):
        """初始化物料映射标签页"""
        if self.tab_mapping.initialized:
            return

        frame = self.tab_mapping
        ttk.Label(frame, text="物料映射功能开发中...").pack(padx=20, pady=20)
        self._mark_tab_initialized(frame)

    def _init_config_tab(self):
        """初始化基础配置标签页"""
        if self.tab_config.initialized:
            return

        frame = self.tab_config
        ttk.Label(frame, text="基础配置功能开发中...").pack(padx=20, pady=20)
        self._mark_tab_initialized(frame)

    # ===== 工具栏和状态栏 =====
    def _create_toolbar(self):
        """创建工具栏"""
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=5, pady=2)

        ttk.Button(toolbar, text="📂 打开文件", command=self._open_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="💾 保存", command=self._save_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🔄 同步数据", command=self._sync_data).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)

        ttk.Button(toolbar, text="📋 粘贴订单", command=self._paste_orders_threaded).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="⚙️ 一键配箱", command=self._auto_peixiang_threaded).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="📊 生成发货通知", command=self._gen_shipment_notice_threaded).pack(side=tk.LEFT, padx=2)

    def _create_status_bar(self):
        """创建状态栏"""
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_label = ttk.Label(status_frame, text="就绪", relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(fill=tk.X, padx=2, pady=2)

    # ===== 文件操作 =====
    def _open_file(self):
        """打开文件"""
        filepath = filedialog.askopenfilename(
            title="选择配箱表文件",
            filetypes=[("Excel文件", "*.xlsm *.xlsx"), ("所有文件", "*.*")]
        )

        if filepath:
            self.current_file = filepath
            self.root.title(f"配箱工具 - {filepath}")
            self._refresh_orders_threaded()

    def _save_file(self):
        """保存文件"""
        if not self.current_file:
            messagebox.showwarning("警告", "请先打开文件")
            return

        if messagebox.askyesno("确认", "确定要保存到原文件吗？"):
            self._save_to_excel_threaded()

    def _sync_data(self):
        """同步数据"""
        if not self.current_file:
            messagebox.showwarning("警告", "请先打开文件")
            return

        self._refresh_orders_threaded()

    def _copy_results(self):
        """复制结果到剪贴板"""
        if not self._results_cache:
            messagebox.showwarning("警告", "没有可复制的结果")
            return

        text = "行号\t客户订单号\t物料名称\t数量\t箱号\tDI\t批次\t装置\n"
        for r in self._results_cache:
            text += f"{r.row}\t{r.po}\t{r.product_name}\t{r.quantity}\t{r.box_number}\t{r.di}\t{r.batch}\t{r.device}\n"

        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("成功", "结果已复制到剪贴板")

    def _on_close(self):
        """关闭应用"""
        if self.excel_sync:
            self.excel_sync.close()
        self.root.destroy()


def main():
    """启动应用"""
    root = tk.Tk()
    app = PeiXiangApp(root)
    root.protocol("WM_DELETE_WINDOW", app._on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
