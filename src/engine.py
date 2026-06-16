"""
配箱核心引擎 - 箱号分配算法
还原Excel配箱公式sheet的所有VLOOKUP/COUNTIF逻辑，使用Python字典索引实现O(1)查询
"""

from typing import Optional, Dict, List, Any, Tuple
from dataclasses import dataclass, field
import re
from collections import defaultdict


# ===== DINO前缀 → 销售组织 映射 (还原SWITCH公式) =====
DINO_TO_HUB = {
    "SH": "上海Hub",
    "HP": "黄埔Hub",
    "XG": "新港Hub",
    "QD": "青岛Hub",
    "NB": "宁波Hub",
    "ST": "汕头Hub",
    "XM": "厦门Hub",
    "NS": "南沙Hub",
    "WF": "潍坊Hub",
    "NJ": "南京Hub",
    "TC": "太仓Hub",
    "SQ": "宿迁Hub",
    "CQ": "重庆Hub",
    "WH": "武汉Hub",
    "TZ": "台州Hub",
    "HF": "合肥Hub",
    "FZ": "福州Hub",
}


def dino_to_hub(dino: str) -> str:
    """DINO编号前两位转Hub名称"""
    if not dino or len(dino) < 2:
        return ""
    return DINO_TO_HUB.get(dino[:2], "")


@dataclass
class BoxItem:
    """汇总表中的箱号库存项"""
    row: int
    dino: str = ""           # 运单号 (M列)
    waybill: str = ""        # 运单号 (N列)
    box_number: str = ""     # 箱号 (O列)
    weight: float = 0        # 重量 (P列)
    batch: str = ""          # 批次 (Q列)
    ship_company: str = ""   # 船公司 (R列)
    device: str = ""         # 装置 (S列)
    product_desc: str = ""   # 产品描述 (H列)
    atd: Optional[Any] = None           # 实际离港日期 (I列)
    etd: Optional[Any] = None           # 计划离港日期 (K列)
    transit_port: str = ""   # 中转港 (L列)
    # 右侧区域
    ae_product: str = ""     # AE: 产品描述
    af_etd: Optional[Any] = None       # AF: 计划离港日期
    ag_box: str = ""         # AG: 箱号
    ak_dino: str = ""        # AK: DINO
    al_weight: float = 0     # AL: 重量
    # 计算字段
    hub: str = ""            # B列: 销售组织
    lookup_key: str = ""     # F列: 重量&组织&产品描述
    final_key: str = ""      # G列: lookup_key + 序号


@dataclass
class OrderItem:
    """订单行项"""
    row: int
    po: str = ""              # 客户订单号 (Z列)
    product_name: str = ""    # 物料名称 (AD列) - ERP名称
    quantity: float = 0       # 数量 (AI列)
    sales_org: str = ""       # 销售组织 (AC列)
    brand: str = ""           # 牌号 (AR列) - 原始牌号
    spec_batch: str = ""      # 配箱指定批次/过渡料 (U列)
    other_req: str = ""       # 其他配箱/物流要求 (V列)
    spec_di: str = ""         # 配箱指定DI/箱号/产线 (W列)
    incoterm: str = ""        # SO Incoterm (X列)
    carrier: str = ""         # LYB陆运承运商 (AE列)
    # 计算字段
    brand_mapped: str = ""    # H列: 映射后的牌号 (VLOOKUP AD→B)
    seq1: str = ""            # C列: 序列1
    seq2: int = 0             # D列: 序列2 (COUNTIF)
    query1: str = ""          # E列: 查询1
    seq3: int = 0             # F列: 序列3 (COUNTIF)
    box_number: str = ""      # J列: 箱号
    di: str = ""              # S列: DI
    batch: str = ""           # M列: 装置相关
    device: str = ""          # M列: 装置
    weight_gross: float = 0   # O列: 毛重
    remaining: int = 0        # P列: 剩余数量
    weight_net: float = 0     # Q列: 净重
    pending: int = 0          # R列: 待离数量
    etd: Optional[Any] = None # O列: 离港时间
    coa: str = ""             # N列: COA


class PeiXiangEngine:
    """配箱核心引擎"""

    def __init__(self):
        # 物料名称映射表 (erp物料 → 物流表物料)
        self.material_mapping: Dict[str, str] = {}
        # 汇总表数据
        self.box_items: List[BoxItem] = []
        # 索引字典 (用于O(1)查询)
        self._index_by_final_key: Dict[str, List[BoxItem]] = defaultdict(list)
        self._index_by_box_number: Dict[str, BoxItem] = {}
        self._index_by_lookup_key: Dict[str, List[BoxItem]] = defaultdict(list)
        self._index_by_pending_key: Dict[str, List[BoxItem]] = defaultdict(list)
        self._index_by_waybill: Dict[str, BoxItem] = {}
        # 物流跟踪数据 (运单号 → 行数据)
        self.logistics_data: Dict[str, Dict] = {}
        # ASN数据
        self.asn_data: List[Dict] = []

    # ===== 牌号映射 =====
    def load_material_mapping(self, mapping: Dict[str, str]):
        """加载erp物料名称 → 物流表物料名称 映射"""
        self.material_mapping = mapping

    def map_brand(self, product_name: str) -> str:
        """VLOOKUP(AD2, A:B, 2, 0) - 物料名称转牌号"""
        return self.material_mapping.get(product_name, "")

    # ===== 汇总表索引构建 =====
    def load_summary_data(self, rows: List[Dict]):
        """
        加载汇总表数据并构建索引
        rows: [{row, hub, product_desc, atd, etd, transit_port, dino, waybill,
                box_number, weight, batch, ship_company, device, ...}]
        """
        self.box_items = []
        self._index_by_final_key = defaultdict(list)
        self._index_by_box_number = {}
        self._index_by_lookup_key = defaultdict(list)
        self._index_by_pending_key = defaultdict(list)
        self._index_by_waybill = {}

        # 第一步：创建BoxItem并计算基础字段
        for r in rows:
            item = BoxItem(
                row=r.get('row', 0),
                dino=r.get('dino', ''),
                waybill=r.get('waybill', ''),
                box_number=r.get('box_number', ''),
                weight=float(r.get('weight', 0) or 0),
                batch=r.get('batch', ''),
                ship_company=r.get('ship_company', ''),
                device=r.get('device', ''),
                product_desc=r.get('product_desc', ''),
                atd=r.get('atd'),
                etd=r.get('etd'),
                transit_port=r.get('transit_port', ''),
                ae_product=r.get('ae_product', ''),
                ak_dino=r.get('ak_dino', ''),
                al_weight=float(r.get('al_weight', 0) or 0),
            )
            self.box_items.append(item)

        # 第二步：计算Hub (B列 = SWITCH(LEFT(DINO,2),...))
        for item in self.box_items:
            item.hub = dino_to_hub(item.dino)
            # 右侧Hub
            right_hub = dino_to_hub(item.ak_dino) if item.ak_dino else ""

        # 第三步：计算lookup_key (F列 = 重量&Hub&产品描述)
        for item in self.box_items:
            if item.product_desc and item.weight > 0:
                item.lookup_key = f"{int(item.weight)}{item.hub}{item.product_desc}"

        # 第四步：计算final_key (G列 = lookup_key & 序号)
        # 需要按行顺序处理，COUNTIF逻辑
        key_count: Dict[str, int] = defaultdict(int)
        for item in self.box_items:
            if item.lookup_key:
                key_count[item.lookup_key] += 1
                item.final_key = f"{item.lookup_key}{key_count[item.lookup_key]}"

        # 第五步：构建索引
        for item in self.box_items:
            if item.final_key:
                self._index_by_final_key[item.final_key].append(item)
            if item.box_number:
                self._index_by_box_number[item.box_number] = item
            if item.lookup_key:
                self._index_by_lookup_key[item.lookup_key].append(item)
            # 待离索引: 重量&Hub&产品描述 (不含序号)
            pending_key = f"{int(item.weight)}{item.hub}{item.product_desc}"
            if pending_key:
                self._index_by_pending_key[pending_key].append(item)
            if item.waybill:
                self._index_by_waybill[item.waybill] = item

    # ===== 配箱核心算法 =====
    def compute_order(self, order: OrderItem, row_idx: int, prev_orders: List[OrderItem]) -> OrderItem:
        """
        对单个订单行执行配箱计算
        还原 配箱公式 sheet 的全部公式逻辑
        row_idx: 当前行号(从1开始)
        prev_orders: 之前已处理的订单行(用于COUNTIF计算)
        """
        # === H列: 牌号 = VLOOKUP(AD2, A:B, 2, 0) ===
        order.brand_mapped = self.map_brand(order.product_name)

        # === C列: 序列1 = IF(AND(IF(H_prev==H_curr,1,0), IF(AC_prev==AC_curr,1,0)),"",0) ===
        if row_idx > 1 and prev_orders:
            prev = prev_orders[-1]
            if order.brand_mapped == prev.brand_mapped and order.sales_org == prev.sales_org:
                order.seq1 = ""
            else:
                order.seq1 = "0"
        else:
            order.seq1 = "0"

        # === D列: 序列2 = COUNTIF($C$1:C_curr, 0) ===
        count_0 = sum(1 for o in prev_orders if o.seq1 == "0") + (1 if order.seq1 == "0" else 0)
        order.seq2 = count_0

        # === E列: 查询1 = IF(U&V&W="...", AI&AC&AD, "") ===
        spec_combined = f"{order.spec_batch}{order.other_req}{order.spec_di}"
        if spec_combined == "..." or (order.spec_batch == "." and order.other_req == "." and order.spec_di == "."):
            order.query1 = f"{order.quantity}{order.sales_org}{order.product_name}"
        else:
            order.query1 = ""

        # === F列: 序列3 = COUNTIF($E$1:E_curr, E_curr) ===
        if order.query1:
            count_e = sum(1 for o in prev_orders if o.query1 == order.query1) + 1
            order.seq3 = count_e
        else:
            order.seq3 = 0

        # === J列: 箱号 (核心!) ===
        # =IF($V2&U2&W2="...", VLOOKUP($AI2&$AC2&$H2&$F2, 汇总!G:O, 9, 0), U2&$V2&"/")
        if order.spec_batch == "." and order.other_req == "." and order.spec_di == ".":
            # 自动匹配箱号
            vlookup_key = f"{int(order.quantity)}{order.sales_org}{order.brand_mapped}{order.seq3}"
            matched_items = self._index_by_final_key.get(vlookup_key, [])
            if matched_items:
                # 取第一个匹配项
                first_match = matched_items[0]
                order.box_number = first_match.box_number
            else:
                order.box_number = "#N/A"
        else:
            # 手动指定: 指定批次 & 指定要求 & "/"
            order.box_number = f"{order.spec_batch}{order.other_req}/"

        # === M列: 装置 = VLOOKUP(J2, 汇总!O:S, 5, 0) ===
        if order.box_number and order.box_number != "#N/A":
            box_item = self._index_by_box_number.get(order.box_number)
            if box_item:
                order.device = box_item.device
            else:
                order.device = ""

        # === N列: COA = VLOOKUP(J2, 汇总!O:S, 3, 0) → 实际取运单号 ===
        if order.box_number and order.box_number != "#N/A":
            box_item = self._index_by_box_number.get(order.box_number)
            if box_item:
                order.coa = box_item.waybill
            else:
                order.coa = ""

        # === O列: 离港时间 = VLOOKUP($AI2&$AC2&$H2&$F2, 汇总!G:K, 5, 0) ===
        if order.spec_batch == "." and order.other_req == "." and order.spec_di == ".":
            vlookup_key = f"{int(order.quantity)}{order.sales_org}{order.brand_mapped}{order.seq3}"
            matched_items = self._index_by_final_key.get(vlookup_key, [])
            if matched_items:
                order.etd = matched_items[0].etd

        # === P列: 剩余数量 = COUNTIF(汇总!F:F, $AI2&$AC2&$H2) ===
        pending_lookup = f"{int(order.quantity)}{order.sales_org}{order.brand_mapped}"
        order.remaining = len(self._index_by_lookup_key.get(pending_lookup, []))

        # === Q列: 净重 = VLOOKUP($AI2&$AC2&$H2&$F2, 汇总!AD:AK, 3, 0) ===
        # 暂留 - 需要右侧索引
        # AD列 = AC&AB (待离数量键)

        # === R列: 待离数量 = COUNTIF(汇总!AC:AC, $AI2&$AC2&$H2) ===
        pending_key = f"{order.quantity}{order.sales_org}{order.brand_mapped}"
        order.pending = len(self._index_by_pending_key.get(pending_key, []))

        # === S列: DI = VLOOKUP(J2, IF({0,1},汇总!M:M,汇总!O:O), 2, 0) ===
        # 即用箱号查O列对应M列的DINO
        if order.box_number and order.box_number != "#N/A":
            box_item = self._index_by_box_number.get(order.box_number)
            if box_item:
                order.di = box_item.dino
            else:
                order.di = ""

        return order

    def compute_all_orders(self, orders: List[OrderItem]) -> List[OrderItem]:
        """
        批量执行配箱计算
        orders: 待配箱的订单列表
        返回: 配箱结果
        """
        processed = []
        for idx, order in enumerate(orders):
            result = self.compute_order(order, idx + 1, processed)
            processed.append(result)
        return processed

    def get_box_details(self, box_number: str) -> Optional[BoxItem]:
        """根据箱号获取详细信息"""
        return self._index_by_box_number.get(box_number)

    def search_available_boxes(self, quantity: float, sales_org: str, product_desc: str) -> List[BoxItem]:
        """搜索可用箱号"""
        key = f"{int(quantity)}{sales_org}{product_desc}"
        return self._index_by_lookup_key.get(key, [])

    def get_statistics(self) -> Dict[str, Any]:
        """获取引擎统计信息"""
        return {
            "material_mapping_count": len(self.material_mapping),
            "box_items_count": len(self.box_items),
            "final_key_index_size": len(self._index_by_final_key),
            "box_number_index_size": len(self._index_by_box_number),
            "lookup_key_index_size": len(self._index_by_lookup_key),
            "logistics_data_count": len(self.logistics_data),
            "asn_data_count": len(self.asn_data),
        }
