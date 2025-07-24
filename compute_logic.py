#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
compute_logic.py
--------------------------------------------------
核心计算逻辑，供后端调用。
支持：
1. 同时处理多个 Excel 文件 (openpyxl)。
2. 按 "Created Time" 日期字段进行过滤 (闭区间)。
3. 以 Seller SKU 为分组键输出各项指标。
"""

from collections import defaultdict
from datetime import datetime, date
from io import BytesIO, TextIOWrapper
from typing import Dict, List, Iterable, Union
import re
import csv
from zipfile import BadZipFile

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

# 列映射
TARGET_COLUMNS = {
    "order_substatus": ["order substatus"],
    "cancel_type": ["cancelation/return type", "cancellation/return type"],
    "seller_sku": ["seller sku"],
    "shipped_time": ["shipped time"],
    "created_time": ["created time"],
}


def _norm(text):
    return str(text).strip().lower() if text is not None else ""


def _locate_cols(headers: List[str]):
    header_map = {_norm(h): idx for idx, h in enumerate(headers) if h is not None}
    idx_map: Dict[str, int] = {}
    for key, aliases in TARGET_COLUMNS.items():
        for a in aliases:
            if a in header_map:
                idx_map[key] = header_map[a]
                break
        if key not in idx_map:
            raise KeyError(f"列缺失: {aliases[0]}")
    return idx_map


def _iter_rows(file_bytes: Union[str, bytes, BytesIO]):
    """遍历文件行，兼容 .xlsx 与 .csv"""
    if isinstance(file_bytes, str):
        with open(file_bytes, "rb") as f:
            data = BytesIO(f.read())
    elif isinstance(file_bytes, bytes):
        data = BytesIO(file_bytes)
    else:
        data = file_bytes
        data.seek(0)

    try:
        wb = load_workbook(data, data_only=True)
        ws = wb.active
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        headers = list(header_row)
        cols = _locate_cols(headers)
        for row in ws.iter_rows(min_row=3, values_only=True):  # 跳过描述行
            yield (
                row[cols["seller_sku"]],
                row[cols["order_substatus"]],
                row[cols["cancel_type"]],
                row[cols["shipped_time"]],
                row[cols["created_time"]],
            )
        wb.close()
    except (InvalidFileException, BadZipFile):
        data.seek(0)
        wrapper = TextIOWrapper(data, encoding="utf-8-sig")
        reader = csv.reader(wrapper)
        headers = next(reader)
        cols = _locate_cols(headers)
        next(reader, None)  # 跳过描述行
        for row in reader:
            yield (
                row[cols["seller_sku"]] if cols["seller_sku"] < len(row) else None,
                row[cols["order_substatus"]] if cols["order_substatus"] < len(row) else None,
                row[cols["cancel_type"]] if cols["cancel_type"] < len(row) else None,
                row[cols["shipped_time"]] if cols["shipped_time"] < len(row) else None,
                row[cols["created_time"]] if cols["created_time"] < len(row) else None,
            )


def _to_date(val) -> Union[date, None]:
    """尽可能解析单元格中的日期/日期时间，失败返回 None"""
    if val is None:
        return None
    # 如果已经是 datetime/date 类型
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val

    # 如果是 Excel 序列号 (openpyxl 会转为数字类型且未设置 date_only)
    if isinstance(val, (int, float)):
        try:
            from openpyxl.utils.datetime import from_excel
            return from_excel(val).date()
        except Exception:
            pass

    # 字符串解析：支持常见格式
    if isinstance(val, str):
        txt = val.strip()
        # 替换中文分隔符
        txt = txt.replace("年", "-").replace("月", "-").replace("日", "")

        patterns = [
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d",
            "%d/%m/%Y %H:%M:%S",
            "%d/%m/%Y",
            "%m/%d/%Y %H:%M:%S",
            "%m/%d/%Y",
        ]
        for fmt in patterns:
            try:
                return datetime.strptime(txt, fmt).date()
            except ValueError:
                continue

        # 兜底：YYYYMMDD
        m = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", txt)
        if m:
            try:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                pass
    return None


def _date_in_range(d: Union[datetime, date, str, int, float, None], start: date, end: date):
    parsed = _to_date(d)
    if parsed is None:
        return False
    return start <= parsed <= end


def compute_metrics(file_streams: Iterable[BytesIO], start_date: date, end_date: date):
    """核心接口：返回 (Workbook, stats_dict)"""
    stats: Dict[str, Dict[str, int]] = defaultdict(lambda: defaultdict(int))
    total_rows = 0

    for fs in file_streams:
        for seller_sku, sub, cancel, shipped, created in _iter_rows(fs):
            if seller_sku is None:
                continue
            if not _date_in_range(created, start_date, end_date):
                continue

            sku = str(seller_sku)
            s = stats[sku]
            s["total"] += 1
            total_rows += 1

            sub_lower = _norm(sub)
            cancel_lower = _norm(cancel)
            shipped_empty = shipped is None or str(shipped).strip() == ""

            completed_set = {"已完成", "completed"}
            delivered_set = {"已送达", "delivered"}
            canceled_set = {"已取消", "canceled"}
            in_transit_set = {"运输中", "in transit"}

            if sub_lower in completed_set and cancel_lower == "":
                s["completed"] += 1
            elif sub_lower in delivered_set:
                s["delivered"] += 1
            elif "return" in sub_lower or "refund" in sub_lower:
                s["refund"] += 1
            elif sub_lower in canceled_set:
                if shipped_empty:
                    s["cancel_before"] += 1
                else:
                    s["cancel_after"] += 1
            elif sub_lower in in_transit_set:
                s["in_transit"] += 1

    wb = Workbook()
    ws = wb.active
    ws.title = "订单指标"
    headers = [
        "Seller SKU", "订单数", "签收率(%)", "已完成率(%)", "已送达率(%)", "退款率(%)", "发货前取消率(%)", "发货后取消率(%)", "仍在途率(%)",
    ]
    ws.append(headers)

    for sku, m in sorted(stats.items(), key=lambda x: (-x[1]["total"], x[0])):
        total = m["total"]
        if total == 0:
            continue
        completed_rate = m.get("completed", 0) / total * 100
        delivered_rate = m.get("delivered", 0) / total * 100
        refund_rate = m.get("refund", 0) / total * 100
        cancel_before_rate = m.get("cancel_before", 0) / total * 100
        cancel_after_rate = m.get("cancel_after", 0) / total * 100
        in_transit_rate = m.get("in_transit", 0) / total * 100
        sign_rate = completed_rate + delivered_rate + refund_rate

        ws.append([
            sku,
            total,
            round(sign_rate, 2),
            round(completed_rate, 2),
            round(delivered_rate, 2),
            round(refund_rate, 2),
            round(cancel_before_rate, 2),
            round(cancel_after_rate, 2),
            round(in_transit_rate, 2),
        ])

    for idx in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(idx)].width = 14

    return wb, stats 