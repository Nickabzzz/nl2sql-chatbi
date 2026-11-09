# -*- coding: utf-8 -*-
"""生成评测集 v2：题干写清口径 + 备选解读实测。

v1 的 9 个错误里有 4 个是"题目本身口径不清"造成的，所以这一版：
1. 每道题的题干都写明 状态范围 / 时间范围 / 分母 / 是否含运费；
2. 对有歧义的题，额外算出"另一种解读"的答案，并给出差距——用数据证明歧义真实存在；
3. 输出 Excel（三个 sheet）：题目与答案 / 口径与备选解读 / 测试记录。

用法：python build_eval_v2.py
"""
from __future__ import annotations

import os

import json
from decimal import Decimal
from pathlib import Path

import pymysql
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

CONN = dict(host="127.0.0.1", port=13307, user="root", password=os.environ.get("DB_PASSWORD", "changeme"),
            charset="utf8mb4", database="olist")
OUT_JSONL = Path(r".\nl2sql_cases_v2.jsonl")
OUT_XLSX = Path(r".\NL2SQL评测集_v2.xlsx")

OCT = ("o.order_purchase_timestamp>='2026-10-01' AND o.order_purchase_timestamp<'2026-11-01'")
SEP = ("o.order_purchase_timestamp>='2026-09-01' AND o.order_purchase_timestamp<'2026-10-01'")
AUG_OCT = "o.order_purchase_timestamp>='2026-08-01'"

# (编号, 分类, 修订后的问题, 标准SQL, 易错点, [(备选解读, SQL）, ...])
CASES = [
    ("Q01", "基础聚合", "上个月（2026年10月）的销售额是多少？只统计已交付订单，金额为商品金额之和，不含运费",
     f"SELECT ROUND(SUM(oi.price),2) FROM orders o JOIN order_items oi ON o.order_id=oi.order_id "
     f"WHERE o.order_status='delivered' AND {OCT}", "不筛状态会虚高；含运费会偏高",
     [("包含运费", f"SELECT ROUND(SUM(oi.price+oi.freight_value),2) FROM orders o JOIN order_items oi ON o.order_id=oi.order_id WHERE o.order_status='delivered' AND {OCT}")]),

    ("Q02", "基础聚合", "上个月（2026年10月）已交付的订单有多少笔？",
     f"SELECT COUNT(DISTINCT o.order_id) FROM orders o WHERE o.order_status='delivered' AND {OCT}",
     "把全部状态也算进去会多约 161 笔",
     [("包含全部状态", f"SELECT COUNT(DISTINCT o.order_id) FROM orders o WHERE {OCT}")]),

    ("Q03", "派生指标", "上个月已交付订单的平均客单价是多少？销售额÷订单数，两者同一口径、不含运费",
     f"SELECT ROUND(SUM(oi.price)/COUNT(DISTINCT o.order_id),2) FROM orders o JOIN order_items oi ON o.order_id=oi.order_id "
     f"WHERE o.order_status='delivered' AND {OCT}", "分子分母口径不一致就错",
     [("分子分母都用全部状态", f"SELECT ROUND(SUM(oi.price)/COUNT(DISTINCT o.order_id),2) FROM orders o JOIN order_items oi ON o.order_id=oi.order_id WHERE {OCT}")]),

    ("Q04", "时间趋势", "最近三个月（2026年8/9/10月）每个月的已交付销售额分别是多少？不含运费",
     f"SELECT DATE_FORMAT(o.order_purchase_timestamp,'%Y-%m') m, ROUND(SUM(oi.price),2) gmv FROM orders o "
     f"JOIN order_items oi ON o.order_id=oi.order_id WHERE o.order_status='delivered' AND {AUG_OCT} GROUP BY m ORDER BY m",
     "漏掉状态过滤，三个月全部偏高",
     [("包含全部状态", f"SELECT DATE_FORMAT(o.order_purchase_timestamp,'%Y-%m') m, ROUND(SUM(oi.price),2) gmv FROM orders o JOIN order_items oi ON o.order_id=oi.order_id WHERE {AUG_OCT} GROUP BY m ORDER BY m")]),

    ("Q05", "维度分析", "近三个月已交付订单中，销售额最高的 5 个品类是什么？品类用英文名",
     f"SELECT COALESCE(t.product_category_name_english,p.product_category_name) cat, ROUND(SUM(oi.price),2) gmv "
     f"FROM order_items oi JOIN orders o ON oi.order_id=o.order_id JOIN products p ON oi.product_id=p.product_id "
     f"LEFT JOIN category_translation t ON p.product_category_name=t.product_category_name "
     f"WHERE o.order_status='delivered' AND {AUG_OCT} GROUP BY cat ORDER BY gmv DESC LIMIT 5",
     "漏 join 翻译表会得到葡语品类名", []),

    ("Q06", "维度分析", "在全部订单（包含取消、已发货等所有状态）中，订单数量最多的 5 个州是哪些？",
     "SELECT c.customer_state st, COUNT(DISTINCT o.order_id) n FROM orders o JOIN customers c ON o.customer_id=c.customer_id "
     "GROUP BY st ORDER BY n DESC LIMIT 5", "这题不加状态过滤，加了反而错",
     [("仅已交付", "SELECT c.customer_state st, COUNT(DISTINCT o.order_id) n FROM orders o JOIN customers c ON o.customer_id=c.customer_id WHERE o.order_status='delivered' GROUP BY st ORDER BY n DESC LIMIT 5")]),

    ("Q07", "评价", "2026 年的平均评分是多少？保留两位小数",
     "SELECT ROUND(AVG(review_score),2) FROM order_reviews WHERE review_creation_date>='2026-01-01'", "时间条件写错会变", []),

    ("Q08", "评价", "差评（1 到 2 分）占全部评价的比例是多少？百分比保留两位小数",
     "SELECT ROUND(100*SUM(review_score<=2)/COUNT(*),2) FROM order_reviews", "分母是全部评价，不是已交付订单", []),

    ("Q09", "物流", "在已签收的订单中，从下单到签收平均需要多少天？保留两位小数",
     "SELECT ROUND(AVG(DATEDIFF(order_delivered_customer_date,order_purchase_timestamp)),2) FROM orders "
     "WHERE order_delivered_customer_date IS NOT NULL", "未签收订单要排除，否则偏小", []),

    ("Q10", "物流", "已签收订单中，延迟送达（签收晚于预计送达）的订单占比是多少？百分比保留两位小数",
     "SELECT ROUND(100*SUM(order_delivered_customer_date>order_estimated_delivery_date)/COUNT(*),2) FROM orders "
     "WHERE order_delivered_customer_date IS NOT NULL", "分母只能用已签收订单", []),

    ("Q11", "支付", "2026 年 10 月，各种支付方式各自有多少笔订单？",
     f"SELECT p.payment_type, COUNT(DISTINCT p.order_id) n FROM order_payments p JOIN orders o ON p.order_id=o.order_id "
     f"WHERE {OCT} GROUP BY p.payment_type ORDER BY n DESC", "不限时间范围会把全部历史算进来",
     [("全部历史", "SELECT payment_type, COUNT(DISTINCT order_id) n FROM order_payments GROUP BY payment_type ORDER BY n DESC")]),

    ("Q12", "支付", "按订单数计算，使用分期（分期数大于1）的订单占比是多少？百分比保留两位小数",
     "SELECT ROUND(100*COUNT(DISTINCT CASE WHEN payment_installments>1 THEN order_id END)/COUNT(DISTINCT order_id),2) FROM order_payments",
     "分母用支付记录会得到另一个数",
     [("按支付记录计算", "SELECT ROUND(100*SUM(payment_installments>1)/COUNT(*),2) FROM order_payments")]),

    ("Q13", "卖家", "在全部订单中，销售额最高的 5 个卖家是谁？",
     "SELECT oi.seller_id, ROUND(SUM(oi.price),2) gmv FROM order_items oi GROUP BY oi.seller_id ORDER BY gmv DESC LIMIT 5",
     "这题不限状态；加了过滤会变",
     [("仅已交付", "SELECT oi.seller_id, ROUND(SUM(oi.price),2) gmv FROM order_items oi JOIN orders o ON oi.order_id=o.order_id WHERE o.order_status='delivered' GROUP BY oi.seller_id ORDER BY gmv DESC LIMIT 5")]),

    ("Q14", "客户", "下过两次及以上订单的客户有多少人？按客户真实唯一编号统计",
     "SELECT COUNT(*) FROM (SELECT c.customer_unique_id FROM orders o JOIN customers c ON o.customer_id=c.customer_id "
     "GROUP BY c.customer_unique_id HAVING COUNT(*)>1) x", "用 customer_id 会因为一单一号而全部失误", []),

    ("Q15", "客户", "上个月有多少客户是第一次下单？按客户真实唯一编号统计",
     f"SELECT COUNT(*) FROM (SELECT c.customer_unique_id, MIN(o.order_purchase_timestamp) t FROM orders o "
     f"JOIN customers c ON o.customer_id=c.customer_id GROUP BY c.customer_unique_id "
     f"HAVING t>='2026-10-01' AND t<'2026-11-01') x", "先把首单时间算出来再筛", []),

    ("Q16", "状态", "取消订单占总订单的比例是多少？百分比保留两位小数",
     "SELECT ROUND(100*SUM(order_status='canceled')/COUNT(*),2) FROM orders", "分母是全部订单", []),

    ("Q17", "派生指标", "上个月已交付订单的平均每单运费是多少？保留两位小数",
     f"SELECT ROUND(SUM(oi.freight_value)/COUNT(DISTINCT o.order_id),2) FROM orders o JOIN order_items oi ON o.order_id=oi.order_id "
     f"WHERE o.order_status='delivered' AND {OCT}", "不筛状态会得到另一个数",
     [("包含全部状态", f"SELECT ROUND(SUM(oi.freight_value)/COUNT(DISTINCT o.order_id),2) FROM orders o JOIN order_items oi ON o.order_id=oi.order_id WHERE {OCT}")]),

    ("Q18", "趋势对比", "十月与九月相比，已交付销售额的环比增长率是多少？百分比保留两位小数",
     f"SELECT ROUND(100*(SUM(CASE WHEN o.order_purchase_timestamp>='2026-10-01' THEN oi.price END)"
     f"-SUM(CASE WHEN {SEP} THEN oi.price END))/SUM(CASE WHEN {SEP} THEN oi.price END),2) "
     f"FROM orders o JOIN order_items oi ON o.order_id=oi.order_id WHERE o.order_status='delivered' "
     f"AND o.order_purchase_timestamp>='2026-09-01' AND o.order_purchase_timestamp<'2026-11-01'",
     "基期写错会得到完全不同的数",
     [("包含全部状态", f"SELECT ROUND(100*(SUM(CASE WHEN o.order_purchase_timestamp>='2026-10-01' THEN oi.price END)"
                       f"-SUM(CASE WHEN {SEP} THEN oi.price END))/SUM(CASE WHEN {SEP} THEN oi.price END),2) "
                       f"FROM orders o JOIN order_items oi ON o.order_id=oi.order_id "
                       f"WHERE o.order_purchase_timestamp>='2026-09-01' AND o.order_purchase_timestamp<'2026-11-01'")]),

    ("T01", "陷阱·重复计数", "全部订单（含取消）一共卖出了多少件商品？按商品明细行计数",
     "SELECT COUNT(*) FROM order_items", "和「订单数」是两回事；加状态过滤会变",
     [("仅已交付", "SELECT COUNT(*) FROM order_items oi JOIN orders o ON oi.order_id=o.order_id WHERE o.order_status='delivered'")]),

    ("T02", "陷阱·重复计数", "全部订单（含取消）的总支付金额是多少？",
     "SELECT ROUND(SUM(payment_value),2) FROM order_payments", "先 JOIN 订单表再 SUM 会翻倍",
     [("仅已交付", "SELECT ROUND(SUM(p.payment_value),2) FROM order_payments p JOIN orders o ON p.order_id=o.order_id WHERE o.order_status='delivered'")]),

    ("T03", "陷阱·空值", "有多少个商品的品类是未知的？",
     "SELECT COUNT(*) FROM products WHERE product_category_name IS NULL", "用 ='' 会得到 0", []),

    ("T04", "陷阱·状态过滤", "十月包含取消订单在内，销售额是多少？用于对比「只算已交付」差多少",
     f"SELECT ROUND(SUM(oi.price),2) FROM orders o JOIN order_items oi ON o.order_id=oi.order_id WHERE {OCT}",
     "这题故意不筛状态；减去 838576.64 就是虚高部分",
     [("仅已交付", f"SELECT ROUND(SUM(oi.price),2) FROM orders o JOIN order_items oi ON o.order_id=oi.order_id WHERE o.order_status='delivered' AND {OCT}")]),
]

HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
TRAP_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def scalar(value):
    if isinstance(value, Decimal):
        return float(value)
    return value


def run(cur, sql):
    cur.execute(sql)
    data = cur.fetchall()
    if len(data) == 1 and len(data[0]) == 1:
        return scalar(data[0][0])
    return [[str(x) for x in row] for row in data]


def style_header(ws, headers, widths):
    ws.append(headers)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        cell = ws.cell(row=1, column=i)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"


def main() -> int:
    conn = pymysql.connect(**CONN)
    cur = conn.cursor()

    rows = []
    for cid, cat, q, sql, note, alts in CASES:
        expected = run(cur, sql)
        alt_rows = []
        for label, alt_sql in alts:
            alt_rows.append({"label": label, "value": run(cur, alt_sql)})
        rows.append({"id": cid, "category": cat, "question": q, "sql": sql,
                     "expected": expected, "note": note, "alts": alt_rows})
        shown = str(expected)[:44]
        print(f"{cid} {cat:<14} {shown:<46} 备选 {len(alt_rows)}")

    OUT_JSONL.write_text("\n".join(json.dumps(r, ensure_ascii=False) for r in rows), encoding="utf-8")

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "题目与标准答案"
    style_header(ws1, ["编号", "分类", "问题（口径已验证）", "标准答案", "标准 SQL", "易错点"],
                 [7, 16, 46, 30, 62, 26])
    for r in rows:
        exp = r["expected"]
        exp = "\n".join(" | ".join(x) for x in exp) if isinstance(exp, list) else exp
        ws1.append([r["id"], r["category"], r["question"], exp, r["sql"], r["note"]])

    ws2 = wb.create_sheet("口径对照")
    style_header(ws2, ["编号", "问题", "标准答案", "另一种解读", "那种解读的答案", "差距说明"],
                 [7, 40, 24, 20, 24, 30])
    for r in rows:
        exp = r["expected"]
        exp = " | ".join(" ".join(x[1:]) if len(x) > 1 else x[0] for x in exp) if isinstance(exp, list) else exp
        if not r["alts"]:
            ws2.append([r["id"], r["question"], exp, "—", "—", "无歧义"])
            continue
        for alt in r["alts"]:
            v = alt["value"]
            v = " | ".join(" ".join(x[1:]) if len(x) > 1 else x[0] for x in v) if isinstance(v, list) else v
            ws2.append([r["id"], r["question"], exp, alt["label"], v, ""])

    ws3 = wb.create_sheet("我的测试记录")
    style_header(ws3, ["编号", "问题", "标准答案", "模型给的答案", "是否正确", "错误类型", "备注"],
                 [7, 44, 26, 26, 12, 16, 26])
    for r in rows:
        exp = r["expected"]
        exp = " | ".join(" ".join(x[1:]) if len(x) > 1 else x[0] for x in exp) if isinstance(exp, list) else exp
        ws3.append([r["id"], r["question"], exp, "", "", "", ""])

    for ws in (ws1, ws2, ws3):
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = BORDER
            if row[1].value and "陷阱" in str(row[1].value):
                for cell in row:
                    cell.fill = TRAP_FILL

    wb.save(OUT_XLSX)
    conn.close()
    print(f"\n已生成:\n  {OUT_JSONL}\n  {OUT_XLSX}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
