# -*- coding: utf-8 -*-
"""把 Olist 的 8 个 CSV 导入 MySQL。

为什么自己写导入而不是让 SQLBot 直接读 CSV：
- 需要控制字段类型（金额用 decimal、时间用 datetime），否则生成的 SQL 会算错；
- 需要给表加中文注释，中文提问才能正确映射到英文字段；
- 需要修正原数据集的拼写错误（lenght -> length），否则模型会照着错字段名编。

用法：python load_olist.py
"""
from __future__ import annotations

import os

import csv
from pathlib import Path

import pymysql

HOST, PORT = "127.0.0.1", 13307
USER, PWD = "root", os.environ.get("DB_PASSWORD", "changeme")
DB = "olist"
DATA = Path(r".\olist")

DDL = [
    """CREATE TABLE IF NOT EXISTS customers (
        customer_id VARCHAR(64) PRIMARY KEY COMMENT '客户唯一编号（每笔订单一个）',
        customer_unique_id VARCHAR(64) COMMENT '客户真实唯一编号（同一人可多次下单）',
        customer_zip_code_prefix VARCHAR(10) COMMENT '邮编前五位',
        customer_city VARCHAR(64) COMMENT '客户所在城市',
        customer_state CHAR(2) COMMENT '客户所在州（两位缩写）',
        KEY idx_unique (customer_unique_id)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='客户表'""",
    """CREATE TABLE IF NOT EXISTS orders (
        order_id VARCHAR(64) PRIMARY KEY COMMENT '订单编号',
        customer_id VARCHAR(64) COMMENT '客户编号，关联 customers.customer_id',
        order_status VARCHAR(32) COMMENT '订单状态：delivered已完成 shipped已发货 canceled已取消 等',
        order_purchase_timestamp DATETIME COMMENT '下单时间',
        order_approved_at DATETIME NULL COMMENT '支付审核通过时间',
        order_delivered_carrier_date DATETIME NULL COMMENT '交给物流时间',
        order_delivered_customer_date DATETIME NULL COMMENT '客户签收时间',
        order_estimated_delivery_date DATETIME NULL COMMENT '预计送达时间',
        KEY idx_purchase (order_purchase_timestamp),
        KEY idx_customer (customer_id)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='订单主表'""",
    """CREATE TABLE IF NOT EXISTS order_items (
        order_id VARCHAR(64) NOT NULL COMMENT '订单编号',
        order_item_id INT NOT NULL COMMENT '同一订单内的商品序号',
        product_id VARCHAR(64) NOT NULL COMMENT '商品编号',
        seller_id VARCHAR(64) NOT NULL COMMENT '卖家编号',
        shipping_limit_date DATETIME NULL COMMENT '卖家发货截止时间',
        price DECIMAL(10,2) COMMENT '商品单价（雷亚尔）',
        freight_value DECIMAL(10,2) COMMENT '运费（雷亚尔）',
        PRIMARY KEY (order_id, order_item_id),
        KEY idx_product (product_id),
        KEY idx_seller (seller_id)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='订单商品明细表（一个订单可含多个商品）'""",
    """CREATE TABLE IF NOT EXISTS order_payments (
        order_id VARCHAR(64) NOT NULL COMMENT '订单编号',
        payment_sequential INT NOT NULL COMMENT '支付序号（同一订单可多次支付）',
        payment_type VARCHAR(32) COMMENT '支付方式：credit_card信用卡 boleto票据 voucher代金券 debit_card借记卡',
        payment_installments INT COMMENT '分期数',
        payment_value DECIMAL(10,2) COMMENT '本笔支付金额（雷亚尔）',
        PRIMARY KEY (order_id, payment_sequential),
        KEY idx_type (payment_type)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='订单支付表'""",
    """CREATE TABLE IF NOT EXISTS order_reviews (
        review_id VARCHAR(64) NOT NULL COMMENT '评价编号',
        order_id VARCHAR(64) NOT NULL COMMENT '订单编号',
        review_score INT COMMENT '评分 1-5',
        review_comment_title VARCHAR(255) COMMENT '评价标题',
        review_comment_message TEXT COMMENT '评价内容',
        review_creation_date DATETIME COMMENT '评价问卷发出时间',
        review_answer_timestamp DATETIME COMMENT '评价提交时间',
        KEY idx_order (order_id),
        KEY idx_score (review_score)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='订单评价表'""",
    """CREATE TABLE IF NOT EXISTS products (
        product_id VARCHAR(64) PRIMARY KEY COMMENT '商品编号',
        product_category_name VARCHAR(128) COMMENT '商品品类（葡语）',
        product_name_length INT COMMENT '商品名称长度',
        product_description_length INT COMMENT '商品描述长度',
        product_photos_qty INT COMMENT '商品图片数量',
        product_weight_g INT COMMENT '重量（克）',
        product_length_cm INT COMMENT '长（厘米）',
        product_height_cm INT COMMENT '高（厘米）',
        product_width_cm INT COMMENT '宽（厘米）',
        KEY idx_cat (product_category_name)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='商品表'""",
    """CREATE TABLE IF NOT EXISTS sellers (
        seller_id VARCHAR(64) PRIMARY KEY COMMENT '卖家编号',
        seller_zip_code_prefix VARCHAR(10) COMMENT '邮编前五位',
        seller_city VARCHAR(64) COMMENT '卖家所在城市',
        seller_state CHAR(2) COMMENT '卖家所在州'
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='卖家表'""",
    """CREATE TABLE IF NOT EXISTS category_translation (
        product_category_name VARCHAR(128) PRIMARY KEY COMMENT '品类葡语名',
        product_category_name_english VARCHAR(128) COMMENT '品类英文名'
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='品类名称翻译表（葡语->英文）'""",
]

TABLES = [
    ("customers", "olist_customers_dataset.csv"),
    ("orders", "olist_orders_dataset.csv"),
    ("order_items", "olist_order_items_dataset.csv"),
    ("order_payments", "olist_order_payments_dataset.csv"),
    ("order_reviews", "olist_order_reviews_dataset.csv"),
    ("products", "olist_products_dataset.csv"),
    ("sellers", "olist_sellers_dataset.csv"),
    ("category_translation", "product_category_name_translation.csv"),
]

REMAP = {
    "product_name_lenght": "product_name_length",
    "product_description_lenght": "product_description_length",
}

# 原始数据 2018-09/10 只有 20 单（采集截止造成的稀疏尾巴）。
# 平移后这些会落在"未来"，所以统一删掉，保证每个月的数据都是完整的。
CUTOFF = "2026-11-01"


def load(conn, table: str, path: Path) -> int:
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        header = [REMAP.get(h.strip(), h.strip()) for h in next(reader)]
        cols = ",".join(f"`{c}`" for c in header)
        holders = ",".join(["%s"] * len(header))
        sql = f"INSERT INTO `{table}` ({cols}) VALUES ({holders})"
        batch: list[list] = []
        total = 0
        cur = conn.cursor()
        for row in reader:
            if not row or all(x == "" for x in row):
                continue
            batch.append([None if x == "" else x for x in row])
            if len(batch) >= 5000:
                cur.executemany(sql, batch)
                total += len(batch)
                batch.clear()
        if batch:
            cur.executemany(sql, batch)
            total += len(batch)
        conn.commit()
        cur.close()
        return total


def main() -> int:
    conn = pymysql.connect(host=HOST, port=PORT, user=USER, password=PWD,
                           charset="utf8mb4", autocommit=False)
    with conn.cursor() as cur:
        cur.execute(f"CREATE DATABASE IF NOT EXISTS `{DB}` DEFAULT CHARSET utf8mb4")
        cur.execute(f"USE `{DB}`")
        for stmt in DDL:
            cur.execute(stmt)
    conn.commit()
    print("表结构已创建")

    for table, fname in TABLES:
        path = DATA / fname
        if not path.exists():
            print(f"跳过 {table}（缺文件）")
            continue
        n = load(conn, table, path)
        print(f"{table:<22} 导入 {n:>7} 行")

    # 时间整体平移到近期：原始数据是 2016-09 ~ 2018-08，+98 个月后落在 2024-11 ~ 2026-10
    with conn.cursor() as cur:
        cur.execute("""
            UPDATE orders SET
              order_purchase_timestamp = order_purchase_timestamp + INTERVAL 98 MONTH,
              order_approved_at = order_approved_at + INTERVAL 98 MONTH,
              order_delivered_carrier_date = order_delivered_carrier_date + INTERVAL 98 MONTH,
              order_delivered_customer_date = order_delivered_customer_date + INTERVAL 98 MONTH,
              order_estimated_delivery_date = order_estimated_delivery_date + INTERVAL 98 MONTH
        """)
        cur.execute("UPDATE order_items SET shipping_limit_date = shipping_limit_date + INTERVAL 98 MONTH")
        cur.execute("UPDATE order_reviews SET review_creation_date = review_creation_date + INTERVAL 98 MONTH,"
                    " review_answer_timestamp = review_answer_timestamp + INTERVAL 98 MONTH")
        for t in ("order_items", "order_payments", "order_reviews"):
            cur.execute(f"DELETE FROM {t} WHERE order_id IN "
                        f"(SELECT order_id FROM orders WHERE order_purchase_timestamp >= %s)", (CUTOFF,))
        cur.execute("DELETE FROM orders WHERE order_purchase_timestamp >= %s", (CUTOFF,))
    conn.commit()
    print("时间已整体平移 +8 年 2 个月，并清理稀疏尾巴")

    with conn.cursor() as cur:
        cur.execute("SELECT MIN(order_purchase_timestamp), MAX(order_purchase_timestamp) FROM orders")
        lo, hi = cur.fetchone()
        print(f"订单时间范围：{lo} ~ {hi}")
        cur.execute("SELECT COUNT(*) FROM orders")
        print(f"订单总数：{cur.fetchone()[0]}")
    conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
