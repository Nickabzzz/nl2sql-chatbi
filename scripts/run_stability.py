# -*- coding: utf-8 -*-
"""加 Q12 的 SQL 示例，然后跑 3 轮稳定性测试。

判定改进（修掉之前两个判分 bug）：
1. 数字带百分号 / 千分位也能解析；
2. 多列结果会在**所有列**里找匹配值（T04 的标准答案在第二列）。
"""
from __future__ import annotations

import base64
import json
import os
import statistics
import time
from pathlib import Path

import httpx
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import padding
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = "http://127.0.0.1:8000"
PWD = os.environ["SQLBOT_PWD"]
DS_ID = 2
ROUNDS = 3
CASES = Path(r".\nl2sql_cases_v2.jsonl")
OUT_JSON = Path(r".\stability_results.json")
OUT_XLSX = Path(r".\NL2SQL稳定性报告.xlsx")

TRAINING = {
    "question": "按订单数计算，使用分期（分期数大于1）的订单占比是多少",
    "description": (
        "SELECT ROUND(100*COUNT(DISTINCT CASE WHEN payment_installments>1 THEN order_id END)"
        "/COUNT(DISTINCT order_id),2) AS installment_order_ratio FROM order_payments"
    ),
    "datasource": DS_ID,
    "enabled": True,
}


def login(c: httpx.Client) -> str:
    pem = c.get("/api/v1/system/config/key").json()["data"]["public_key"]
    key = serialization.load_pem_public_key(pem.encode())
    enc = lambda t: base64.b64encode(key.encrypt(t.encode(), padding.PKCS1v15())).decode()
    r = c.post("/api/v1/login/access-token", data={"username": enc("admin"), "password": enc(PWD)})
    r.raise_for_status()
    return r.json()["data"]["access_token"]


def ask(c, h, question):
    chat = c.post("/api/v1/chat/start", headers=h,
                  json={"datasource": DS_ID, "question": question, "origin": 0}).json()["data"]
    chat_id = chat["id"]
    buffers, record_id, error = {}, None, None
    with c.stream("POST", "/api/v1/chat/question", headers=h,
                  json={"chat_id": chat_id, "question": question, "datasource_id": DS_ID}) as resp:
        for line in resp.iter_lines():
            if not line.startswith("data:"):
                continue
            try:
                evt = json.loads(line[5:].strip())
            except json.JSONDecodeError:
                continue
            k = evt.get("type")
            if k == "id":
                record_id = evt.get("id")
            elif k == "error":
                error = evt.get("content") or evt.get("message")
            elif "content" in evt:
                buffers[k] = buffers.get(k, "") + (evt.get("content") or "")
    sql = None
    if buffers.get("sql-result"):
        try:
            sql = json.loads(buffers["sql-result"]).get("sql")
        except json.JSONDecodeError:
            pass
    fields, rows = [], []
    if record_id:
        d = c.get(f"/api/v1/chat/record/{record_id}/data", headers=h).json().get("data") or {}
        fields, rows = d.get("fields") or [], d.get("data") or []
    return {"sql": sql, "fields": fields, "rows": rows, "error": error}


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("%", "").replace(" ", "")
    try:
        return float(s)
    except ValueError:
        return None


def judge(expected, rows, fields):
    if not rows:
        return "无数据"
    if isinstance(expected, list):
        if len(rows) < len(expected):
            return "行数不足"
        for i, exp_row in enumerate(expected):
            if str(next(iter(rows[i].values()))) != str(exp_row[0]):
                return "错"
        return "对"
    target = num(expected)
    for row in rows[:3]:
        for key in fields or row.keys():
            if key in row and num(row[key]) is not None and target is not None:
                if abs(num(row[key]) - target) <= max(0.011, abs(target) * 0.0005):
                    return "对"
    return "错"


def main() -> int:
    cases = [json.loads(l) for l in CASES.read_text(encoding="utf-8").splitlines() if l.strip()]
    with httpx.Client(base_url=BASE, timeout=180) as c:
        h = {"X-SQLBOT-TOKEN": f"Bearer {login(c)}"}

        r = c.put("/api/v1/system/data-training", headers=h, json=TRAINING)
        print("新增 SQL 示例:", r.status_code, r.text[:100])

        results = {case["id"]: {"question": case["question"], "expected": case["expected"], "rounds": []}
                   for case in cases}
        for rnd in range(1, ROUNDS + 1):
            t0 = time.time()
            ok = 0
            for case in cases:
                try:
                    res = ask(c, h, case["question"])
                    verdict = judge(case["expected"], res["rows"], res["fields"])
                except Exception as exc:
                    res, verdict = {"error": str(exc), "sql": None, "rows": []}, "异常"
                results[case["id"]]["rounds"].append(
                    {"verdict": verdict, "answer": (res["rows"][0] if res["rows"] else None),
                     "sql": res["sql"], "error": res["error"]})
                ok += verdict == "对"
            print(f"第 {rnd} 轮：{ok}/{len(cases)} = {ok/len(cases):.1%}  用时 {time.time()-t0:.0f}s")

    summary = []
    for cid, info in results.items():
        verdicts = [r["verdict"] for r in info["rounds"]]
        correct = verdicts.count("对")
        summary.append({"id": cid, "question": info["question"], "expected": info["expected"],
                        "verdicts": verdicts, "correct": correct,
                        "stable": correct == ROUNDS})
    stable = sum(1 for s in summary if s["stable"])
    print(f"\n3 轮全部答对（稳定）：{stable}/22")
    rates = [sum(1 for r in info["rounds"] if r["verdict"] == "对") / ROUNDS for info in results.values()]
    print(f"三轮平均准确率：{statistics.mean(rates):.1%}")
    print(f"逐题波动：最低 {min(rates):.0%}｜最高 {max(rates):.0%}")

    OUT_JSON.write_text(json.dumps({"summary": summary, "detail": list(results.values())},
                                   ensure_ascii=False, indent=2), encoding="utf-8")

    wb = Workbook()
    ws = wb.active
    ws.title = "稳定性汇总"
    head = ["编号", "问题", "标准答案", "第1轮", "第2轮", "第3轮", "答对轮次", "是否稳定"]
    ws.append(head)
    for i, w in enumerate([7, 44, 22, 10, 10, 10, 10, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        cell = ws.cell(row=1, column=i)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for s in summary:
        exp = s["expected"]
        exp = " | ".join(" ".join(map(str, x)) for x in exp) if isinstance(exp, list) else exp
        ws.append([s["id"], s["question"], exp, *s["verdicts"], s["correct"], "稳定" if s["stable"] else "波动"])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        color = "DFF5E1" if row[6].value == ROUNDS else ("FFF2CC" if row[6].value else "FFD7D7")
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=color)
    ws.freeze_panes = "A2"
    wb.save(OUT_XLSX)
    print(f"报告 -> {OUT_XLSX}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
