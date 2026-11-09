# -*- coding: utf-8 -*-
"""自动跑完 22 道 NL2SQL 评测题并判分。

链路：登录 → 开会话 → 读 SSE（拿 SQL）→ 取 record 数据 → 与标准答案比对 → 出报告

用法：python run_eval_api.py
"""
from __future__ import annotations

import base64
import json
import os
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
DS_ID = int(os.environ.get("SQLBOT_DS", "2"))
CASES = Path(r".\nl2sql_cases_v2.jsonl")
OUT_JSON = Path(r".\nl2sql_api_results.json")
OUT_XLSX = Path(r".\NL2SQL跑分报告.xlsx")


def login(c: httpx.Client) -> str:
    pem = c.get("/api/v1/system/config/key").json()["data"]["public_key"]
    key = serialization.load_pem_public_key(pem.encode())
    enc = lambda t: base64.b64encode(key.encrypt(t.encode(), padding.PKCS1v15())).decode()
    r = c.post("/api/v1/login/access-token", data={"username": enc("admin"), "password": enc(PWD)})
    r.raise_for_status()
    return r.json()["data"]["access_token"]


def ask(c: httpx.Client, h: dict, question: str) -> dict:
    chat = c.post("/api/v1/chat/start", headers=h,
                  json={"datasource": DS_ID, "question": question, "origin": 0})
    chat.raise_for_status()
    chat_id = chat.json()["data"]["id"]

    buffers: dict[str, str] = {}
    record_id, error = None, None
    with c.stream("POST", "/api/v1/chat/question", headers=h,
                  json={"chat_id": chat_id, "question": question, "datasource_id": DS_ID}) as resp:
        resp.raise_for_status()
        for line in resp.iter_lines():
            if not line or not line.startswith("data:"):
                continue
            try:
                evt = json.loads(line[5:].strip())
            except json.JSONDecodeError:
                continue
            kind = evt.get("type")
            if kind == "id":
                record_id = evt.get("id")
            elif kind == "error":
                error = evt.get("content") or evt.get("message")
            elif "content" in evt:
                buffers[kind] = buffers.get(kind, "") + (evt.get("content") or "")

    sql, success = None, None
    if buffers.get("sql-result"):
        try:
            parsed = json.loads(buffers["sql-result"])
            success, sql = parsed.get("success"), parsed.get("sql")
        except json.JSONDecodeError:
            pass

    fields, rows = [], []
    if record_id:
        try:
            d = c.get(f"/api/v1/chat/record/{record_id}/data", headers=h).json().get("data") or {}
            fields = d.get("fields") or []
            rows = d.get("data") or []
        except Exception as exc:
            error = error or f"取数据失败: {exc}"
    return {"sql": sql, "success": success, "fields": fields, "rows": rows, "error": error}


def judge(expected, rows, fields):
    """返回 (判定, 模型答案, 误差%)"""
    if not rows:
        return "无数据", None, None
    first = rows[0]
    val = first.get(fields[0]) if fields else next(iter(first.values()))
    if isinstance(expected, list):
        # 多行结果：比较行数与第一列序列
        if len(rows) < len(expected):
            return "行数不足", str(len(rows)), None
        ok = True
        for i, exp_row in enumerate(expected):
            model_row = rows[i]
            exp_first = str(exp_row[0])
            got_first = str(next(iter(model_row.values())))
            if got_first != exp_first:
                ok = False
                break
        return ("对" if ok else "错"), str(val), None
    try:
        a, b = float(val), float(expected)
        diff = abs(a - b)
        pct = 0.0 if b == 0 else diff / abs(b) * 100
        if diff <= 0.011:
            return "对", a, pct
        return ("接近" if pct <= 5 else "错"), a, pct
    except (TypeError, ValueError):
        return ("对" if str(val) == str(expected) else "错"), val, None


def main() -> int:
    cases = [json.loads(l) for l in CASES.read_text(encoding="utf-8").splitlines() if l.strip()]
    results = []
    with httpx.Client(base_url=BASE, timeout=180) as c:
        h = {"X-SQLBOT-TOKEN": f"Bearer {login(c)}"}
        for i, case in enumerate(cases, 1):
            t0 = time.time()
            try:
                r = ask(c, h, case["question"])
            except Exception as exc:
                r = {"sql": None, "success": False, "fields": [], "rows": [], "error": f"{type(exc).__name__}: {exc}"}
            verdict, answer, pct = judge(case["expected"], r["rows"], r["fields"])
            results.append({**case, "model_sql": r["sql"], "model_rows": r["rows"][:5],
                            "model_answer": answer, "verdict": verdict, "diff_pct": pct,
                            "error": r["error"], "seconds": round(time.time() - t0, 1)})
            print(f"[{i:>2}/22] {case['id']:<4} {verdict:<6} 模型={str(answer)[:26]:<28} 期望={str(case['expected'])[:26]:<28} {r['error'] or ''}")
            time.sleep(0.4)

    OUT_JSON.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")

    ok = sum(1 for r in results if r["verdict"] == "对")
    near = sum(1 for r in results if r["verdict"] == "接近")
    print(f"\n准确率: {ok}/22 = {ok/22:.1%}  （接近 {near} 条，无数据 {sum(1 for r in results if r['verdict']=='无数据')} 条）")

    wb = Workbook()
    ws = wb.active
    ws.title = "跑分结果"
    head = ["编号", "分类", "问题", "标准答案", "模型答案", "判定", "误差%", "耗时s", "模型 SQL", "错误"]
    ws.append(head)
    for i, w in enumerate([7, 16, 42, 26, 22, 8, 8, 8, 60, 20], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        cell = ws.cell(row=1, column=i)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for r in results:
        exp = r["expected"]
        exp = " | ".join(" ".join(map(str, x)) for x in exp) if isinstance(exp, list) else exp
        ws.append([r["id"], r["category"], r["question"], exp, r["model_answer"],
                   r["verdict"], None if r["diff_pct"] is None else round(r["diff_pct"], 2),
                   r["seconds"], r["model_sql"], r["error"]])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        if row[5].value == "错":
            for cell in row:
                cell.fill = PatternFill("solid", fgColor="FFD7D7")
        elif row[5].value == "对":
            for cell in row:
                cell.fill = PatternFill("solid", fgColor="DFF5E1")
    ws.freeze_panes = "A2"
    wb.save(OUT_XLSX)
    print(f"报告 -> {OUT_XLSX}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
