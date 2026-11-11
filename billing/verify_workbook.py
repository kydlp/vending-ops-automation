# -*- coding: utf-8 -*-
"""生成したブックの検証。月次運用では、請求書PDFを出す前にこれを通す。

1) 全数式を走査し、使用関数が Excel2007 世代／Google スプレッドシート互換かを確認
2) LibreOffice で再計算し、エラーセル（#REF! 等）が出ていないかを確認
3) 請求額を Python で独立に再現し、シートの値と突き合わせる
4) 「このまま発行してはいけない」条件を洗い出す（発行日未記入・0円・条件未確定など）

使い方: python verify_workbook.py [ブック.xlsx] [対象年月]
戻り値: 問題があれば終了コード 1（月次スクリプトはこれを見て PDF 出力を止める）
"""
import datetime, math, os, re, shutil, subprocess, sys, tempfile
from collections import defaultdict
import openpyxl

sys.stdout.reconfigure(encoding="utf-8")

from build_workbook import (MACHINES, MNAME, WHOLESALE, WHOLESALE_ID, OTHER_SALES,
                            OTHER_AMOUNT, CLOSED_WEEKDAYS, load_year, read_contracts,
                            invoice_sheets)

SOFFICE = os.environ.get("SOFFICE", r"C:\Program Files\LibreOffice\program\soffice.exe")
ALLOWED = {"SUM", "SUMIFS", "COUNTIFS", "INDEX", "MATCH", "IFERROR", "IF",
           "ROUND", "ROUNDUP", "MAX", "SUBSTITUTE", "TODAY"}
ERRORS = ("#REF!", "#N/A", "#VALUE!", "#NAME?", "#DIV/0!", "#NULL!", "#NUM!")

src = sys.argv[1] if len(sys.argv) > 1 else "sales-ledger.xlsx"
problems, warnings = [], []


# ---------- 1) 数式の互換性 ----------
wb = openpyxl.load_workbook(src)
target_ym = sys.argv[2] if len(sys.argv) > 2 else wb["集計_号機別"]["B3"].value
funcs, n_formula = set(), 0
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                n_formula += 1
                # 文字列リテラル内（例: "WholesaleX(卸)"）を関数と誤検出しないよう、引用符の中身を除いてから拾う
                body = re.sub(r'"[^"]*"', '""', c.value)
                funcs |= set(re.findall(r"(?<![A-Za-z0-9_.])([A-Z][A-Z0-9\.]*)\s*\(", body))
contracts = read_contracts(wb)
sheets = invoice_sheets(wb)                      # 請求書シート名 → 支払先
payees = sorted(set(sheets.values()))
sheet_of = {v: k for k, v in sheets.items()}
bad = funcs - ALLOWED
print(f"[1] 数式 {n_formula} 個 / 使用関数 {sorted(funcs)}")
print("    互換性:", "OK（すべて2007世代・Googleスプレッドシート互換）" if not bad else f"NG {sorted(bad)}")
if bad:
    problems.append(f"Sheets非互換の関数: {sorted(bad)}")
print(f"    対象年月: {target_ym}")


# ---------- 2) 再計算してエラーセルを見る ----------
tmp = tempfile.mkdtemp()
work = os.path.join(tmp, "v.xlsx")
wb.calculation.fullCalcOnLoad = True
wb.save(work)
subprocess.run([SOFFICE, "--headless", "--norestore", "--convert-to", "xlsx",
                "--outdir", os.path.join(tmp, "out"), work],
               check=True, capture_output=True, timeout=300)
calc = openpyxl.load_workbook(os.path.join(tmp, "out", "v.xlsx"), data_only=True)
errs = defaultdict(int)
for ws in calc.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value in ERRORS:
                errs[(ws.title, c.value)] += 1
print(f"\n[2] 再計算後のエラーセル: {dict(errs) if errs else 'なし'}")
if errs:
    problems.append(f"エラーセル: {dict(errs)}")


# ---------- 3) 請求額を Python で再現して突合 ----------
kaishu = [d for d in load_year("vms_year_kaishu.tsv") if d["ym"] == target_ym]
agg = defaultdict(lambda: dict(qty=0.0, cash=0.0, card=0.0, n=0))
for d in kaishu:
    a = agg[d["code"]]
    a["qty"] += d["qty"]; a["cash"] += d["cash"]; a["card"] += d["card"]; a["n"] += 1
for a in agg.values():
    a["total"] = a["cash"] + a["card"]

y, m = (int(x) for x in target_ym.split("-"))
d1 = datetime.date(y, m, 1)
d2 = datetime.date(y + (m == 12), (m % 12) + 1, 1)
workdays = sum(1 for k in range((d2 - d1).days)
               if (d1 + datetime.timedelta(days=k)).weekday() not in CLOSED_WEEKDAYS)

other = {}
for ym, _kind, name, qty, price, _n in OTHER_SALES:
    if ym == target_ym:
        other[name] = qty * price if (qty and price) else OTHER_AMOUNT.get(name, 0)


def expected(payee):
    total = 0
    for (p, target, method, rate, fee, monthly, floor_, taxmode, _n) in contracts:
        if p != payee:
            continue
        tid = target
        if tid == WHOLESALE_ID:
            sales, cash = other.get(WHOLESALE, 0), 0
        else:
            a = agg.get(tid, dict(total=0, cash=0))
            sales, cash = a["total"], a["cash"]
        if method == "料率":
            # 浮動小数の誤差（例: 77350*0.14=10829.000000000002）で切り上げが1円ずれないよう丸めてから切り上げる
            base = (math.ceil(round(sales / 1.08 * rate * 1.1, 6)) if taxmode == "税抜換算"
                    else math.ceil(round(sales * rate, 6)))
            total += max(base, floor_)
        elif method == "現金渡し":
            total += cash - fee * workdays - monthly
    return total


print(f"\n[3] 請求額の再現（稼働日数 {workdays}日・定休曜日を除く）")
print(f"    {'支払先':<24}{'シート':>12}{'再現':>12}{'差':>8}")
for payee in payees:
    ws = calc[sheet_of[payee]]
    got = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "合計（税込）":
            got = next((ws.cell(r, c).value for c in range(2, 9)
                        if isinstance(ws.cell(r, c).value, (int, float))), None)
            break
    exp = expected(payee)
    diff = "" if got is None else f"{got - exp:+,}"
    mark = "" if got == exp else "  ← 不一致"
    print(f"    {payee:<24}{got if got is not None else '-':>12,}{exp:>12,}{diff:>8}{mark}")
    if got != exp:
        problems.append(f"{payee}: シート {got} ≠ 再現 {exp}")


# ---------- 4) 発行してよいかのチェック ----------
print("\n[4] 発行前チェック")
setting = {calc["M_請求設定"].cell(r, 1).value: calc["M_請求設定"].cell(r, 2).value
           for r in range(5, 20)}
if not setting.get("発行日"):
    warnings.append("M_請求設定 の発行日が空。PDFを出す前に発行日を入れること")
for k in ("インボイス登録番号", "振込先", "電話番号", "消費税の扱い"):
    if not setting.get(k):
        problems.append(f"M_請求設定 の「{k}」が空")

for payee in payees:
    for (p, target, method, *_r) in contracts:
        if p == payee and method == "要確認":
            problems.append(f"{payee}×号機ID {target}: 契約条件が未確定（要確認）")
    if expected(payee) == 0:
        warnings.append(f"{payee}: 請求額が0円")

# M_契約 に足したのに請求書シートを用意し忘れると、その支払先だけ請求書が出ない。
for p_ in sorted({c[0] for c in contracts}):
    if p_ not in sheet_of:
        problems.append(f"{p_}: M_契約 にあるが請求書シートが無い。"
                        f"予備シート（請求_予備1〜3）の J1 に支払先名を入れること")

for mid, name, *_r in MACHINES:
    if mid not in agg:
        warnings.append(f"号機 {mid}（{name}）: {target_ym} の回収実績が0件")

# VMS に増えた号機がマスタに無いと、売上が誰の集計にも乗らず静かに消える。
known = {m[0] for m in MACHINES}
for code in sorted(set(agg) - known):
    problems.append(f"号機ID {code} が VMS にあるがマスタ（MACHINES / M_号機）に無い。"
                    f"売上 {int(agg[code]['total']):,}円 が集計から漏れている")

# 発行済みの控えと金額が変わっていたら、遡及計算になっている（料率や税方式を後から直した等）。
try:
    lg = calc["発行台帳"]
    for r in range(5, lg.max_row + 1):
        ym_, payee_, _no, _d, amt = (lg.cell(r, c).value for c in range(1, 6))
        if ym_ == target_ym and payee_ and isinstance(amt, (int, float)):
            now = expected(payee_)
            if abs(now - amt) > 0.5:
                problems.append(f"{payee_}: 発行台帳は {int(amt):,}円 だが今の計算は {int(now):,}円。"
                                f"発行済みの控えと食い違う")
except KeyError:
    pass
sheet_names = {calc["集計_号機別"].cell(r, 2).value for r in range(6, 6 + len(MACHINES))}
if None in sheet_names or "" in sheet_names:
    problems.append("集計_号機別 の号機名が引けていない（M_号機 と号機IDが合っていない）")

for w in warnings:
    print("    ⚠ " + w)
for x in problems:
    print("    ✗ " + x)
if not warnings and not problems:
    print("    問題なし")

shutil.rmtree(tmp, ignore_errors=True)
print(f"\n判定: {'NG（' + str(len(problems)) + '件）' if problems else 'OK'}"
      f"{' / 警告 ' + str(len(warnings)) + '件' if warnings else ''}")
sys.exit(1 if problems else 0)
