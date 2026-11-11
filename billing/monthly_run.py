# -*- coding: utf-8 -*-
"""月次の請求まわりを回す（Excel 正本の運用）。

  ingest  … 前月分を VMS から取り、Drive のブックにデータ行を追記する（毎月2日。定時実行）
  remind  … 金額を出して確認を促す（4日ごろ。現在は手動実行。定時実行への切り替えを予定）
  draft   … 検証を通してから請求書PDFをドラフト出力する（5日ごろ。手動）

実運用ではこのほかに import-other（入力用スプシの「その他売上」を Sheets API で読み、
対象月の行だけ置き換える）があるが、この公開版には含めていない。

ブックは Drive 上の1個が唯一の正。毎月作り直さない（作り直すと手入力が消える）。
build_workbook.py は初回の型作り専用で、運用中は走らせないこと。
"""
import datetime
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.request

import openpyxl
from copy import copy
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
BOOK = os.environ.get("BOOK", os.path.join(os.environ.get("LEDGER_DIR", "."),
                                           "sales-ledger.xlsx"))
DRIVE = os.path.dirname(BOOK)
EXPORTER = os.path.join(REPO, "vms-exporter", "lib", "vms-exporter.exe")
CONFIG = os.path.join(REPO, "vms-exporter", "config.local.json")
SOFFICE = os.environ.get("SOFFICE", os.path.join("C:\\", "Program Files", "LibreOffice",
                                                 "program", "soffice.exe"))
DATA = os.path.join(HERE, "data")
LOGDIR = os.path.join(HERE, "logs")

sys.path.insert(0, HERE)
from build_workbook import load_year, write_evidence, read_contracts, invoice_sheets

WEEKDAYS = ["月", "火", "水", "木", "金", "土", "日"]


def log(msg):
    print(datetime.datetime.now().strftime("%H:%M:%S ") + msg)
    os.makedirs(LOGDIR, exist_ok=True)
    path = os.path.join(LOGDIR, "monthly_%s.log" % datetime.date.today().strftime("%Y%m"))
    with open(path, "a", encoding="utf-8") as f:
        f.write(datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S ") + msg + "\n")


def discord(msg):
    try:
        cfg = json.load(open(CONFIG, encoding="utf-8-sig"))
        url = cfg.get("notifyWebhookUrl") or cfg.get("discordWebhookUrl", "")
    except Exception:
        url = ""
    if not url:
        log("Discord Webhook 未設定のため送信スキップ")
        return
    body = json.dumps({"content": msg[:1900]}).encode("utf-8")
    # 既定の User-Agent（Python-urllib）は Discord 側に 403 で弾かれる
    req = urllib.request.Request(url, data=body, headers={
        "Content-Type": "application/json",
        "User-Agent": "vending-monthly/1.0 (+vms-exporter)",
    })
    try:
        urllib.request.urlopen(req, timeout=30)
    except Exception as exc:
        log("Discord送信に失敗: %s" % exc)


def issue_date(ym=None, today=None):
    """請求書の発行日＝PDFを作った日。発行日は運用側の判断で「作成当日」に決まった。"""
    return today or datetime.date.today()


def prev_month(today=None):
    d = today or datetime.date.today()
    last = datetime.date(d.year, d.month, 1) - datetime.timedelta(days=1)
    return "%04d-%02d" % (last.year, last.month)


def check_lock():
    """Excel で開かれていると書けない。Drive の競合コピーも避けたいので先に見る。"""
    lock = os.path.join(DRIVE, "~$" + os.path.basename(BOOK))
    if os.path.exists(lock):
        raise SystemExit("ブックが開かれています（%s）。閉じてから再実行してください。" % lock)
    if not os.path.exists(BOOK):
        raise SystemExit("ブックが見つかりません: %s" % BOOK)
    try:
        with open(BOOK, "r+b"):
            pass
    except OSError as exc:
        raise SystemExit("ブックに書き込めません（開かれている可能性）: %s" % exc)


def recalc(path):
    """LibreOffice で再計算した写しを返す。openpyxl は数式を計算しないため。"""
    tmp = tempfile.mkdtemp()
    work = os.path.join(tmp, "r.xlsx")
    wb = openpyxl.load_workbook(path)
    wb.calculation.fullCalcOnLoad = True
    wb.save(work)
    subprocess.run([SOFFICE, "--headless", "--norestore", "--convert-to", "xlsx",
                    "--outdir", os.path.join(tmp, "out"), work],
                   check=True, capture_output=True, timeout=600)
    return openpyxl.load_workbook(os.path.join(tmp, "out", "r.xlsx"), data_only=True), tmp


def totals(calc):
    """再計算済みブックから (支払先, 合計税込, シート名, 請求書番号) を拾う。"""
    out = []
    for sheet, payee in sorted(invoice_sheets(calc).items(), key=lambda kv: kv[1]):
        ws = calc[sheet]
        for r in range(1, ws.max_row + 1):
            if ws.cell(r, 1).value == "合計（税込）":
                v = next((ws.cell(r, c).value for c in range(2, 9)
                          if isinstance(ws.cell(r, c).value, (int, float))), 0)
                out.append((payee, v, sheet, ws["F2"].value or ""))
                break
    return out


def fmt(rows):
    lines = ["　%-26s %10s円" % (r[0], format(int(r[1]), ",")) for r in rows]
    lines.append("　%-26s %10s円" % ("合計", format(int(sum(r[1] for r in rows)), ",")))
    return "\n".join(lines)


# ---------------------------------------------------------------- ingest
def fetch(ym):
    out = os.path.join(DATA, ym)
    os.makedirs(out, exist_ok=True)
    log("VMS から %s を取得中…" % ym)
    r = subprocess.run([EXPORTER, "export-sales", ym, ym, out],
                       capture_output=True, timeout=1800)
    if r.returncode != 0:
        raise SystemExit("取得に失敗: " + r.stdout.decode("utf-8", "replace")[-500:])
    # 検証スクリプトが読む位置にも最新月を置く
    for name in ("vms_year.tsv", "vms_year_kaishu.tsv"):
        shutil.copyfile(os.path.join(out, name), os.path.join(HERE, name))
    return out


def append_rows(ws, rows, ym):
    """取込シートに1ヶ月ぶん追記する。すでにその月が入っていれば何もしない（再実行しても安全）。"""
    have = {ws.cell(r, 1).value for r in range(5, ws.max_row + 1)}
    if ym in have:
        return 0
    r = ws.max_row + 1
    src = ws.max_row                      # 直前の行から書式を写す
    for d in rows:
        y, m, dd = (int(x) for x in d["date"].split("-"))
        vals = [d["ym"], d["date"], d["code"], d["machine"], d["qty"], d["cash"], d["card"],
                "=F%d+G%d" % (r, r), WEEKDAYS[datetime.date(y, m, dd).weekday()]]
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            s = ws.cell(src, i)
            c.font = copy(s.font)
            c.border = copy(s.border)
            c.number_format = s.number_format
        r += 1
    return len(rows)


def cmd_ingest(ym):
    check_lock()
    out = fetch(ym)
    daily = load_year(os.path.join(out, "vms_year.tsv"))
    kaishu = load_year(os.path.join(out, "vms_year_kaishu.tsv"))
    wb = openpyxl.load_workbook(BOOK)
    n1 = append_rows(wb["取込_夜間"], daily, ym)
    n2 = append_rows(wb["取込_回収時"], kaishu, ym)
    wb["集計_号機別"]["B3"] = ym
    wb.save(BOOK)
    if n1 == 0 and n2 == 0:
        msg = "ℹ️ %s はすでに取り込み済みでした。追記していません。" % ym
    else:
        msg = ("📥 **%s の売上を取り込みました**\n"
               "夜間検量 %d行 ／ 回収時検量 %d行 → %s\n"
               "入力が要るものがあれば5日までにお願いします（卸売上・イレギュラー稼働）。\n"
               "4日に金額をお知らせし、5日に請求書ドラフトを出します。"
               % (ym, n1, n2, os.path.basename(BOOK)))
    log(msg.replace("\n", " / "))
    discord(msg)


# ---------------------------------------------------------------- remind
def cmd_remind(ym):
    calc, tmp = recalc(BOOK)
    rows = totals(calc)
    shutil.rmtree(tmp, ignore_errors=True)
    log("remind: " + " / ".join("%s=%s" % (r[0], r[1]) for r in rows))
    discord("🔎 **%s の請求額（ご確認ください）**\n```\n%s\n```\n"
            "違っていれば5日までに %s を直してください。この内容で5日にドラフトを出します。"
            % (ym, fmt(rows), os.path.basename(BOOK)))


# ---------------------------------------------------------------- draft
def rewrite_evidence(wb, ym):
    """各請求書の【算定根拠】を対象月で書き直す。回収回数が月ごとに違うので数式にできない。"""
    kaishu = load_year(os.path.join(HERE, "vms_year_kaishu.tsv"))
    detail = {}
    for d in kaishu:
        detail.setdefault((d["ym"], d["code"]), []).append(d)
    for v in detail.values():
        v.sort(key=lambda x: x["date"])
    contracts = read_contracts(wb)
    for sheet, payee in invoice_sheets(wb).items():
        ws = wb[sheet]
        top = next((r for r in range(1, ws.max_row + 1)
                    if isinstance(ws.cell(r, 1).value, str)
                    and ws.cell(r, 1).value.startswith("【算定根拠】")), None)
        if top is None:
            continue
        ws.delete_rows(top, ws.max_row - top + 1)
        rows_c = [c for c in contracts if c[0] == payee]
        ids = [c[1] for c in rows_c if isinstance(c[1], int)]
        end = write_evidence(ws, top, ids, ym, detail)
        ncols = 8 if any(c[2] == "現金渡し" for c in rows_c) else 6
        ws.print_area = "A1:%s%d" % (get_column_letter(ncols), end)


def cmd_draft(ym):
    check_lock()
    log("検証を実行します…")
    r = subprocess.run([sys.executable, os.path.join(HERE, "verify_workbook.py"), BOOK, ym],
                       capture_output=True, cwd=HERE, timeout=1800)
    report = r.stdout.decode("utf-8", "replace")
    print(report)
    if r.returncode != 0:
        tail = "\n".join(l for l in report.splitlines() if l.strip().startswith(("✗", "⚠")))
        discord("🛑 **%s の請求書は出力を止めました**\n検証で問題があります。\n```\n%s\n```" % (ym, tail))
        raise SystemExit(1)

    issue = issue_date(ym)      # PDF作成日
    wb = openpyxl.load_workbook(BOOK)
    for r_ in range(5, 20):
        if wb["M_請求設定"].cell(r_, 1).value == "発行日":
            c = wb["M_請求設定"].cell(r_, 2)
            c.value = issue
            c.number_format = "yyyy/mm/dd"
    wb["集計_号機別"]["B3"] = ym
    rewrite_evidence(wb, ym)
    wb.save(BOOK)

    outdir = os.path.join(DRIVE, "請求書_%s_ドラフト" % ym)
    shutil.rmtree(outdir, ignore_errors=True)
    from export_invoices import export
    made = export(BOOK, outdir)
    log("PDF %d枚を出力: %s" % (len(made), outdir))

    calc, tmp = recalc(BOOK)
    rows = totals(calc)
    shutil.rmtree(tmp, ignore_errors=True)

    wb = openpyxl.load_workbook(BOOK)
    ledger = wb["発行台帳"]
    r_ = next(r for r in range(5, ledger.max_row + 2) if not ledger.cell(r, 1).value)
    for payee, amt, sheet, no in rows:
        ledger.cell(r_, 1, ym)
        ledger.cell(r_, 2, payee)
        ledger.cell(r_, 3, no)
        ledger.cell(r_, 4, issue).number_format = "yyyy/mm/dd"
        ledger.cell(r_, 5, amt)
        ledger.cell(r_, 6, "%s_%s.pdf" % (ym, sheet[len("請求_"):]))
        ledger.cell(r_, 7, "ドラフト（未承認）")
        r_ += 1
    wb.save(BOOK)

    discord("📄 **%s の請求書ドラフトができました**（%d枚・発行日 %s）\n```\n%s\n```\n"
            "保存先: %s\n**内容を確認して、承認できたら発行してください。**"
            % (ym, len(made), issue, fmt(rows), outdir))


COMMANDS = {"ingest": cmd_ingest, "remind": cmd_remind, "draft": cmd_draft}

if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    ym = sys.argv[2] if len(sys.argv) > 2 else prev_month()
    if cmd not in COMMANDS:
        raise SystemExit(__doc__)
    if not re.match(r"^\d{4}-\d{2}$", ym):
        raise SystemExit("対象年月の形式が不正: %s" % ym)
    try:
        COMMANDS[cmd](ym)
    except SystemExit:
        raise
    except Exception as exc:
        log("ERROR: %r" % exc)
        discord("⚠ 月次処理（%s %s）でエラー: %s" % (cmd, ym, exc))
        raise
