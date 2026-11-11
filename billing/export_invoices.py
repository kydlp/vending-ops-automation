# -*- coding: utf-8 -*-
"""請求書シートを支払先ごとに1枚のPDFへ書き出す。

LibreOffice はブック単位でしかPDF化できないため、支払先ごとに
「その請求書シート以外を非表示にしたコピー」を作って変換する。
非表示シートも数式の参照先としては生きているので、金額は壊れない。
"""
import os, re, shutil, subprocess, sys, tempfile
import openpyxl

SOFFICE = os.environ.get("SOFFICE", r"C:\Program Files\LibreOffice\program\soffice.exe")
INVALID = re.compile(r'[\\/:*?"<>|]')


def export(src, outdir, month=None):
    os.makedirs(outdir, exist_ok=True)
    wb = openpyxl.load_workbook(src)
    if month:
        wb["集計_号機別"]["B3"] = month
        src2 = os.path.join(tempfile.mkdtemp(), os.path.basename(src))
        wb.save(src2); src = src2
        wb = openpyxl.load_workbook(src)
    ym = month or wb["集計_号機別"]["B3"].value
    # 支払先（J1）が空のシートは未使用の予備。PDF化しない。
    sheets = [s for s in wb.sheetnames if s.startswith("請求_") and wb[s]["J1"].value]
    if not sheets:
        raise SystemExit("請求書シートが見つからない")

    made = []
    tmp = tempfile.mkdtemp()
    for name in sheets:
        work = os.path.join(tmp, "w.xlsx")
        shutil.copyfile(src, work)
        b = openpyxl.load_workbook(work)
        for s in b.worksheets:
            s.sheet_state = "visible" if s.title == name else "hidden"
            # 非表示にするだけでは LibreOffice が印刷範囲のあるシートを全部出してしまう。
            # 対象以外の印刷範囲を消して、1枚だけがPDFになるようにする。
            if s.title != name:
                s.print_area = None
        b.active = b.sheetnames.index(name)
        b.calculation.fullCalcOnLoad = True
        b.save(work)

        subprocess.run([SOFFICE, "--headless", "--norestore",
                        "--convert-to", "pdf:calc_pdf_Export",
                        "--outdir", tmp, work],
                       check=True, capture_output=True, timeout=180)
        produced = os.path.join(tmp, "w.pdf")
        if not os.path.exists(produced):
            raise SystemExit(f"PDF化に失敗: {name}")

        payee = name[len("請求_"):]
        num = b[name]["F2"].value
        code = num.rsplit('"', 2)[-2] if isinstance(num, str) and '"' in num else ""
        dst = os.path.join(outdir, INVALID.sub("_", f"{ym}_{payee}.pdf"))
        shutil.move(produced, dst)
        made.append(dst)
    shutil.rmtree(tmp, ignore_errors=True)
    return made


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "sales-ledger.xlsx"
    outdir = sys.argv[2] if len(sys.argv) > 2 else "invoices"
    month = sys.argv[3] if len(sys.argv) > 3 else None
    for p in export(src, outdir, month):
        print(p)
