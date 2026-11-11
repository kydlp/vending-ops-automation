# -*- coding: utf-8 -*-
"""合成データ（sample-data/sales_sample.tsv）を、build_workbook.py / verify_workbook.py が読む
2本のTSV（夜間検量・回収時検量）に分けて、このフォルダに置く。

本番では VMS からの取得（monthly_run.py ingest）がこの2本を作る。
合成データの回収時の1行は「前回の回収日〜回収日の前日」の夜間の合計にしてある。
分けたあと、回収時の各行が夜間と合っているかを確かめ、号機×月の合計を並べて出す。
回収時と夜間の月合計の差は、月末時点でまだ回収していない分と、前月から持ち越した分だけになる。
使い方: python prepare_sample.py
"""
import io, os, sys
from collections import defaultdict

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "sample-data", "sales_sample.tsv")
OUT = {"夜間": "vms_year.tsv", "回収時": "vms_year_kaishu.tsv"}

lines = io.open(SRC, encoding="utf-8").read().splitlines()
header, body = lines[0], [l for l in lines[1:] if l.strip()]
for measure, name in OUT.items():
    rows = [l for l in body if l.split("\t")[0] == measure]
    with io.open(os.path.join(HERE, name), "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join([header] + rows) + "\n")
    print(f"{name}: {len(rows)}行")

recs = [l.split("\t") for l in body]
night = defaultdict(int)
for p in recs:
    if p[0] == "夜間":
        night[(p[3], p[2])] += int(p[6])
visits = sorted({p[2] for p in recs if p[0] == "回収時"})
start = min(p[2] for p in recs)
bad = 0
for p in recs:
    if p[0] != "回収時":
        continue
    prev = max([v for v in visits if v < p[2]], default=start)
    want = sum(t for (code, d), t in night.items() if code == p[3] and prev <= d < p[2])
    if want != int(p[6]):
        bad += 1
        print(f"  回収時 {p[2]} 号機{p[3]}: {int(p[6]):,} ≠ 夜間の合計 {want:,}")

month = defaultdict(int)
for p in recs:
    month[(p[1], p[3], p[0])] += int(p[6])
print("\n年月     号機        夜間      回収時        差")
for ym, code in sorted({(p[1], p[3]) for p in recs}, key=lambda k: (k[0], int(k[1]))):
    a, b = month[(ym, code, "夜間")], month[(ym, code, "回収時")]
    print(f"{ym}  {code:>4} {a:>11,} {b:>11,} {b - a:>+9,}")
if bad:
    raise SystemExit(f"回収時の {bad} 行が夜間の合計と合いません")
