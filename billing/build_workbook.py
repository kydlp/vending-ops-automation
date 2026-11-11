# -*- coding: utf-8 -*-
"""sales-volumeブックの再構成プロトタイプを生成する。
元ファイル: 【2026.06】sales-volume.xlsx（シート1=手入力／シート2=VMS貼付／オーナー別11枚）
方針: VMS を唯一の正とし、マスタ（号機・単価・契約）を外出し、精算は1枚に統合。
"""
import io, json, sys, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.shapes import GraphicalProperties

# dataviz skill / reference palette（light）— 検証済み: 8スロットとも PASS
# node scripts/validate_palette.js "#2a78d6,#eb6834,#1baf7a,#eda100,#e87ba4,#008300,#4a3aa7,#e34948" --mode light
SERIES = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "4A3AA7", "E34948"]
SURFACE = "FCFCFB"      # チャート面（積み上げの隙間に使う）
GRID = "E1E0D9"         # ヘアライングリッド
AXIS = "C3C2B7"         # 軸
MUTED = "898781"        # 目盛ラベル


def _axis_chrome(ch):
    """グリッドと軸をヘアラインの実線にする（破線は使わない）。"""
    for ax in (ch.x_axis, ch.y_axis):
        ax.majorGridlines = None
        ax.spPr = GraphicalProperties(ln=LineProperties(solidFill=AXIS, w=9525))
        ax.txPr = None
        ax.delete = False
    ch.y_axis.majorGridlines = openpyxl.chart.axis.ChartLines(
        spPr=GraphicalProperties(ln=LineProperties(solidFill=GRID, w=9525)))


def _paint(ser, hex_color, gap=True):
    """系列を塗る。積み上げの境目は枠線ではなく“面と同じ色の細線”で隙間を作る。"""
    ser.graphicalProperties = GraphicalProperties(solidFill=hex_color)
    ser.graphicalProperties.ln = LineProperties(
        solidFill=SURFACE, w=19050) if gap else LineProperties(noFill=True)

FONT = "Meiryo"
BLUE = Font(name=FONT, size=10, color="0000FF")          # 手入力
BLACK = Font(name=FONT, size=10)                          # 数式
BOLD = Font(name=FONT, size=10, bold=True)
TITLE = Font(name=FONT, size=13, bold=True)
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(name=FONT, size=10, bold=True, color="FFFFFF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")       # 編集してよいセル
NOTE_FONT = Font(name=FONT, size=9, color="808080")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
YEN = '¥#,##0;(¥#,##0);-'
NUM = '#,##0;(#,##0);-'
PCT = '0.0%'

MACHINES = [
    # 号機ID, VMS名, 旧スプシ名, 決済端末シリアル, 実際の設置場所, 満タン本数（公開用の架空値）
    (1, "ClientA SiteA店前",            "ClientA",   "SN0000001", "ClientA SiteA店", 60),
    (2, "SiteB モール",         "SiteB",   "SN0000002", "SiteB モール", 60),
    (3, "SiteC モール",   "SiteC",       "SN0000003", "SiteC モール（ClientB SiteC店）", 60),
    (4, "ClientKSiteD店",     "SiteD",     "SN0000004", "ClientKSiteD店", 60),
    (5, "ClientE SiteE店",         "SiteE",         "SN0000005", "SiteE モール", 60),
    (6, "SiteF モール2階",  "SiteF",     "SN0000006", "SiteF モール", 60),
    (7, "ClientG",           "ClientG",   "SN0000007", "SiteG モール（※旧スプシ通称との対応は要確認）", 60),
    (8, "ClientKSiteH店", "SiteH",       "SN0000008", "ClientKSiteH店", 60),
]
WHOLESALE = "WholesaleX(卸)"
WHOLESALE_ID = "W1"
# 突合キーは号機ID（＝VMS の号機コード）。号機名が変わっても数式が壊れないようにする。
MID = dict((m[1], m[0]) for m in MACHINES)
MID[WHOLESALE] = WHOLESALE_ID
MNAME = dict((m[0], m[1]) for m in MACHINES)

# 支払先, 対象ID, 方式, 料率, 日額, 月額固定, 最低保証, 税方式, 備考
#   対象IDは号機ID（＝VMS の号機コード）。号機名で結ぶと改名で壊れるため。
#   1=ClientA SiteA 2=SiteB モール 3=SiteC 4=ClientKSiteD
#   5=ClientESiteE 6=SiteF 7=ClientG 8=ClientKSiteH W1=WholesaleX(卸)
#   方式  : 料率／現金渡し
#   日額  : 稼働1日あたりの費用。稼働日数は M_カレンダー（定休曜日を除く）＋入力_月次調整
#   最低保証: 料率方式で、この額を下回ったらこの額を払う（ClientG＝売上7.5万円未満は固定6,000円）
#   税方式: 税込   … 賃料 = 税込売上 × 料率（料率の結果がそのまま税込）
#           税抜換算 … 賃料 = 税込売上 ÷ 1.08（軽減税率で税抜化）× 料率 × 1.1（賃料に課税）
CONTRACTS = [
    ("ClientA",                     1,  "料率", 0.12, 0, 0, 0, "税込", ""),
    ("ClientB株式会社", 3,  "料率", 0.10, 0, 0, 0, "税込", ""),
    ("ClientC株式会社",                       3,  "料率", 0.05, 0, 0, 0, "税込", ""),
    ("ClientC株式会社",           WHOLESALE_ID,  "料率", 0.05, 0, 0, 0, "税込", "卸値500円/本"),
    ("ClientD株式会社",   2,  "料率", 0.14, 0, 0, 0, "税込", ""),
    ("ClientE",                            5,  "料率", 0.11, 0, 0, 0, "税込", ""),
    ("ClientF株式会社", 6,  "料率", 0.07, 0, 0, 0, "税抜換算",
     "売上→÷1.08→×7%→×1.1。合計値で切り上げ"),
    ("ClientG",                         7,  "料率", 0.08, 0, 0, 6000, "税込",
     "売上7.5万円未満は固定6,000円／7.5万円以上は8%（＝最低保証6,000円と等価）"),
    ("ClientH株式会社",                     3,  "現金渡し", 0, 700, 0, 0, "税込",
     "日額は稼働日ごと（定休曜日以外）。イレギュラーは 入力_月次調整 で加減"),
    ("ClientI株式会社",                   8,  "現金渡し", 0, 400, 6000, 0, "税込",
     "月額=固定費。臨時の控除は 入力_月次調整 へ"),
    ("ClientJ株式会社",                   5,  "現金渡し", 0, 0, 0, 0, "税込", ""),
    ("ClientJ株式会社",                   6,  "現金渡し", 0, 0, 0, 0, "税込", ""),
    ("ClientJ株式会社",                   7,  "現金渡し", 0, 0, 0, 0, "税込", ""),
]
# 除外した契約（点検で確認した結果）
#   ClientC株式会社×ClientB  … 料率契約の対象外
#   ClientKSiteD店 … 請求書は別途作成しているため本ブックでは扱わない

# 支払先 → (シート名に使う短縮名, 正式名称)
PAYEE_SHEET = {
    "ClientA":                     "ClientA",
    "ClientB株式会社": "ClientB",
    "ClientC株式会社":                       "ClientC",
    "ClientD株式会社":   "ClientD",
    "ClientE":                            "ClientE",
    "ClientF株式会社": "ClientF",
    "ClientG":                         "ClientG",
    "ClientH株式会社":                     "ClientH",
    "ClientI株式会社":                   "ClientI",
    "ClientJ株式会社":                   "ClientJ",
    "ClientKSiteD店":       "ClientKSiteD",
}

# 請求書番号の支払先コード。並び順に依存させない（支払先が増えても既存番号が動かない）
PAYEE_CODE = {
    "ClientA":                     "01",
    "ClientB株式会社": "02",
    "ClientC株式会社":                       "03",
    "ClientD株式会社":   "04",
    "ClientE":                            "05",
    "ClientF株式会社": "06",
    "ClientG":                         "07",
    "ClientH株式会社":                     "08",
    "ClientI株式会社":                   "09",
    "ClientJ株式会社":                   "10",
    "ClientKSiteD店":       "11",
}

OTHER_SALES = [
    ("2026-06", "卸",       WHOLESALE,      150, 500, "旧ファイルは行ズレで100,000円計上→正しくは75,000円"),
    ("2026-06", "EC",       "EC-A",          None, None, ""),
    ("2026-06", "EC",       "EC-B",          None, None, ""),
    ("2026-06", "EC",       "EC-C",          None, None, ""),
    ("2026-06", "催事",     "催事A",         None, None, ""),
    ("2026-06", "催事",     "催事B",         None, None, ""),
    ("2026-06", "卸",       "卸B",           None, None, ""),
]
# 金額は公開用の架空値
OTHER_AMOUNT = {"EC-A":120000, "EC-B":300000, "EC-C":500000,
                "催事A":200000, "催事B":100000, "卸B":80000}


def style_header(ws, row, ncols, widths=None):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.font = HEAD_FONT; cell.fill = HEAD_FILL; cell.border = BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def load_year(path="vms_year.tsv"):
    """export-sales の出力を読む。列: 検量/年月/日付/号機コード/号機名/本数/総額/現金/カード。"""
    rows = []
    for i, line in enumerate(io.open(path, encoding="utf-8")):
        if i == 0 or not line.strip():
            continue
        p = line.rstrip('\n').split('\t')
        rows.append(dict(ym=p[1], date=p[2], code=int(p[3]), machine=p[4],
                         qty=float(p[5]), total=float(p[6]),
                         cash=float(p[7]), card=float(p[8])))
    return rows


WEEKDAYS = ["月", "火", "水", "木", "金", "土", "日"]
CLOSED_WEEKDAYS = (6,)    # 定休曜日（0=月 … 6=日）。公開用の架空値


def write_import(wb, title, rows, subtitle, note):
    """取込タブを1枚書き出して、最終行番号を返す。"""
    ws = wb.create_sheet(title)
    ws["A1"] = subtitle; ws["A1"].font = TITLE
    ws["A2"] = note; ws["A2"].font = NOTE_FONT
    ws["A3"] = "取得日: 2026-08-01 ／ 取得元: VMS 売上照会 ／ 手で書き換えないこと"
    ws["A3"].font = NOTE_FONT
    heads = ["年月", "日付", "号機ID", "号機（VMS名）", "本数", "現金", "カード", "合計", "曜日"]
    for i, h in enumerate(heads, 1): ws.cell(4, i, h)
    style_header(ws, 4, len(heads), [10, 12, 8, 24, 10, 14, 14, 14, 8])
    for j, d in enumerate(rows):
        r = 5 + j
        for i, v in enumerate([d["ym"], d["date"], d["code"], d["machine"], d["qty"], d["cash"], d["card"]], 1):
            c = ws.cell(r, i, v); c.font = BLACK; c.border = BOX
        for i in (6, 7): ws.cell(r, i).number_format = YEN
        ws.cell(r, 5).number_format = NUM
        c = ws.cell(r, 8, f"=F{r}+G{r}"); c.font = BLACK; c.number_format = YEN; c.border = BOX
        c = ws.cell(r, 9, d["wd"]); c.font = BLACK; c.border = BOX
    return 4 + len(rows)


NM = '=IFERROR(INDEX(M_号機!$B:$B,MATCH({},M_号機!$A:$A,0)),"")'



def read_contracts(wb):
    """ブックの M_契約 から契約を読む。Python 定数 CONTRACTS ではなくこちらが運用中の正。

    支払先を増やすときは M_契約 に行を足して予備の請求書シートに支払先名を入れるだけでよく、
    ブックを作り直さずに済む（作り直すと手入力が消えるため）。
    列: B支払先 C対象ID E方式 F料率 G日額 H月額固定 I最低保証 J税方式 K備考
    """
    ws = wb["M_契約"]
    out = []
    for r in range(5, ws.max_row + 1):
        payee = ws.cell(r, 2).value
        tid = ws.cell(r, 3).value
        if not payee or tid in (None, ""):
            continue
        out.append((payee, tid, ws.cell(r, 5).value or "",
                    ws.cell(r, 6).value or 0, ws.cell(r, 7).value or 0,
                    ws.cell(r, 8).value or 0, ws.cell(r, 9).value or 0,
                    ws.cell(r, 10).value or "税込", ws.cell(r, 11).value or ""))
    return out


def invoice_sheets(wb):
    """請求書シート名 → 支払先名。支払先は J1 セルに入っている（予備シートは空欄）。"""
    out = {}
    for name in wb.sheetnames:
        if not name.startswith("請求_"):
            continue
        v = wb[name]["J1"].value
        if v:
            out[name] = v
    return out


def write_evidence(ws, dr, target_ids, target_ym, detail):
    """算定根拠（回収日ごとの明細＋号機ごとの小計）を dr 行目から書き、最終行を返す。

    回収回数は号機・月ごとに変わるので行数を数式で組めない。生成時に実データを流し込み、
    月次スクリプトが再出力のたびに書き直す（monthly_run.py draft）。
    """
    ws.cell(dr, 1, f"【算定根拠】VMS 回収時検量　{target_ym}").font = BOLD
    ws.cell(dr + 1, 1,
            f'=IF(集計_号機別!$B$3<>"{target_ym}",'
            f'"⚠ この明細は {target_ym} 時点の出力です。対象年月を変えた場合は月次スクリプトで再出力してください","")'
            ).font = Font(name=FONT, size=9, bold=True, color="C00000")
    dr += 3
    for tid in target_ids:
        recs = detail.get((target_ym, tid), [])
        ws.cell(dr, 1, f"■ {MNAME[tid]}（号機ID {tid}）").font = BOLD
        dr += 1
        for i, h in enumerate(["回収日", "本数", "現金", "カード", "合計"], 1): ws.cell(dr, i, h)
        style_header(ws, dr, 5)
        dr += 1
        st = dr
        for d in recs:
            vals = [d["date"], d["qty"], d["cash"], d["card"], d["cash"] + d["card"]]
            for i, v in enumerate(vals, 1):
                cc = ws.cell(dr, i, v); cc.font = BLACK; cc.border = BOX
                cc.number_format = NUM if i == 2 else (YEN if i >= 3 else "@")
            dr += 1
        ws.cell(dr, 1, "小計").font = BOLD; ws.cell(dr, 1).border = BOX
        for i in (2, 3, 4, 5):
            L = get_column_letter(i)
            cc = ws.cell(dr, i, f"=SUM({L}{st}:{L}{dr-1})" if recs else 0)
            cc.font = BOLD; cc.number_format = NUM if i == 2 else YEN; cc.border = BOX
        if not recs:
            ws.cell(dr - 1, 1, "（この月の回収実績はありません）").font = NOTE_FONT
        dr += 2
    return dr


def main():
    daily = load_year("vms_year.tsv")              # 夜間検量＝分析の正
    kaishu = load_year("vms_year_kaishu.tsv")      # 回収時検量＝精算の正
    months = sorted({d["ym"] for d in daily} | {d["ym"] for d in kaishu})
    # 請求書の算定根拠は「出力時点の対象月」で書き込む（行数が月ごとに変わるため数式にできない）
    TARGET_YM = sys.argv[1] if len(sys.argv) > 1 else months[-1]
    detail = {}
    for d in kaishu:
        detail.setdefault((d["ym"], d["code"]), []).append(d)
    for v in detail.values():
        v.sort(key=lambda x: x["date"])
    for rows in (daily, kaishu):
        for d in rows:
            y, m, dd = (int(x) for x in d["date"].split("-"))
            d["wd"] = WEEKDAYS[datetime.date(y, m, dd).weekday()]
    # 号機の並び順は「年計の大きい順」で固定する（月を切り替えても並びと色が動かないように）
    year_total = {}
    for d in daily:
        year_total[d["code"]] = year_total.get(d["code"], 0) + d["total"]
    ORDER = [m[0] for m in sorted(MACHINES, key=lambda m: -year_total.get(m[0], 0))]
    COLOR = {mid: SERIES[i] for i, mid in enumerate(ORDER)}   # 色は号機IDに固定
    wb = openpyxl.Workbook()

    # ---------------- README ----------------
    ws = wb.active; ws.title = "README"
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 108
    lines = [
        (TITLE, "OurProduct 販売数・支払精算ブック（再構成プロトタイプ）"),
        (NOTE_FONT, "元ファイル: 【2026.06】sales-volume.xlsx / 作成: 2026-07-31"),
        (None, ""),
        (BOLD, "■ 何が変わったか"),
        (BLACK, "1. 販売数・現金・カードの正は VMS だけ。手入力の二重管理をやめた（旧シート1の I〜AZ 列は廃止）。"),
        (BLACK, "2. 単価・契約条件をマスタに外出しした。値上げや料率変更は、数式ではなくマスタの1行を直すだけで済む。"),
        (BLACK, "3. 支払先ごとの11枚を1枚に統合した（精算_支払明細）。上のプルダウンで支払先を切り替える。"),
        (BLACK, "4. 全部の数式が SUMIFS 参照になった。旧ファイルのようなセル直参照（='シート1'!AE3）をやめたので、行を足しても壊れない。"),
        (None, ""),
        (BOLD, "■ 入力するのはこの2枚だけ（黄色のセル）"),
        (BLACK, "・取込_回収時 … 現金を回収した時点の実績。【精算・請求はこちら】"),
        (BLACK, "・取込_夜間 … 毎晩の確定値。【曜日別・売切・トレンド分析はこちら】精算には使わない"),
        (BLACK, "・入力_その他売上 … EC・催事・卸など、VMS に載らない売上を1行1件で入力する。"),
        (None, ""),
        (BOLD, "■ 条件が変わったときに直す場所（黄色のセル）"),
        (BLACK, "・M_単価 … 値上げ時は行を追加する（適用開始日つき）。過去の単価は消さない。"),
        (BLACK, "・M_契約 … 料率・日額・月額固定費の変更はここ。支払先を増やすときも行追加のみ。"),
        (BLACK, "・M_号機 … 自販機の増設・撤去・名称変更はここ。"),
        (None, ""),
        (BOLD, "■ 色の意味"),
        (BLUE, "青字＋黄色セル … 人が入力する値"),
        (BLACK, "黒字 … 数式。原則さわらない"),
        (None, ""),
        (BOLD, "■ 必ず先に読むこと"),
        (BLACK, "点検 シートに、旧ファイルから引き継いだ不明点と、確認が必要な契約条件をまとめてある。"),
        (BLACK, "特に「ClientKSiteD店」の契約条件は旧ファイルが壊れていて復元できなかった。"),
    ]
    r = 1
    for f, text in lines:
        c = ws.cell(r, 2, text)
        if f: c.font = f
        r += 1

    # ---------------- M_号機 ----------------
    ws = wb.create_sheet("M_号機")
    ws["A1"] = "自販機マスタ"; ws["A1"].font = TITLE
    ws["A2"] = ("※ 1台が最大4通りの名前で呼ばれている（VMS機名／旧スプシ通称／決済端末の登録名／実際の設置場所）。"
                "突合は決済端末のシリアルで行うこと。")
    ws["A2"].font = NOTE_FONT
    heads = ["号機ID", "VMS号機名（正）", "旧スプシ通称", "決済端末シリアル",
             "実際の設置場所", "満タン本数", "状態"]
    for i, h in enumerate(heads, 1): ws.cell(4, i, h)
    style_header(ws, 4, len(heads), [8, 24, 16, 18, 34, 12, 10])
    for j, (mid, name, old, serial, place, cap) in enumerate(MACHINES):
        r = 5 + j
        for i, v in enumerate([mid, name, old, serial, place, cap, "稼働中"], 1):
            c = ws.cell(r, i, v); c.font = BLUE; c.fill = INPUT_FILL; c.border = BOX

    # ---------------- M_単価 ----------------
    ws = wb.create_sheet("M_単価")
    ws["A1"] = "単価マスタ"; ws["A1"].font = TITLE
    ws["A2"] = "※ 自販機の売上金額は VMS 実績をそのまま使うため、この表は卸・EC など手入力売上の換算と、履歴の記録用。"
    ws["A2"].font = NOTE_FONT
    heads = ["適用開始日", "対象", "商品", "単価", "備考"]
    for i, h in enumerate(heads, 1): ws.cell(4, i, h)
    style_header(ws, 4, len(heads), [14, 22, 20, 12, 40])
    # 単価は公開用の架空値
    prices = [("2026-04-01", "自販機（全機）", "OurProduct", 800, "値上げ前"),
              ("2026-06-01", "自販機（全機）", "OurProduct", 850, "2026-06-01 値上げ"),
              ("2026-04-01", "卸", "OurProduct", 500, "WholesaleX 卸値")]
    for j, row in enumerate(prices):
        r = 5 + j
        for i, v in enumerate(row, 1):
            c = ws.cell(r, i, v); c.font = BLUE; c.fill = INPUT_FILL; c.border = BOX
            if i == 4: c.number_format = YEN

    # ---------------- M_カレンダー ----------------
    # 日額は「稼働日ごと」（点検で確認）。定休曜日（CLOSED_WEEKDAYS）以外が稼働日なので、月ごとの日数をここで持つ。
    # 対象年月のプルダウンもこの表を参照する（旧ファイルで新しい月を選べなかった問題の解消）。
    ws = wb.create_sheet("M_カレンダー")
    ws["A1"] = "稼働日数カレンダー"; ws["A1"].font = TITLE
    ws["A2"] = ("※ 稼働日＝定休曜日以外の日。イレギュラーな稼働・休止は "
                "入力_月次調整 で日数を加減する（この表は直さない）。")
    ws["A2"].font = NOTE_FONT
    heads = ["年月", "暦日数", "稼働日数（定休曜日を除く）"]
    for i, h in enumerate(heads, 1): ws.cell(4, i, h)
    style_header(ws, 4, len(heads), [12, 10, 22])
    y0, m0 = (int(x) for x in months[0].split("-"))
    ye, me = (int(x) for x in months[-1].split("-"))
    cal_months = []
    yy, mm = y0, m0
    for _ in range((ye - y0) * 12 + (me - m0) + 1 + 24):    # データ期間＋24ヶ月先まで選べるようにする
        d1 = datetime.date(yy, mm, 1)
        d2 = datetime.date(yy + (mm == 12), (mm % 12) + 1, 1)
        ndays = (d2 - d1).days
        work = sum(1 for k in range(ndays)
                   if (d1 + datetime.timedelta(days=k)).weekday() not in CLOSED_WEEKDAYS)
        cal_months.append(("%04d-%02d" % (yy, mm), ndays, work))
        yy, mm = yy + (mm == 12), (mm % 12) + 1
    for j, (ym, nd, wk) in enumerate(cal_months):
        r = 5 + j
        for i, v in enumerate([ym, nd, wk], 1):
            c = ws.cell(r, i, v); c.font = BLACK; c.border = BOX
        ws.cell(r, 1).number_format = "@"
    cal_last = 4 + len(cal_months)

    # ---------------- M_契約 ----------------
    ws = wb.create_sheet("M_契約")
    ws["A1"] = "契約マスタ（支払条件）"; ws["A1"].font = TITLE
    ws["A2"] = "※ 1台の自販機に複数の支払先がぶら下がる（SiteCは3社、SiteF・SiteG モールは2社）。1行＝1つの支払条件。"
    ws["A2"].font = NOTE_FONT
    heads = ["キー（自動）", "支払先", "対象ID", "対象（表示）", "方式", "料率", "日額",
             "月額固定", "最低保証", "税方式", "備考"]
    for i, h in enumerate(heads, 1): ws.cell(4, i, h)
    style_header(ws, 4, len(heads), [30, 32, 9, 24, 12, 9, 10, 12, 12, 12, 46])
    for j, (payee, target, method, rate, daily_fee, monthly, floor_, taxmode, note) in enumerate(CONTRACTS):
        r = 5 + j
        ws.cell(r, 1, f"=B{r}&\"|\"&C{r}").font = BLACK
        # 対象（表示）は号機IDから引く。VMS側の改名に自動追従させるため。
        disp = WHOLESALE if target == WHOLESALE_ID else NM.format("$C" + str(r))
        vals = [payee, target, disp, method, rate, daily_fee, monthly, floor_, taxmode, note]
        for i, v in enumerate(vals, 2):
            c = ws.cell(r, i, v); c.font = BLUE; c.fill = INPUT_FILL; c.border = BOX
        ws.cell(r, 4).font = BLACK; ws.cell(r, 4).fill = PatternFill()
        ws.cell(r, 6).number_format = PCT
        for i in (7, 8, 9): ws.cell(r, i).number_format = YEN
        ws.cell(r, 1).border = BOX
    n_contract = 4 + len(CONTRACTS)
    ws.cell(n_contract + 2, 2, "方式の意味").font = BOLD
    ws.cell(n_contract + 3, 2, "料率　　… 支払額 = 対象売上 × 料率（税方式で計算式が変わる）。最低保証があればそれを下回らない").font = BLACK
    ws.cell(n_contract + 4, 2, "現金渡し… 請求額 = 回収現金 − 日額×稼働日数 − 月額固定 − 金額調整（キャッシュレス分は当社が受領）").font = BLACK
    ws.cell(n_contract + 5, 2, "要確認　… 条件が未確定。精算シートでは 0 円として扱い、警告を出す").font = BLACK
    ws.cell(n_contract + 6, 2, "端数　　… すべて合計値で切り上げ（ROUNDUP）").font = BLACK
    # 支払先のユニークリスト（プルダウン用）
    payees = sorted(set(c[0] for c in CONTRACTS))
    ws.cell(4, 11, "支払先一覧（プルダウン用）").font = BOLD
    ws.cell(4, 12, "支払先コード（固定）").font = BOLD
    ws.column_dimensions['K'].width = 34
    ws.column_dimensions['L'].width = 18
    for j, p in enumerate(payees):
        ws.cell(5 + j, 11, p).font = BLACK
        # 請求書番号に使う。支払先が増減しても既存の番号がずれないよう、並び順ではなく固定値で持つ
        c = ws.cell(5 + j, 12, PAYEE_CODE[p]); c.font = BLUE; c.fill = INPUT_FILL; c.border = BOX
    ws.cell(6 + len(payees), 11, "※ コードは一度振ったら変えないこと（請求書番号の一意性が壊れる）").font = NOTE_FONT
    payee_code_rng = f"M_契約!$L$5:$L${4+len(payees)}"
    payee_name_rng = f"M_契約!$K$5:$K${4+len(payees)}"

    # ---------------- 取込（回収時／夜間の2本立て） ----------------
    last_k = write_import(
        wb, "取込_回収時", kaishu, "VMS 回収時検量（精算・請求の正）",
        f"※ 現金を回収した時点の実績（{months[0]}〜{months[-1]}・{len(kaishu):,}行）。"
        "オーナー精算と請求はこちらを使う。回収日にまとめて計上されるため日別の需要は分からない。")
    last_in = write_import(
        wb, "取込_夜間", daily, "VMS 夜間検量（分析の正）",
        f"※ 毎晩の確定値（{months[0]}〜{months[-1]}・{len(daily):,}行）。"
        "曜日別需要・売切検知・トレンド分析はこちらを使う。精算には使わない。")

    # ---------------- 入力_その他売上 ----------------
    ws = wb.create_sheet("入力_その他売上")
    ws["A1"] = "その他売上（VMS に載らないもの）"; ws["A1"].font = TITLE
    ws["A2"] = "※ EC・催事・卸など。数量と単価を入れれば金額は自動計算。金額しか分からない場合は F 列に直接入れてよい。"
    ws["A2"].font = NOTE_FONT
    heads = ["年月", "区分", "名称", "数量", "単価", "金額", "備考"]
    for i, h in enumerate(heads, 1): ws.cell(4, i, h)
    style_header(ws, 4, len(heads), [10, 10, 22, 10, 10, 16, 52])
    for j, (ym, kind, name, qty, price, note) in enumerate(OTHER_SALES):
        r = 5 + j
        for i, v in enumerate([ym, kind, name, qty, price], 1):
            c = ws.cell(r, i, v); c.font = BLUE; c.fill = INPUT_FILL; c.border = BOX
        ws.cell(r, 5).number_format = YEN
        if qty and price:
            c = ws.cell(r, 6, f"=D{r}*E{r}"); c.font = BLACK
        else:
            c = ws.cell(r, 6, OTHER_AMOUNT.get(name)); c.font = BLUE; c.fill = INPUT_FILL
        c.number_format = YEN; c.border = BOX
        c = ws.cell(r, 7, note); c.font = NOTE_FONT; c.border = BOX
    last_other = 4 + len(OTHER_SALES)

    # ---------------- 入力_月次調整 ----------------
    ws = wb.create_sheet("入力_月次調整")
    ws["A1"] = "月次の調整（稼働日数・金額）"; ws["A1"].font = TITLE
    ws["A2"] = ("※ イレギュラーが起きた月だけ1行足す。無ければ空のままでよい。"
                "　稼働日数の調整＝定休曜日に稼働した日は +1、休んだ日は −1。"
                "　金額の調整＝控除として差し引く額。")
    ws["A2"].font = NOTE_FONT
    heads = ["年月", "支払先", "対象ID", "稼働日数の調整（日）", "金額の調整（円・控除）", "備考"]
    for i, h in enumerate(heads, 1): ws.cell(4, i, h)
    style_header(ws, 4, len(heads), [12, 34, 10, 20, 22, 44])
    for r in range(5, 25):                      # 入力用の空行
        for i in range(1, 7):
            c = ws.cell(r, i); c.font = BLUE; c.fill = INPUT_FILL; c.border = BOX
        ws.cell(r, 1).number_format = "@"
        ws.cell(r, 5).number_format = YEN
    ws.cell(26, 1, "※ 支払先と対象IDは M_契約 の表記に合わせること（対象IDは号機ID）。").font = NOTE_FONT

    # ---------------- 集計_号機別 ----------------
    ws = wb.create_sheet("集計_号機別")
    ws["A1"] = "号機別 月次集計"; ws["A1"].font = TITLE
    ws["A3"] = "対象年月"; ws["A3"].font = BOLD
    c = ws["B3"]; c.value = TARGET_YM; c.font = BLUE; c.fill = INPUT_FILL; c.border = BOX
    c.number_format = "@"   # 「2026-07」を日付に自動変換させない（SUMIFS の突合先は文字列）
    ws["C3"] = f"← ここを変えると下の表と精算シートが切り替わる（データがあるのは {months[0]}〜{months[-1]}）"
    ws["C3"].font = NOTE_FONT
    dv_m = DataValidation(type="list", formula1=f"M_カレンダー!$A$5:$A${cal_last}", allow_blank=False)
    ws.add_data_validation(dv_m); dv_m.add(ws["B3"])
    ws["A4"] = "※ この表は【回収時検量】＝現金を回収した時点の実績。精算・請求はここを使う。右端2列は夜間検量との差＝月末に機内へ残っている売上。"
    ws["A4"].font = NOTE_FONT
    heads = ["号機ID", "号機（VMS名）", "本数", "現金", "カード", "合計", "回収回数", "現金比率",
             "参考:夜間検量", "差（機内滞留）"]
    for i, h in enumerate(heads, 1): ws.cell(5, i, h)
    style_header(ws, 5, len(heads), [8, 24, 10, 14, 14, 14, 10, 10, 15, 15])
    for j, mid in enumerate(ORDER):
        r = 6 + j
        c = ws.cell(r, 1, mid); c.font = BLACK; c.border = BOX
        f = {
            2: NM.format("$A" + str(r)),
            3: f'=SUMIFS(取込_回収時!$E:$E,取込_回収時!$C:$C,$A{r},取込_回収時!$A:$A,$B$3)',
            4: f'=SUMIFS(取込_回収時!$F:$F,取込_回収時!$C:$C,$A{r},取込_回収時!$A:$A,$B$3)',
            5: f'=SUMIFS(取込_回収時!$G:$G,取込_回収時!$C:$C,$A{r},取込_回収時!$A:$A,$B$3)',
            6: f'=D{r}+E{r}',
            7: f'=COUNTIFS(取込_回収時!$C:$C,$A{r},取込_回収時!$A:$A,$B$3)',
            8: f'=IF(F{r}=0,0,D{r}/F{r})',
            9: f'=SUMIFS(取込_夜間!$H:$H,取込_夜間!$C:$C,$A{r},取込_夜間!$A:$A,$B$3)',
            10: f'=I{r}-F{r}',
        }
        for i, fx in f.items():
            cc = ws.cell(r, i, fx); cc.font = BLACK; cc.border = BOX
            cc.number_format = NUM if i in (3, 7) else (PCT if i == 8 else ("General" if i == 2 else YEN))
    tr = 6 + len(MACHINES)
    ws.cell(tr, 1, "合計").font = BOLD
    for i in (3, 4, 5, 6, 9, 10):
        cc = ws.cell(tr, i, f"=SUM({get_column_letter(i)}6:{get_column_letter(i)}{tr-1})")
        cc.font = BOLD; cc.number_format = NUM if i == 3 else YEN; cc.border = BOX
    ws.cell(tr, 1).border = BOX
    agg_last = tr - 1

    # グラフ①：号機別の現金／カード内訳（積み上げ横棒・売上の大きい順）
    ch = BarChart(); ch.type = "bar"; ch.grouping = "stacked"; ch.overlap = 100
    ch.title = "号機別 売上と決済内訳（現金／カード・単位:円）"
    ch.add_data(Reference(ws, min_col=4, max_col=5, min_row=5, max_row=agg_last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=6, max_row=agg_last))
    _paint(ch.series[0], SERIES[0]); _paint(ch.series[1], SERIES[1])
    ch.gapWidth = 70; ch.legend.position = "b"
    ch.width, ch.height = 22, 11
    _axis_chrome(ch)
    ws.add_chart(ch, "A17")

    # ---------------- 集計_月次推移 ----------------
    ws = wb.create_sheet("集計_月次推移")
    ws["A1"] = "号機別 月次推移（売上金額）"; ws["A1"].font = TITLE
    ws["A2"] = "※ 空欄でなく 0 の月＝その号機が稼働していなかった月。単価の変更履歴は M_単価 を参照。"
    ws["A2"].font = NOTE_FONT
    ws.cell(4, 1, "号機（VMS名）")
    for i, ym in enumerate(months): ws.cell(4, 2 + i, ym)
    ws.cell(4, 2 + len(months), "年計")
    style_header(ws, 4, 2 + len(months), [24] + [12] * len(months) + [14])
    for j, mid in enumerate(ORDER):
        r = 5 + j
        c = ws.cell(r, 1, NM.format(mid)); c.font = BLACK; c.border = BOX
        for i, ym in enumerate(months):
            cc = ws.cell(r, 2 + i,
                         f'=SUMIFS(取込_夜間!$H:$H,取込_夜間!$C:$C,{mid},'
                         f'取込_夜間!$A:$A,{get_column_letter(2+i)}$4)')
            cc.font = BLACK; cc.number_format = YEN; cc.border = BOX
        cc = ws.cell(r, 2 + len(months),
                     f"=SUM(B{r}:{get_column_letter(1+len(months))}{r})")
        cc.font = BOLD; cc.number_format = YEN; cc.border = BOX
    tr2 = 5 + len(MACHINES)
    ws.cell(tr2, 1, "合計").font = BOLD; ws.cell(tr2, 1).border = BOX
    for i in range(2, 3 + len(months)):
        L = get_column_letter(i)
        cc = ws.cell(tr2, i, f"=SUM({L}5:{L}{tr2-1})")
        cc.font = BOLD; cc.number_format = YEN; cc.border = BOX
    last_m_col = 1 + len(months)

    # グラフ②：全機合計の月次推移（単系列＝凡例なし。季節性を1本で見せる）
    ch = LineChart()
    ch.title = "全機合計の月次推移（単位:円）"
    ch.add_data(Reference(ws, min_col=1, max_col=last_m_col, min_row=tr2, max_row=tr2), from_rows=True, titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, max_col=last_m_col, min_row=4))
    s = ch.series[0]
    s.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=SERIES[0], w=25400))
    s.marker = Marker(symbol="circle", size=7)
    s.smooth = False
    ch.legend = None
    ch.width, ch.height = 22, 9
    _axis_chrome(ch)
    ws.add_chart(ch, "A16")

    # グラフ③：号機別の積み上げ（構成の変化）
    ch = BarChart(); ch.type = "col"; ch.grouping = "stacked"; ch.overlap = 100
    ch.title = "号機別 月次推移（積み上げ・単位:円）"
    ch.add_data(Reference(ws, min_col=1, max_col=last_m_col, min_row=5, max_row=tr2 - 1),
                from_rows=True, titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, max_col=last_m_col, min_row=4))
    for i, s in enumerate(ch.series):
        _paint(s, COLOR[ORDER[i]])
    ch.gapWidth = 60; ch.legend.position = "b"
    ch.width, ch.height = 22, 12
    _axis_chrome(ch)
    ws.add_chart(ch, "A36")

    # ---------------- 集計_曜日別 ----------------
    ws = wb.create_sheet("集計_曜日別")
    ws["A1"] = "曜日別の需要（1日あたり平均本数・直近1年）"; ws["A1"].font = TITLE
    ws["A2"] = ("※ 販売実績のあった日だけを分母にした平均。")
    ws["A2"].font = NOTE_FONT
    ws.cell(4, 1, "曜日")
    for i, mid in enumerate(ORDER): ws.cell(4, 2 + i, NM.format(mid))
    ws.cell(4, 2 + len(ORDER), "全機合計")
    style_header(ws, 4, 2 + len(ORDER), [8] + [17] * len(ORDER) + [12])
    for j, wd in enumerate(WEEKDAYS):
        r = 5 + j
        c = ws.cell(r, 1, wd); c.font = BLACK; c.border = BOX
        for i, mid in enumerate(ORDER):
            cc = ws.cell(r, 2 + i,
                         f'=IFERROR(SUMIFS(取込_夜間!$E:$E,取込_夜間!$I:$I,$A{r},'
                         f'取込_夜間!$C:$C,{mid})/'
                         f'COUNTIFS(取込_夜間!$I:$I,$A{r},取込_夜間!$C:$C,{mid}),0)')
            cc.font = BLACK; cc.number_format = "0.0"; cc.border = BOX
        cc = ws.cell(r, 2 + len(ORDER),
                     f"=SUM(B{r}:{get_column_letter(1+len(ORDER))}{r})")
        cc.font = BOLD; cc.number_format = "0.0"; cc.border = BOX

    ch = BarChart(); ch.type = "col"
    ch.title = "曜日別 1日あたり販売本数（全機合計）"
    ch.add_data(Reference(ws, min_col=2 + len(ORDER), min_row=4, max_row=11), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=11))
    _paint(ch.series[0], SERIES[0], gap=False)
    ch.gapWidth = 80; ch.legend = None
    ch.width, ch.height = 20, 10
    _axis_chrome(ch)
    ws.add_chart(ch, "A14")
    ws.cell(13, 1, "※ 上の表がそのままグラフの数値（表とグラフは同じ数字）。号機ごとの差を見るときは表を横に読む。").font = NOTE_FONT

    # ---------------- 精算_支払明細 ----------------
    ws = wb.create_sheet("精算_支払明細")
    ws["A1"] = "支払先別 精算明細"; ws["A1"].font = TITLE
    ws["A3"] = "支払先"; ws["A3"].font = BOLD
    c = ws["B3"]; c.value = payees[0]; c.font = BLUE; c.fill = INPUT_FILL; c.border = BOX
    ws["D3"] = "対象年月"; ws["D3"].font = BOLD
    c = ws["E3"]; c.value = "=集計_号機別!$B$3"; c.font = BLACK; c.border = BOX
    ws["F3"] = "← 年月の変更は 集計_号機別 で行う"; ws["F3"].font = NOTE_FONT
    dv = DataValidation(type="list", formula1=f"M_契約!$K$5:$K${4+len(payees)}", allow_blank=False)
    ws.add_data_validation(dv); dv.add(ws["B3"])

    heads = ["対象", "方式", "売上", "現金", "稼働日数", "料率", "日額×稼働日数", "月額固定",
             "金額調整", "支払額／請求額"]
    for i, h in enumerate(heads, 1): ws.cell(5, i, h)
    style_header(ws, 5, len(heads), [24, 12, 14, 14, 10, 9, 14, 12, 12, 18])
    targets = [m[0] for m in MACHINES] + [WHOLESALE_ID]
    ck = "M_契約!$A:$A"
    for j, t in enumerate(targets):
        r = 6 + j
        disp = WHOLESALE if t == WHOLESALE_ID else NM.format(t)
        c = ws.cell(r, 1, disp); c.font = BLACK; c.border = BOX
        key = f'$B$3&"|"&"{t}"'
        ws.cell(r, 2, f'=IFERROR(INDEX(M_契約!$E:$E,MATCH({key},{ck},0)),"")')
        ws.cell(r, 3, f'=IF($B{r}="","",IF("{t}"="{WHOLESALE_ID}",'
                      f'SUMIFS(入力_その他売上!$F$5:$F${last_other},入力_その他売上!$C$5:$C${last_other},$A{r},'
                      f'入力_その他売上!$A$5:$A${last_other},$E$3),'
                      f'SUMIFS(集計_号機別!$F:$F,集計_号機別!$A:$A,"{t}")))')
        ws.cell(r, 4, f'=IF($B{r}="","",SUMIFS(集計_号機別!$D:$D,集計_号機別!$A:$A,"{t}"))')
        # 稼働日数 = カレンダーの基準日数（定休曜日を除く）＋ 入力_月次調整 の加減
        adj = (f'入力_月次調整!$A:$A,$E$3,入力_月次調整!$B:$B,$B$3,入力_月次調整!$C:$C,"{t}"')
        ws.cell(r, 5, f'=IF($B{r}="","",IFERROR(INDEX(M_カレンダー!$C:$C,MATCH($E$3,M_カレンダー!$A:$A,0)),0)'
                      f'+SUMIFS(入力_月次調整!$D:$D,{adj}))')
        ws.cell(r, 6, f'=IF($B{r}="","",IFERROR(INDEX(M_契約!$F:$F,MATCH({key},{ck},0)),0))')
        ws.cell(r, 7, f'=IF($B{r}="","",IFERROR(INDEX(M_契約!$G:$G,MATCH({key},{ck},0)),0)*$E{r})')
        ws.cell(r, 8, f'=IF($B{r}="","",IFERROR(INDEX(M_契約!$H:$H,MATCH({key},{ck},0)),0))')
        ws.cell(r, 9, f'=IF($B{r}="","",SUMIFS(入力_月次調整!$E:$E,{adj}))')
        # 料率  : 税方式で式が変わる → 最低保証を下回らせない → 金額調整を控除。端数は合計値で切り上げ
        # 現金渡し: 回収現金 − 日額×稼働日数 − 月額固定 − 金額調整
        tax = f'IFERROR(INDEX(M_契約!$J:$J,MATCH({key},{ck},0)),"税込")'
        flr = f'IFERROR(INDEX(M_契約!$I:$I,MATCH({key},{ck},0)),0)'
        base = (f'IF({tax}="税抜換算",ROUNDUP($C{r}/1.08*$F{r}*1.1,0),ROUNDUP($C{r}*$F{r},0))')
        ws.cell(r, 10, f'=IF($B{r}="","",IF($B{r}="料率",MAX({base},{flr})-$I{r},'
                       f'IF($B{r}="現金渡し",$D{r}-$G{r}-$H{r}-$I{r},0)))')
        for i in range(2, 11):
            cc = ws.cell(r, i); cc.font = BLACK; cc.border = BOX
            cc.number_format = PCT if i == 6 else (NUM if i == 5 else YEN)
    pr = 6 + len(targets)
    ws.cell(pr, 1, "合計").font = BOLD
    cc = ws.cell(pr, 10, f"=SUM(J6:J{pr-1})"); cc.font = BOLD; cc.number_format = YEN; cc.border = BOX
    ws.cell(pr, 1).border = BOX
    ws.cell(pr + 2, 1, '=IF(COUNTIFS(M_契約!$B:$B,$B$3,M_契約!$E:$E,"要確認")>0,'
            '"⚠ この支払先には条件未確定の契約が含まれる。点検シートを参照","")')
    ws.cell(pr + 2, 1).font = Font(name=FONT, size=10, bold=True, color="C00000")
    ws.cell(pr + 4, 1, "※ 空欄の行＝この支払先の契約対象ではない号機。").font = NOTE_FONT

    # ---------------- M_請求設定 ----------------
    ws = wb.create_sheet("M_請求設定")
    ws["A1"] = "請求書・支払明細書の共通設定"; ws["A1"].font = TITLE
    ws["A2"] = "※ ここを直すと全部の請求書シートに反映される。黄色は入力必須。"
    ws["A2"].font = NOTE_FONT
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 46
    ws.column_dimensions['C'].width = 44
    settings = [
        ("発行者名", "OurCompany", ""),
        ("郵便番号", "〒000-0000", ""),
        ("住所", "東京都○○区○○1-2-3", ""),
        ("電話番号", "000-0000-0000", ""),
        ("インボイス登録番号", "T0000000000000", "自社の登録番号。支払明細書（仕入明細書）では相手方の番号が必要・別途確認"),
        ("振込先", "○○銀行　○○支店（普通）0000000", ""),
        ("消費税の扱い", "賃料は消費税10%（税込表示）。売価は軽減税率8%の税込単価。",
         "SiteFのみ 税抜換算方式（M_契約 の税方式列を参照）"),
        ("請求書番号の付け方", "INV-{年月}-{支払先コード}", "例: INV-202607-10（末尾は支払先ごとに固定。同じ月に重複しない）"),
        ("支払期日", "翌月末日", ""),
        ("発行日", datetime.date(2026, 7, 5), "サンプルの発行日。空欄だと請求書に記入を促す注意書きが出る。過去分を再出力するときは、控えと合わせて実際の発行日を入れる"),
    ]
    for i, h in enumerate(["項目", "値", "備考"], 1): ws.cell(4, i, h)
    style_header(ws, 4, 3)
    for j, (k, v, note) in enumerate(settings):
        r = 5 + j
        c = ws.cell(r, 1, k); c.font = BOLD; c.border = BOX
        c = ws.cell(r, 2, v); c.font = BLUE; c.fill = INPUT_FILL; c.border = BOX
        if isinstance(v, datetime.date): c.number_format = "yyyy/mm/dd"
        c = ws.cell(r, 3, note)
        c.font = Font(name=FONT, size=9, color="C00000") if note.startswith("★") else NOTE_FONT
        c.border = BOX
    S = {k: f"M_請求設定!$B${5+j}" for j, (k, _, _) in enumerate(settings)}

    # ---------------- 請求書シート（支払先ごとに1枚＋予備） ----------------
    # 支払先と対象は数式にベタ書きせず、セル（J1／J列）に持たせる。
    # こうしておくと予備シートに支払先名と対象IDを入れるだけで1社増やせる（ブックの作り直しが要らない）。
    # J列は印刷範囲の外に置くので請求書には出ない。
    def build_invoice(sheet, payee, rows_c, cash_mode, n_rows, spare=False):
        ws = wb.create_sheet(sheet)
        widths = [24, 14, 11, 10, 13, 13, 14, 34] if cash_mode else [24, 14, 11, 10, 13, 40]
        for col, w in zip("ABCDEFGH", widths):
            ws.column_dimensions[col].width = w
        ws.column_dimensions["J"].width = 10

        ws["I1"] = "支払先(必須)"; ws["I1"].font = NOTE_FONT
        c = ws["J1"]; c.value = payee; c.font = BLUE; c.fill = INPUT_FILL; c.border = BOX

        doc = "請 求 書" if cash_mode else "支 払 明 細 書"
        ws["A1"] = doc; ws["A1"].font = Font(name=FONT, size=20, bold=True)
        ws["E1"] = "発行日"; ws["E1"].font = BOLD
        # TODAY() を使うと開き直すたびに発行日が変わる。空欄なら記入を促す。
        ws["F1"] = f'=IF({S["発行日"]}="","★M_請求設定 に発行日を入れてください",{S["発行日"]})'
        ws["F1"].number_format = "yyyy/mm/dd"; ws["F1"].font = BLACK
        ws["E2"] = "請求書番号"; ws["E2"].font = BOLD
        ws["F2"] = ('=IF($J$1="","","INV-"&SUBSTITUTE(集計_号機別!$B$3,"-","")&"-"&'
                    f'IFERROR(INDEX({payee_code_rng},MATCH($J$1,{payee_name_rng},0)),"??"))')
        ws["F2"].font = BLACK

        ws["A3"] = '=IF($J$1="","（J1 に支払先を入力してください）",$J$1&"　御中")'
        ws["A3"].font = Font(name=FONT, size=14, bold=True)
        ws["A3"].border = Border(bottom=Side(style="medium", color="1F3864"))

        ws["E4"] = f"={S['発行者名']}"; ws["E4"].font = BOLD
        ws["E5"] = f"={S['郵便番号']}"; ws["E5"].font = NOTE_FONT
        ws["E6"] = f"={S['住所']}"; ws["E6"].font = NOTE_FONT
        ws["E7"] = f'=IF({S["電話番号"]}="","","TEL "&{S["電話番号"]})'; ws["E7"].font = NOTE_FONT
        ws["E8"] = f'=IF({S["インボイス登録番号"]}="","（登録番号 未設定）","登録番号 "&{S["インボイス登録番号"]})'
        ws["E8"].font = NOTE_FONT

        ws["A5"] = "件名"; ws["A5"].font = BOLD
        ws["B5"] = "自動販売機 売上金 ご請求" if cash_mode else "自動販売機 設置手数料 お支払明細"
        ws["B5"].font = BLACK
        ws["A6"] = "対象期間"; ws["A6"].font = BOLD
        ws["B6"] = "=集計_号機別!$B$3"; ws["B6"].font = BLACK

        ws["A8"] = "ご請求金額" if cash_mode else "お支払金額"; ws["A8"].font = BOLD
        # 数値セルは列幅を超えてもはみ出し表示されないため、金額欄は結合して幅を確保する。
        ws.merge_cells("B8:D8")
        ws["B8"].font = Font(name=FONT, size=18, bold=True)
        ws["B8"].number_format = YEN
        ws["B8"].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[8].height = 26

        hr = 10
        if cash_mode:
            heads = ["対象", "回収現金", "日額", "稼働日数", "管理手数料", "その他控除", "ご請求額", "備考"]
        else:
            heads = ["対象", "対象売上", "料率", "税方式", "手数料額", "備考"]
        ncols = len(heads)
        tot_col = 7 if cash_mode else 5
        for i, h in enumerate(heads, 1): ws.cell(hr, i, h)
        style_header(ws, hr, ncols)
        ws.cell(hr, 10, "対象ID").font = NOTE_FONT

        first = hr + 1
        for j in range(n_rows):
            r = first + j
            tid = rows_c[j][1] if j < len(rows_c) else None
            note = rows_c[j][8] if j < len(rows_c) else ""
            method = rows_c[j][2] if j < len(rows_c) else ""
            c = ws.cell(r, 10, tid); c.font = BLUE; c.fill = INPUT_FILL; c.border = BOX

            blank = f'$J{r}=""'
            key = f'$J$1&"|"&$J{r}'

            def M(col, dflt="0"):
                return (f'IFERROR(INDEX(M_契約!${col}:${col},'
                        f'MATCH({key},M_契約!$A:$A,0)),{dflt})')

            ws.cell(r, 1, f'=IF({blank},"",IF($J{r}="{WHOLESALE_ID}","{WHOLESALE}",'
                          f'IFERROR(INDEX(M_号機!$B:$B,MATCH($J{r},M_号機!$A:$A,0)),"")))').font = BLACK
            method_f = M("E", '""')
            sales = (f'IF($J{r}="{WHOLESALE_ID}",'
                     f'SUMIFS(入力_その他売上!$F:$F,入力_その他売上!$C:$C,"{WHOLESALE}",'
                     f'入力_その他売上!$A:$A,集計_号機別!$B$3),'
                     f'SUMIFS(集計_号機別!$F:$F,集計_号機別!$A:$A,$J{r}))')
            cash = (f'IF($J{r}="{WHOLESALE_ID}",0,'
                    f'SUMIFS(集計_号機別!$D:$D,集計_号機別!$A:$A,$J{r}))')
            # 稼働日数・金額調整は精算シートと同じ定義（M_カレンダー ＋ 入力_月次調整）
            adjk = (f'入力_月次調整!$A:$A,集計_号機別!$B$3,入力_月次調整!$B:$B,$J$1,'
                    f'入力_月次調整!$C:$C,$J{r}')
            wdays = ('(IFERROR(INDEX(M_カレンダー!$C:$C,MATCH(集計_号機別!$B$3,M_カレンダー!$A:$A,0)),0)'
                     f'+SUMIFS(入力_月次調整!$D:$D,{adjk}))')
            adjamt = f'SUMIFS(入力_月次調整!$E:$E,{adjk})'

            if cash_mode:
                # 「単価×数量」を見せる（先方が根拠を確認できるように）
                ws.cell(r, 2, f'=IF({blank},"",{cash})').number_format = YEN
                ws.cell(r, 3, f'=IF({blank},"",{M("G")})').number_format = YEN
                ws.cell(r, 4, f'=IF({blank},"",{wdays})').number_format = NUM
                ws.cell(r, 5, f'=IF({blank},"",C{r}*D{r})').number_format = YEN
                ws.cell(r, 6, f'=IF({blank},"",{M("H")}+{adjamt})').number_format = YEN
                ws.cell(r, 7, f'=IF({blank},"",B{r}-E{r}-F{r})').number_format = YEN
            else:
                ws.cell(r, 2, f'=IF({blank},"",{sales})').number_format = YEN
                ws.cell(r, 3, f'=IF({blank},"",{M("F")})').number_format = PCT
                ws.cell(r, 4, f'=IF({blank},"",{M("J", chr(34) + "税込" + chr(34))})').font = BLACK
                base = (f'IF(D{r}="税抜換算",ROUNDUP(B{r}/1.08*C{r}*1.1,0),ROUNDUP(B{r}*C{r},0))')
                ws.cell(r, 5, f'=IF({blank},"",IF({method_f}="料率",'
                              f'MAX({base},{M("I")})-{adjamt},0))').number_format = YEN
            ws.cell(r, ncols, "⚠条件未確定のため0円" if method == "要確認" else note).font = NOTE_FONT
            for i in range(1, ncols + 1):
                cc = ws.cell(r, i)
                cc.font = NOTE_FONT if i == ncols else BLACK
                cc.border = BOX
                cc.alignment = Alignment(vertical="center", wrap_text=(i == ncols))

        last = first + n_rows - 1
        tl = get_column_letter(tot_col)
        tr3 = last + 1
        ws.cell(tr3, 1, "合計（税込）").font = BOLD; ws.cell(tr3, 1).border = BOX
        cc = ws.cell(tr3, tot_col, f"=SUM({tl}{first}:{tl}{last})")
        cc.font = BOLD; cc.number_format = YEN; cc.border = BOX
        for i in range(2, ncols + 1):
            ws.cell(tr3, i).border = BOX
        # 適格請求書は税率ごとの区分記載が要る。賃料は10%課税・税込表示なので税抜を逆算して並べる。
        ws.cell(tr3 + 1, 1, "　うち 10%対象（税抜）").font = BLACK
        c = ws.cell(tr3 + 1, tot_col, f"=ROUNDUP({tl}{tr3}/1.1,0)")
        c.font = BLACK; c.number_format = YEN; c.border = BOX
        ws.cell(tr3 + 2, 1, "　うち 消費税（10%）").font = BLACK
        c = ws.cell(tr3 + 2, tot_col, f"={tl}{tr3}-{tl}{tr3+1}")
        c.font = BLACK; c.number_format = YEN; c.border = BOX
        tot_row = tr3
        tr3 += 2

        ws["B8"] = f"={tl}{tot_row}"
        ws["B8"].font = Font(name=FONT, size=18, bold=True)
        ws["B8"].number_format = YEN
        ws["B8"].alignment = Alignment(horizontal="left", vertical="center")

        fr = tr3 + 2
        if cash_mode:
            ws.cell(fr, 1, "お振込先").font = BOLD
            ws.cell(fr, 2, f'=IF({S["振込先"]}="","★M_請求設定 に振込先を入力してください",{S["振込先"]})')
            ws.cell(fr, 2).font = BLACK
            ws.cell(fr + 1, 1, "お支払期日").font = BOLD
            ws.cell(fr + 1, 2, f"={S['支払期日']}").font = BLACK
            fr += 2
        ws.cell(fr, 1, "消費税").font = BOLD
        ws.cell(fr, 2, f'=IF({S["消費税の扱い"]}="","★M_請求設定 で消費税の扱いを確定してください",{S["消費税の扱い"]})')
        ws.cell(fr, 2).font = BLACK
        ws.cell(fr + 2, 1,
                "※ 金額は VMS の【回収時検量】（現金を回収した時点の実績）にもとづく。"
                ).font = NOTE_FONT
        ws.cell(fr + 3, 1, "※ 対象期間は 集計_号機別 の対象年月と連動する。").font = NOTE_FONT
        if spare:
            ws.cell(fr + 4, 1,
                    "※ 予備シート。使うときは J1 に支払先名、J列に対象ID（号機ID）を入れ、"
                    "M_契約 にも同じ支払先で行を足すこと。").font = NOTE_FONT

        ids = [c[1] for c in rows_c if c[1] != WHOLESALE_ID]
        dr = write_evidence(ws, fr + 6, ids, TARGET_YM, detail)
        ws.print_area = f"A1:{get_column_letter(ncols)}{dr}"
        ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0    # 算定根拠の明細は月ごとに伸びるので縦は複数ページ可
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    for payee in sorted({c[0] for c in CONTRACTS}):
        rows_c = [c for c in CONTRACTS if c[0] == payee]
        cash_mode = any(c[2] == "現金渡し" for c in rows_c)
        build_invoice("請求_" + PAYEE_SHEET[payee], payee, rows_c, cash_mode, len(rows_c))

    # 予備。支払先が増えたときにブックを作り直さずに済ませるための空シート。
    for i, cash_mode in ((1, False), (2, False), (3, True)):
        build_invoice(f"請求_予備{i}", None, [], cash_mode, 3, spare=True)

    # ---------------- 発行台帳 ----------------
    ws = wb.create_sheet("発行台帳")
    ws["A1"] = "請求書 発行台帳"; ws["A1"].font = TITLE
    ws["A2"] = ("※ PDFを発行したら1行足す。ブック上の請求書は対象年月を切り替えると数字が変わるので、"
                "「いつ・どの番号で・いくらで出したか」はこの台帳とPDFが正になる。")
    ws["A2"].font = NOTE_FONT
    heads = ["対象年月", "支払先", "請求書番号", "発行日", "金額（税込）", "PDFファイル名", "備考"]
    for i, h in enumerate(heads, 1): ws.cell(4, i, h)
    style_header(ws, 4, len(heads), [12, 34, 20, 12, 14, 34, 30])
    for r in range(5, 45):
        for i in range(1, 8):
            c = ws.cell(r, i); c.font = BLUE; c.fill = INPUT_FILL; c.border = BOX
        ws.cell(r, 1).number_format = "@"
        ws.cell(r, 4).number_format = "yyyy/mm/dd"
        ws.cell(r, 5).number_format = YEN

    # ---------------- 点検 ----------------
    ws = wb.create_sheet("点検")
    ws["A1"] = "点検・申し送り"; ws["A1"].font = TITLE
    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 26
    ws.column_dimensions['C'].width = 92
    # 公開版: 実案件の点検メモから、取引先固有の契約条件・確認日・実額を除いて要旨だけ残している。
    items = [
        ("旧ファイルの不具合", "合計行が #REF! になり、手数料の計算そのものが成立していないシートがあった。"),
        ("旧ファイルの不具合", "卸売上の参照セルが1行ズレており、卸の売上が二重計上されて手数料が過大になっていた。本ブックでは正しい本数で計上している。"),
        ("旧ファイルの不具合", "日次は料率で計算しているのに、合計欄だけ固定額のベタ打ちで上書きされているシートがあった。"),
        ("旧ファイルの不具合", "翌月1日の行が当月分として混入しているシートがあった。本ブックでは対象月のみ集計。"),
        ("数値の不一致", "旧シート1（手入力）と VMS の月合計が、一部の号機でズレていた。本ブックは VMS を正とするため、この不一致は構造的に発生しない。"),
        ("運用", "単価は M_単価 に履歴で持つ。旧ファイルは数式に単価が直接書かれており、値上げした月から計算が破綻する構造だった。"),
        ("運用", "自販機の売上金額は VMS 実績をそのまま採用するため、単価を掛け直す計算は行わない（値上げ月に新旧単価が混在しても、実額なので影響を受けない）。"),
        ("名称の不一致", "1台が最大4通りの名前で呼ばれている（VMS機名／旧スプシ通称／決済端末の登録名／実際の設置場所）。M_号機 の対応表を正とし、突合は号機IDと端末シリアルで行う。"),
        ("要確認（契約）", "条件が未確定の契約は M_契約 の方式を「要確認」にする。精算シートでは0円として扱い、警告を出す。"),
        ("データ", "取込タブは VMS から自動取得した直近1年の実績。検証月で、夜間検量は別途取得した月次値と、回収時検量は旧ファイルの貼付データと、それぞれ1円単位で一致することを確認済み。"),
        ("検量方式（重要）", "VMS には「夜間検量」と「回収時検量」の2方式がある。旧ファイルの貼付データは【回収時検量】だった。"
                             "両者の差は取得タイミングのズレではなく方式の違い＝「月末時点でまだ回収されていない機内の売上」。"),
        ("検量方式の使い分け", "精算・請求は【回収時検量】を正とする（回収した現金の実額と会計を一致させるため。"
                               "特に現金渡し方式の支払先は現金実額ベースなので必須）。"
                               "一方、曜日別需要・売切検知・補充計画は【夜間検量】でないと出せない（回収時は回収日にまとめて計上されるため日別の需要が分からない）。"
                               "→ 両方を取り込み、用途で使い分ける。"),
    ]
    for i, h in enumerate(["", "区分", "内容"], 1): ws.cell(3, i, h)
    style_header(ws, 3, 3)
    for j, (kind, text) in enumerate(items):
        r = 4 + j
        ws.cell(r, 1, j + 1).font = BLACK
        c = ws.cell(r, 2, kind); c.font = BOLD; c.border = BOX
        c = ws.cell(r, 3, text); c.font = BLACK; c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30
        ws.cell(r, 1).border = BOX

    # 入力セル（黄色）だけ編集可にして、数式セルの誤上書きを仕組みで止める。
    # パスワードは掛けない（解除できないと保守できなくなるため）。
    key = INPUT_FILL.start_color.rgb
    for sh in wb.worksheets:
        for row in sh.iter_rows():
            for cell in row:
                if cell.fill is not None and cell.fill.start_color.rgb == key:
                    cell.protection = Protection(locked=False)
        sh.protection.sheet = True
        sh.protection.enable()

    for s in wb.worksheets:
        s.sheet_view.showGridLines = False
    # openpyxl は数式のキャッシュ値を持たないため、開いた側で必ず再計算させる（PDF化に必須）
    wb.calculation.fullCalcOnLoad = True
    wb.save("sales-ledger.xlsx")
    print("saved")


if __name__ == "__main__":
    main()
