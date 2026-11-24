import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1]
FONT = "Meiryo UI"
NAVY, TEAL = "14213D", "0E8C9A"
YELLOW = PatternFill("solid", fgColor="FFF2CC")
HEAD = PatternFill("solid", fgColor=NAVY)
SEC = PatternFill("solid", fgColor="DCEBEE")
thin = Side(style="thin", color="B7C4CF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

ROWS = [
    ("1. 基本情報", None, None),
    ("製品名・品番", "例: 製品A(型番)", "—"),
    ("製品の種類", "LIST_TYPE", "第1回"),
    ("記入日・記入者", "", "—"),
    ("2. 温度", None, None),
    ("温度の範囲(何℃から何℃か)", "", "第2回: 言葉ではなく数字で確かめる"),
    ("管理するのはどの温度か", "LIST_TARGET", "第2回: 製品の温度と周りの温度は違う"),
    ("一時的に外れてもよい範囲・時間", "", "決まっていなければ「なし」と確かめる"),
    ("3. 湿度", None, None),
    ("湿度の範囲(何%以下など)", "", "第3回: 高すぎても低すぎても困る"),
    ("防湿袋・乾燥剤・湿度インジケーターの有無", "", "第2回: 袋を勝手に開けない"),
    ("静電気の対策(帯電防止の包装など)", "", "第3回"),
    ("4. 電池(電池・電池を組み込んだ製品の場合)", None, None),
    ("充電率の条件", "", "第2回: 保管と輸送で違うことがある"),
    ("SDS第14項: 国連番号・輸送時の危険性クラス", "", "第6回: SDSの「輸送上の注意」"),
    ("保管と運ぶ方法(航空・海上・陸上)ごとの扱いと確認先", "", "第6回: 法律と運び方で決まり方が違う"),
    ("全固体電池の場合: 種類(硫化物系など)と湿度の条件", "", "第6回: 硫化物系は湿気に注意"),
    ("危険物・消防法の扱いの確認先と結果", "", "第3回・第6回: 自分で判断せず相談する"),
    ("5. 取り扱い", None, None),
    ("衝撃・積み方の注意(向き・段数・固定)", "", "第4回"),
    ("6. 記録と、外れたときの対応", None, None),
    ("記録の要否・間隔・提出の形", "", "第5回: 記録は「使う」もの"),
    ("外れたときの連絡先(社内・社外)", "", "第5回: すぐに報告する"),
    ("7. 根拠", None, None),
    ("根拠の種類", "LIST_BASIS", "第2回: 根拠は4種類"),
    ("根拠の文書名・版・ページ", "", "仕様書・指示書・契約書など"),
    ("確認した相手・日付", "", "口頭の場合は記録を残す"),
    ("8. まとめ", None, None),
    ("わからないこと・次にやること", "", "—"),
]

EXAMPLE = {
    "製品名・品番": "充電式の計測器 B(型番 XX-100)※架空",
    "製品の種類": "電池を組み込んだ製品",
    "記入日・記入者": "2026年12月1日 物流課 山田 ※架空",
    "温度の範囲(何℃から何℃か)": "保管: 0〜35℃ / 輸送: 0〜35℃ ※説明用の架空の値",
    "管理するのはどの温度か": "周りの空気の温度",
    "一時的に外れてもよい範囲・時間": "決まりなし(メーカーに確認済み)",
    "湿度の範囲(何%以下など)": "80%以下・結露させないこと ※説明用の架空の値",
    "防湿袋・乾燥剤・湿度インジケーターの有無": "なし(個箱に乾燥剤のみ)",
    "静電気の対策(帯電防止の包装など)": "特になし(メーカーに確認済み)",
    "充電率の条件": "出荷時 30%程度(メーカー回答)。輸送の条件は運送会社に確認中",
    "SDS第14項: 国連番号・輸送時の危険性クラス": "UN3481(機器に組み込まれた電池)。クラスはSDSの記載どおり転記 ※説明用の架空の記入",
    "保管と運ぶ方法(航空・海上・陸上)ごとの扱いと確認先": "陸上のみ。航空で送る場合は運送会社に事前確認",
    "全固体電池の場合: 種類(硫化物系など)と湿度の条件": "該当なし(リチウムイオン電池)",
    "危険物・消防法の扱いの確認先と結果": "社内の安全管理担当に相談中。倉庫会社には事前に連絡済み",
    "衝撃・積み方の注意(向き・段数・固定)": "天地無用。段積みは4段まで(仕様書 5章)",
    "記録の要否・間隔・提出の形": "夏季(6〜9月)のみ温度記録。10分間隔。月末に取引先へ提出",
    "外れたときの連絡先(社内・社外)": "物流課長 / 品質保証部 / メーカー窓口",
    "根拠の種類": "③ メーカーの仕様書",
    "根拠の文書名・版・ページ": "製品B 取扱説明書 第2版 3章・5章",
    "確認した相手・日付": "メーカー品質保証部 佐藤様 2026年11月28日(メール) ※架空",
    "わからないこと・次にやること": "夏の荷待ちの時間を運送会社と確認する",
}

def f(bold=False, color="1B1F24", size=11):
    return Font(name=FONT, bold=bold, color=color, size=size)

def build(ws, example):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 38
    ws["A1"] = "温湿度条件の確認シート" + ("(記入例)" if example else "")
    ws["A1"].font = f(True, NAVY, 16)
    ws["A2"] = "講座「新しい製品の物流担当になったら — 温度と湿度の管理 入門」第2回・第6回の演習用"
    ws["A2"].font = f(False, "59636E", 10)
    if example:
        ws["A3"] = "※ この記入例の会社・人物・製品・数値は、すべて説明用の架空のものです。実際の条件は、担当する製品の仕様書や取引先の指示で確かめてください。"
        ws["A3"].font = f(True, "C0392B", 10)
    else:
        ws["A3"] = "使い方: 黄色の欄に記入します。「▼」のある欄は一覧から選べます。わからない欄は空けたままにせず「不明」と書き、8. に次にやることを書きます。"
        ws["A3"].font = f(False, "1B1F24", 10)
    ws.merge_cells("A3:C3")
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[3].height = 32
    r = 5
    for c, h in zip("ABC", ["確認する項目", "記入欄", "ヒント(講座の回)"]):
        cell = ws[f"{c}{r}"]; cell.value = h; cell.font = f(True, "FFFFFF"); cell.fill = HEAD; cell.border = BOX
        cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[r].height = 22
    dv_type = DataValidation(type="list", formula1='"電子部品,電池(セル・電池パック),電池を組み込んだ製品,その他"', allow_blank=True)
    dv_target = DataValidation(type="list", formula1='"製品そのものの温度,周りの空気の温度,両方,不明"', allow_blank=True)
    dv_basis = DataValidation(type="list", formula1='"① 法令,② 業界・国際の規格,③ メーカーの仕様書,④ 契約・取引先の指示,不明"', allow_blank=True)
    for dv in (dv_type, dv_target, dv_basis):
        ws.add_data_validation(dv)
    for label, kind, hint in ROWS:
        r += 1
        if kind is None:
            ws[f"A{r}"] = label; ws.merge_cells(f"A{r}:C{r}")
            ws[f"A{r}"].font = f(True, TEAL); ws[f"A{r}"].fill = SEC
            ws.row_dimensions[r].height = 20
            continue
        ws[f"A{r}"] = label; ws[f"A{r}"].font = f()
        b = ws[f"B{r}"]
        if example:
            b.value = EXAMPLE.get(label, "")
        elif kind == "例: 製品A(型番)":
            b.value = None
        b.fill = YELLOW; b.font = f()
        if kind == "LIST_TYPE": dv_type.add(b)
        if kind == "LIST_TARGET": dv_target.add(b)
        if kind == "LIST_BASIS": dv_basis.add(b)
        if kind and kind.startswith("LIST") and not example:
            ws[f"A{r}"] = label + " ▼"
        ws[f"C{r}"] = hint; ws[f"C{r}"].font = f(False, "59636E", 10)
        for c in "ABC":
            ws[f"{c}{r}"].border = BOX
            ws[f"{c}{r}"].alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 36
    r += 2
    ws[f"A{r}"] = "この配布物は、講座の受講者の個人・社内での利用に限ります。再配布はできません。法令やルールは変わることがあるため、最新の情報を確かめてください。"
    ws[f"A{r}"].font = f(False, "59636E", 9)
    ws.merge_cells(f"A{r}:C{r}")
    ws[f"A{r}"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[r].height = 28
    ws.print_area = f"A1:C{r}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.freeze_panes = "A6"

wb = Workbook()
ws1 = wb.active; ws1.title = "記入用"; build(ws1, False)
ws2 = wb.create_sheet("記入例"); build(ws2, True)
wb.save(OUT)
print("saved", OUT)
