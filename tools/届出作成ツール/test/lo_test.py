# -*- coding: utf-8 -*-
"""LibreOffice（VBA互換モード）での自動試験。Windows Excel での最終確認の前の予備試験。

使い方（リポジトリ直下）:
    python3 "tools/届出作成ツール/test/lo_test.py" <作業フォルダ>

作業フォルダにブックをコピーして試験する（リポジトリ内のブックは変更しない）。
注意：LibreOffice と Excel では StrConv（Shift_JIS 判定）などの挙動が違うため、
      文字コード関連の試験結果は参考扱い。最終確認は Windows Excel で行う。
"""
import shutil
import subprocess
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue

HERE = Path(__file__).resolve().parent
TOOL = HERE.parent
BOOK = TOOL / "届出作成ツール_v1.0.xlsx"


def pv(n, v):
    p = PropertyValue()
    p.Name, p.Value = n, v
    return p


def connect():
    proc = subprocess.Popen(["soffice", "--headless", "--norestore", "--nologo",
                             '--accept=socket,host=localhost,port=2002;urp;'],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(60):
        try:
            ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            return proc, ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        except Exception:
            time.sleep(0.5)
    raise RuntimeError("LibreOffice に接続できません")


def serial_to_date(v):
    return (datetime(1899, 12, 30) + timedelta(days=v)).date()


class Book:
    def __init__(self, doc):
        self.doc = doc

    def sh(self, name):
        return self.doc.Sheets.getByName(name)

    def cell(self, sheet, addr):
        return self.sh(sheet).getCellRangeByName(addr)

    def set(self, sheet, addr, v):
        c = self.cell(sheet, addr)
        if isinstance(v, date):
            c.setValue((datetime(v.year, v.month, v.day) - datetime(1899, 12, 30)).days)
        elif isinstance(v, (int, float)):
            c.setValue(v)
        else:
            c.setString(v)

    def val(self, sheet, addr):
        c = self.cell(sheet, addr)
        return c.getString()

    def find_col(self, sheet, header_text, row=4):
        s = self.sh(sheet)
        for col in range(0, 300):
            t = s.getCellByPosition(col, row - 1).getString().replace("\n", "")
            if t == header_text:
                return col
        raise KeyError(header_text)


# ---------------------------------------------------------------- 機能A：チェックリスト
CASES = {
    "D001": dict(区分="採用", 基準日=date(2026, 11, 1), 生年月日=date(1990, 5, 12), 国籍="日本", 在留資格="該当なし",
                 社保="有", 雇保="有", 扶養="有", 特徴="有", 共済="無", 役職="無"),
    "D002": dict(区分="採用", 基準日=date(2026, 11, 1), 生年月日=date(1998, 2, 3), 国籍="外国籍", 在留資格="特定技能",
                 社保="有", 雇保="有", 扶養="無", 特徴="無", 共済="無", 役職="無"),
    "D003": dict(区分="採用", 基準日=date(2026, 11, 16), 生年月日=date(1955, 8, 20), 国籍="日本", 在留資格="該当なし",
                 社保="有", 雇保="有", 扶養="無", 特徴="無", 共済="無", 役職="無"),
    "D004": dict(区分="退職", 基準日=date(2026, 11, 30), 生年月日=date(1985, 1, 30), 国籍="日本", 在留資格="該当なし",
                 社保="有", 雇保="有", 特徴="有", 共済="有", 役職="無", 離職票="有", 退職金="有", 退職理由="自己都合"),
    "D005": dict(区分="退職", 基準日=date(2026, 11, 15), 生年月日=date(1992, 7, 7), 国籍="日本", 在留資格="該当なし",
                 社保="無", 雇保="無", 特徴="無", 共済="無", 役職="有", 離職票="無", 退職金="無", 退職理由="自己都合"),
}
EXPECT = {   # 届出名の一部 → 期限（日付 or 文言）
    "D001": {"資格取得届": "2026/11/5", "被扶養者": "2026/11/5", "雇用保険 資格取得届": "2026/12/10",
             "特別徴収への切替": "随時", "扶養控除等申告書": "初回給与"},
    "D002": {"資格取得届": "2026/11/5", "ローマ字氏名届": "2026/11/5", "雇用保険 資格取得届": "2026/12/10",
             "外国人雇用状況届出（雇保加入者）": "同時", "特定技能雇用契約に係る届出（締結）": "2026/11/14",
             "扶養控除等申告書": "初回給与"},
    "D003": {"資格取得届": "2026/11/20", "70歳以上被用者 該当届": "2026/11/20", "雇用保険 資格取得届": "2026/12/10",
             "扶養控除等申告書": "初回給与"},
    "D004": {"資格喪失届": "2026/12/5", "雇用保険 資格喪失届": "2026/12/11", "離職証明書": "2026/12/11",
             "給与所得者異動届出書": "2026/12/10", "退職手当共済 退職届": "共済規程", "退職手当金請求書": "退職後",
             "給与所得の源泉徴収票": "2026/12/30", "退職所得の源泉徴収票": "2026/12/30", "退職証明書": "請求"},
    "D005": {"変更届": "2026/11/24", "体制届": "前月15日", "給与所得の源泉徴収票": "2026/12/15", "退職証明書": "請求"},
}
INPUT_ROW = {"氏名": 4, "区分": 5, "基準日": 6, "生年月日": 7, "国籍": 8, "在留資格": 9, "社保": 10, "雇保": 11,
             "扶養": 12, "特徴": 13, "共済": 14, "役職": 15, "離職票": 16, "退職金": 17, "障害者": 18, "退職理由": 19}


def test_checklist(bk, results):
    s = "チェックリスト"
    for pid, case in CASES.items():
        for k, r in INPUT_ROW.items():
            bk.set(s, f"D{r}", "")
        bk.set(s, "D4", pid)
        for k, v in case.items():
            bk.set(s, f"D{INPUT_ROW[k]}", v)
        bk.doc.calculateAll()
        got = {}
        for r in range(23, 48):
            nm = bk.val(s, f"C{r}")
            if nm:
                c = bk.cell(s, f"E{r}")
                due = c.getString()
                if c.FormulaResultType2 == 1:      # 数値（日付）
                    d = serial_to_date(c.getValue())
                    due = f"{d.year}/{d.month}/{d.day}"
                got[nm] = due
        exp = EXPECT[pid]
        ok = True
        lines = []
        for key, want in exp.items():
            hit = [(n, d) for n, d in got.items() if key in n]
            if not hit:
                ok = False
                lines.append(f"  NG 届出が出ない：{key}")
            elif want not in hit[0][1]:
                ok = False
                lines.append(f"  NG 期限違い：{hit[0][0]} 期待={want} 実際={hit[0][1]}")
        unexpected = [n for n in got if not any(k in n for k in exp)]
        if unexpected:
            ok = False
            lines.append("  NG 想定外の届出：" + "、".join(unexpected))
        results.append((f"機能A {pid}（{case['区分']}）", ok, "\n".join(lines) or f"  {len(got)}件すべて期待どおり"))
        if pid == "D004":
            results.append(("機能A D004 出力一覧", True, "\n".join(f"  {n}：{d}" for n, d in got.items())))


# ---------------------------------------------------------------- 機能B：VBA
TEST_MODULE = '''Option VBASupport 1
Option Explicit
Public Sub RunKind(ByVal kind As String, ByVal mode As String)
    gTestMode = True
    ThisWorkbook.Worksheets("出力").Range("B3").Value = kind
    If mode = "check" Then
        btnCheck
    ElseIf mode = "clear" Then
        btnClearNyushoku
    Else
        btnExport
    End If
End Sub
'''


def install_vba(doc):
    src = (TOOL / "vba" / "modExport_utf8.bas").read_text(encoding="utf-8")
    body = "\n".join(l for l in src.splitlines() if not l.startswith("Attribute VB_Name"))
    # LibreOffice の StrConv(vbFromUnicode) は日本語を常に「?」にするため、Shift_JIS 判定だけ外して試験する
    # （判定そのものは Windows Excel で確認する項目）
    assert "lvl = SjisLevel(s)" in body
    body = body.replace("lvl = SjisLevel(s)", "lvl = 0")
    libs = doc.BasicLibraries
    libs.VBACompatibilityMode = True
    lib = libs.getByName("Standard")
    lib.insertByName("modExport", "Option VBASupport 1\n" + body)
    lib.insertByName("modTest", TEST_MODULE)


def run(doc, kind, mode):
    sp = doc.getScriptProvider()
    s = sp.getScript("vnd.sun.star.script:Standard.modTest.RunKind?language=Basic&location=document")
    s.invoke((kind, mode), (), ())


def messages(bk):
    out = []
    for r in range(11, 200):
        lv = bk.val("出力", f"A{r}")
        if not lv:
            break
        out.append(" | ".join(bk.val("出力", f"{c}{r}") for c in "ABCDEF"))
    return out


def confirm_rows(bk):
    s = bk.sh("内容確認")
    rows = []
    for r in range(3, 30):
        vals = [s.getCellByPosition(c, r).getString() for c in range(0, 40)]
        if not any(vals):
            break
        while vals and vals[-1] == "":
            vals.pop()
        rows.append(vals)
    return rows


def test_vba(bk, work, results):
    doc = bk.doc
    install_vba(doc)
    # 届書コードはCSV仕様書で確定する値。試験では仮の値を入れる
    lay = bk.sh("レイアウト定義")
    for r in range(4, 200):
        if lay.getCellByPosition(3, r).getString() == "届書コード":
            lay.getCellByPosition(5, r).setString("TEST")

    # 1) 社保取得：正常
    run(doc, "社保取得", "check")
    msg = messages(bk)
    results.append(("機能B 社保取得 チェック", not any(m.startswith("エラー") for m in msg),
                    "  " + bk.val("出力", "B6") + ("\n  " + "\n  ".join(msg) if msg else "")))
    run(doc, "社保取得", "export")
    out_dir = work / "出力"
    csvs = sorted(out_dir.glob("社保取得_*.csv")) if out_dir.exists() else []
    detail = "  " + bk.val("出力", "B6")
    if csvs:
        detail += "\n  --- " + csvs[-1].name + "\n  " + "\n  ".join(csvs[-1].read_text(encoding="utf-8", errors="replace").splitlines())
    results.append(("機能B 社保取得 CSV出力", len(csvs) == 1, detail))
    nyu_done = bk.find_col("入職台帳", "出力済")
    marks = [bk.sh("入職台帳").getCellByPosition(nyu_done, r).getString() for r in range(4, 7)]
    results.append(("機能B 出力済の記録", all("社保取得:" in m for m in marks), "  " + " / ".join(marks)))
    log = [bk.val("ログ", f"{c}5") for c in "ABCDEF"]
    results.append(("機能B ログ記録", "成功" in log[5], "  " + " | ".join(log)))

    # 2) 必須欠落 → エラーでファイルを作らない
    sex = bk.find_col("入職台帳", "性別")
    bk.sh("入職台帳").getCellByPosition(sex, 4).setString("")
    run(doc, "社保取得", "export")
    msg = messages(bk)
    n_csv = len(list(out_dir.glob("社保取得_*.csv")))
    results.append(("機能B 必須欠落でエラー・未出力", any("必須項目が空" in m for m in msg) and n_csv == 1,
                    "  " + "\n  ".join(msg[:5]) + f"\n  CSV数={n_csv}（1のままなら未出力）"))
    bk.sh("入職台帳").getCellByPosition(sex, 4).setString("男")

    # 3) カンマ入り → エラー（カナ欄は全角変換で「，」になり区切りにならないため、変換のない氏名欄で試す）
    kana = bk.find_col("入職台帳", "氏名★必須")
    old = bk.sh("入職台帳").getCellByPosition(kana, 4).getString()
    bk.sh("入職台帳").getCellByPosition(kana, 4).setString("山田 太郎,")
    run(doc, "社保取得", "check")
    msg = messages(bk)
    results.append(("機能B カンマでエラー", any("カンマ" in m for m in msg),
                    "  " + bk.val("出力", "B6") + "\n  " + "\n  ".join(msg[:3])))
    bk.sh("入職台帳").getCellByPosition(kana, 4).setString(old)

    # 4) 機種依存文字（髙）→ 警告 は LibreOffice では判定できないため、Windows Excel で確認する（テスト記録に記載）

    # 5) 社保喪失（D005は対象条件で除外される）
    run(doc, "社保喪失", "export")
    msg = messages(bk)
    csvs = sorted(out_dir.glob("社保喪失_*.csv"))
    detail = "  " + bk.val("出力", "B6") + "\n  " + "\n  ".join(msg)
    if csvs:
        detail += "\n  --- " + csvs[-1].name + "\n  " + csvs[-1].read_text(encoding="utf-8", errors="replace")
    results.append(("機能B 社保喪失 CSV（D004のみ）", len(csvs) == 1 and any("除外" in m for m in msg), detail))

    # 6) 雇保取得・喪失（転記一覧）
    for kind in ("雇保取得", "雇保喪失"):
        run(doc, kind, "export")
        msg = messages(bk)
        rows = confirm_rows(bk)
        ok = not any(m.startswith("エラー") for m in msg) and len(rows) >= 2
        detail = "  " + bk.val("出力", "B6") + ("\n  " + "\n  ".join(msg) if msg else "")
        detail += "\n  " + "\n  ".join(" | ".join(r) for r in rows)
        results.append((f"機能B {kind} 転記一覧", ok, detail))

    # 7) 2回目の出力：出力済でも出力でき、出力済欄に追記される
    run(doc, "社保取得", "export")
    marks = [bk.sh("入職台帳").getCellByPosition(nyu_done, r).getString() for r in range(4, 7)]
    n_csv = len(list(out_dir.glob("社保取得_*.csv")))
    results.append(("機能B 2回目の出力（追記・別ファイル）", all(m.count("社保取得:") == 2 for m in marks) and n_csv == 2,
                    "  " + bk.val("出力", "B6") + f"\n  CSV数={n_csv}\n  " + marks[0]))

    # 8) 台帳クリア：値は消え、表示用の数式は残る
    run(doc, "社保取得", "clear")
    ws = bk.sh("入職台帳")
    c5 = ws.getCellRangeByName("C5")
    left = [ws.getCellByPosition(c, 4).getString() for c in (0, 1, 4)]
    results.append(("機能B 台帳クリア", c5.getFormula().startswith("=") and not any(left) and c5.getString() == "",
                    "  " + bk.val("出力", "B6") + f"\n  C5の数式={c5.getFormula()[:40]}… 残った値={left}"))

    # 9) 順番の重複 → エラー
    for r in range(4, 200):
        if lay.getCellByPosition(0, r).getString() == "社保取得" and lay.getCellByPosition(2, r).getValue() == 3:
            lay.getCellByPosition(2, r).setValue(2)
            break
    run(doc, "社保取得", "check")
    msg = messages(bk)
    results.append(("機能B 順番の重複でエラー", any("順番" in m for m in msg), "  " + "\n  ".join(msg[:3])))


def main():
    work = Path(sys.argv[1]).resolve()
    if work.exists():
        shutil.rmtree(work)
    work.mkdir(parents=True)
    book = work / "届出作成ツール_試験.xlsx"
    shutil.copy(BOOK, book)
    proc, desktop = connect()
    results = []
    try:
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(book)), "_blank", 0,
                                           (pv("Hidden", False), pv("MacroExecutionMode", 4)))
        bk = Book(doc)
        test_checklist(bk, results)
        test_vba(bk, work, results)
        doc.storeToURL(uno.systemPathToFileUrl(str(work / "試験後.ods")), (pv("FilterName", "calc8"),))
        doc.close(True)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        proc.wait(timeout=30)
    n_ok = sum(1 for _, ok, _ in results if ok)
    for name, ok, detail in results:
        print(("OK " if ok else "NG ") + name)
        if detail:
            print(detail)
    print(f"\n結果：{n_ok}/{len(results)} 合格")


if __name__ == "__main__":
    main()
