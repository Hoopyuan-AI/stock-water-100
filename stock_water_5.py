#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""臺灣上市股價水溫 5 檔測試版。Python 3.9+，僅使用標準函式庫。

執行: py -3 stock_water_5.py
輸出: water_output/五檔水溫監測.xlsx、五檔水溫總覽.csv、五檔歷史價格.csv
說明: Yahoo Finance 圖表 API 回傳之股息/拆股調整價，重基準到最新原始收盤價。
     需要網路；抓價失敗時明示錯誤，不會用模擬資料冒充實際行情。
"""
from __future__ import annotations
import argparse
import csv
import datetime as dt
import json
import math
from pathlib import Path
import random
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

STOCKS = [
    ("2330", "台積電"),
    ("2454", "聯發科"),
    ("2308", "台達電"),
    ("2317", "鴻海"),
    ("3711", "日月光投控"),
]
MA_NEAR_PCT = 0.02  # 股價距離10日/10月線 <= 2% 即「貼近」；可修改
TZ = dt.timezone(dt.timedelta(hours=8))  # 臺灣標準時間 UTC+8，Windows 無須另裝時區資料
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/125.0 Safari/537.36",
    "Accept": "application/json,text/plain,*/*",
}


def yahoo_chart(stock_code: str, retries: int = 3):
    """查詢兩年日價；查不到就報錯。"""
    stock = f"{stock_code}.TW"
    query = urllib.parse.urlencode({"range": "2y", "interval": "1d", "events": "div,splits"})
    reasons = []
    for host in ("query1.finance.yahoo.com", "query2.finance.yahoo.com"):
        url = f"https://{host}/v8/finance/chart/{stock}?{query}"
        for attempt in range(retries):
            try:
                req = urllib.request.Request(url, headers=HEADERS)
                with urllib.request.urlopen(req, timeout=20) as response:
                    obj = json.load(response)
                chart = obj.get("chart", {})
                if chart.get("error"):
                    raise RuntimeError(str(chart["error"]))
                result = chart.get("result")
                if not result:
                    raise RuntimeError("回傳資料為空")
                return result[0], url
            except (urllib.error.URLError, TimeoutError, ValueError, RuntimeError) as exc:
                reasons.append(f"{host} 第{attempt+1}次：{type(exc).__name__}: {exc}")
                if attempt < retries-1:
                    time.sleep(1.5 * (attempt + 1))
    raise RuntimeError("; ".join(reasons[-4:]))


def clean_history(result: dict):
    """將 Yahoo 調整收盤價轉換為以最新原始收盤價為基準的可比較歷史價。"""
    ind = result.get("indicators", {})
    quote = (ind.get("quote") or [{}])[0]
    adj_group = (ind.get("adjclose") or [{}])[0]
    timestamps = result.get("timestamp") or []
    closes = quote.get("close") or []
    adjusts = adj_group.get("adjclose") or []
    if not timestamps or not adjusts or len(adjusts) != len(timestamps):
        raise ValueError("缺少每日 Adj Close，無法依約定採除權息調整價，不以未還原價替代")
    series = []
    for epoch, close, adj in zip(timestamps, closes, adjusts):
        if close is None or adj is None or close <= 0 or adj <= 0:
            continue
        date = dt.datetime.fromtimestamp(int(epoch), TZ).date()
        series.append((date, float(close), float(adj)))
    if len(series) < 300:
        raise ValueError(f"只有 {len(series)} 個有效交易日，難以計算10月線與斜率")
    series.sort(key=lambda x: x[0])
    latest_day, latest_price, latest_adjusted = series[-1]
    # Adj Close 舊日報酬含權息；重訂基準，使最新價與實際最新收盤價一致。
    factor = latest_price / latest_adjusted
    history = [{"date": day.isoformat(), "raw_close": round(raw, 6),
                "normalized_close": round(adj * factor, 6)}
               for day, raw, adj in series]
    return history


def avg(numbers):
    if not numbers:
        raise ValueError("無可計算的資料")
    return sum(numbers) / len(numbers)


def moving_stats(values: list[float], n=10):
    if len(values) < n + 1:
        raise ValueError(f"需要至少 {n+1} 筆資料，實際僅 {len(values)} 筆")
    current, previous = avg(values[-n:]), avg(values[-n-1:-1])
    # 足以辨認持平：浮點計算誤差以 1e-10 的相對值當作 0
    diff = current - previous
    threshold = max(abs(current), abs(previous), 1.0) * 1e-10
    direction = "上揚" if diff > threshold else ("下彎" if diff < -threshold else "持平")
    return current, previous, direction


def period_closes(history, mode):
    periods = {}
    for row in history:
        day = dt.date.fromisoformat(row["date"])
        if mode == "week":
            iso = day.isocalendar()
            key = (iso.year, iso.week)
        else:
            key = (day.year, day.month)
        periods[key] = row["normalized_close"]  # 每週/月最後一個交易日收盤
    return list(periods.values())  # Python 3.9+ dict insertion order


def classify(close, day_ma, week_ma, month_ma, ddir, wdir, mdir, near=MA_NEAR_PCT):
    # 規則優先順序：三線齊揚 > 三線齊跌 > 貼長線 > 熱/冷 > 過渡
    if all(x == "上揚" for x in (ddir, wdir, mdir)):
        return "滾水區", "三均線全部上揚"
    if all(x == "下彎" for x in (ddir, wdir, mdir)):
        return "冰水區", "三均線全部下彎"
    near_d = abs(close / day_ma - 1) <= near
    near_m = abs(close / month_ma - 1) <= near
    if near_m and ddir == "上揚" and wdir == "上揚":
        return "溫水區", "貼近10月線，10日與10週均線上揚"
    if close > month_ma * (1 + near) and wdir == "上揚" and near_d:
        return "熱水區", "高於10月線超過容忍範圍，10週線向上且貼近10日線"
    if close < month_ma * (1 - near) and wdir == "上揚" and near_d:
        return "冷水區", "低於10月線超過容忍範圍，10週線向上且貼近10日線"
    return "過渡區", "不符合上述五種條件；不強制歸類"


def analyze(code, name, history, source_url):
    daily = [r["normalized_close"] for r in history]
    weekly = period_closes(history, "week")
    monthly = period_closes(history, "month")
    d, dprev, ddir = moving_stats(daily)
    w, wprev, wdir = moving_stats(weekly)
    m, mprev, mdir = moving_stats(monthly)
    latest = history[-1]
    close = latest["raw_close"]
    zone, reason = classify(close, d, w, m, ddir, wdir, mdir)
    return {
        "股票代號": code, "股票名稱": name, "行情日期": latest["date"], "最新收盤價(元)": round(close, 4),
        "10日均線(元)": round(d, 4), "10週均線(元)": round(w, 4), "10月均線(元)": round(m, 4),
        "10日方向": ddir, "10週方向": wdir, "10月方向": mdir,
        "距10日線(%)": round((close / d - 1) * 100, 4),
        "距10週線(%)": round((close / w - 1) * 100, 4),
        "距10月線(%)": round((close / m - 1) * 100, 4),
        "水溫分類": zone, "判定理由": reason, "來源URL": source_url,
    }


COLUMNS = ["股票代號", "股票名稱", "行情日期", "最新收盤價(元)", "10日均線(元)", "10週均線(元)",
           "10月均線(元)", "10日方向", "10週方向", "10月方向", "距10日線(%)", "距10週線(%)",
           "距10月線(%)", "水溫分類", "判定理由", "來源URL"]
HISTORY_COLUMNS = ["股票代號", "股票名稱", "日期", "實際收盤價(元)", "還原且對齊最新價(元)", "資料來源"]


def write_csv(path, headers, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for row in rows:
            writer.writerow(row)


def escape(value):
    import xml.sax.saxutils as sx
    return sx.escape(str(value), {'"': '&quot;'})


def excel_col(num):
    s = ""
    while num:
        num, rem = divmod(num - 1, 26)
        s = chr(65 + rem) + s
    return s


def excel_sheet_xml(rows, widths, frozen=0, filter_row=None, title_rows=()):
    """僅用 Python 標準函式庫產生 OOXML 工作表，不依賴額外安裝的模組。"""
    from io import StringIO
    out = StringIO()
    out.write('<?xml version="1.0" encoding="utf-8"?>')
    out.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    out.write('<sheetViews><sheetView workbookViewId="0">')
    if frozen:
        out.write(f'<pane ySplit="{frozen}" topLeftCell="A{frozen+1}" activePane="bottomLeft" state="frozen"/>')
    out.write('</sheetView></sheetViews>')
    out.write('<sheetFormatPr defaultRowHeight="17"/>')
    out.write('<cols>')
    for i, width in enumerate(widths, 1):
        out.write(f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>')
    out.write('</cols><sheetData>')
    for rowno, row in enumerate(rows, 1):
        out.write(f'<row r="{rowno}"' + (' ht="31" customHeight="1"' if rowno in title_rows else '') + '>')
        for colno, value in enumerate(row, 1):
            if value is None or value == "":
                continue
            address = f'{excel_col(colno)}{rowno}'
            style = 1 if rowno == 1 else (2 if rowno == 3 else 0)
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                if not math.isfinite(value):
                    continue
                out.write(f'<c r="{address}" s="{4 if style == 0 else style}"><v>{value}</v></c>')
            else:
                if rowno > 3 and str(value) in ('滾水區','熱水區','溫水區','冷水區','冰水區','過渡區','資料異常'):
                    style = {'滾水區':5,'熱水區':6,'溫水區':7,'冷水區':8,'冰水區':9,'過渡區':10,'資料異常':11}[value]
                out.write(f'<c r="{address}" s="{style}" t="inlineStr"><is><t>{escape(value)}</t></is></c>')
        out.write('</row>')
    out.write('</sheetData>')
    if filter_row and len(rows) >= filter_row:
        out.write(f'<autoFilter ref="A{filter_row}:{excel_col(max(len(x) for x in rows))}{len(rows)}"/>')
    out.write('</worksheet>')
    return out.getvalue().encode("utf-8")


def write_native_excel(path: Path, summary: list[dict], history_rows: list[list], report_date: str):
    """產生原生 .xlsx；無需 openpyxl / Excel 安裝。"""
    import zipfile
    title = [f"上市股票五檔水溫測試｜更新日 {report_date}"]
    dashboard = [title, ["判定基礎：10日、10週、10月均線；上揚/下彎與股價位置；貼近容忍 ±2%"], COLUMNS]
    for entry in summary:
        dashboard.append([entry.get(k, "") for k in COLUMNS])
    category_names = ["滾水區", "熱水區", "溫水區", "冷水區", "冰水區", "過渡區", "資料異常"]
    counts = {z: sum(1 for item in summary if item.get("水溫分類") == z) for z in category_names}
    dist = [["五檔水溫分布"], ["資料異常不計入技術分類，不可當成過渡區"], ["分類", "檔數", "占5檔比例(%)"]]
    dist += [[z, counts[z], round(counts[z] / len(STOCKS) * 100, 2)] for z in category_names]
    dist.append(["全部測試檔數", len(summary), 100])
    hist = [["五檔歷史股價｜每日資料"], ["還原且對齊最新價：以 Yahoo 調整收盤價按最新原始收盤價重基準"], HISTORY_COLUMNS]
    hist.extend(history_rows)
    rules = [
        ["五檔水溫系統｜計算及來源說明"],
        ["模式", "股價位置 + 10日/10週/10月均線方向"],
        ["短期均線", "最近10個交易日之調整收盤價平均"],
        ["中期均線", "最近10個週收盤價平均（含未完成當週）"],
        ["長期均線", "最近10個月收盤價平均（含未完成當月）"],
        ["均線方向", "各自與上一日/週/月的相同10期均線比較"],
        ["貼近條件", f"當期收盤價距指定均線不超過 ±{MA_NEAR_PCT:.0%}"],
        ["分類順序 1", "日週月三條均線全上揚：滾水區"],
        ["分類順序 2", "日週月三條均線全下彎：冰水區"],
        ["分類順序 3", "靠近10月線且日、週線上揚：溫水區"],
        ["分類順序 4", "收盤高於10月線 2% 以上，週線上揚，靠近10日線：熱水區"],
        ["分類順序 5", "收盤低於10月線 2% 以上，週線上揚，靠近10日線：冷水區"],
        ["分類順序 6", "其餘不完全符合：過渡區（不硬分類）"],
        ["收盤價", "最新原始收盤價；非盤中即時報價"],
        ["歷史價格", "Yahoo Finance Adj Close調整價；重新對齊最新原始收盤價（不與原始價格混用）"],
        ["Yahoo 資料來源", "https://finance.yahoo.com/quote/2330.TW/history/"],
        ["說明", "此程式適用上市.TW股票；非投資建議。Yahoo可能限速或改版，失敗會清楚註記"],
    ]
    sheets = [
        ("五檔總覽", dashboard, [13,15,16,17,16,16,16,12,12,12,16,16,16,12,47,75], 3, 3),
        ("水溫分布", dist, [26,13,25], 3, 3),
        ("歷史股價", hist, [15,18,16,20,25,76], 3, 3),
        ("規則與來源", rules, [25,100], 0, None),
    ]
    names = [x[0] for x in sheets]
    content_types = ['<?xml version="1.0" encoding="utf-8"?>', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
                     '<Default Extension="xml" ContentType="application/xml"/>',
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
                     '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    for i in range(1, len(sheets)+1):
        content_types.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    content_types.append('</Types>')
    root_rels = ('<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    workbook = ('<?xml version="1.0" encoding="utf-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
                ''.join(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>' for i, name in enumerate(names,1)) +
                '</sheets></workbook>')
    book_rels = ('<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                 ''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,len(sheets)+1)) +
                 f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
                 '</Relationships>')
    styles = '''<?xml version="1.0" encoding="utf-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4"><font><sz val="10"/><name val="Aptos"/></font><font><b/><sz val="15"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/></font><font><b/><sz val="10"/><color rgb="FF17324D"/></font></fonts>
<fills count="12"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF15324D"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF2C658A"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFAFAFA"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFEDD5"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFDBEAFE"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFE0F2FE"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="12"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="3" fillId="5" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="3" fillId="6" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="3" fillId="7" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="3" fillId="8" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="3" fillId="9" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="3" fillId="10" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="3" fillId="11" borderId="0" xfId="0" applyFill="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(content_types))
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", book_rels)
        z.writestr("xl/styles.xml", styles)
        for idx, (name, rows, widths, frozen, filt) in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{idx}.xml", excel_sheet_xml(rows,widths,frozen,filter_row=filt,title_rows=(1,)))


def self_test():
    # 驗證五種判斷邏輯及移動平均平均值，完全不需網路
    assert classify(100, 99, 95, 90, "上揚", "上揚", "上揚")[0] == "滾水區"
    assert classify(100, 105, 110, 115, "下彎", "下彎", "下彎")[0] == "冰水區"
    assert classify(100, 99, 97, 99, "上揚", "上揚", "下彎")[0] == "溫水區"
    assert classify(110, 110, 105, 95, "下彎", "上揚", "上揚")[0] == "熱水區"
    assert classify(90, 90, 91, 100, "下彎", "上揚", "上揚")[0] == "冷水區"
    assert classify(110, 100, 108, 110, "持平", "下彎", "上揚")[0] == "過渡區"
    ma, prior, direction = moving_stats([float(i) for i in range(1, 13)])
    assert abs(ma - 7.5) < 1e-8 and abs(prior - 6.5) < 1e-8 and direction == "上揚"
    print("PASS：六種分類（含過渡區）、均線平均與方向測試正常。")


def main():
    parser = argparse.ArgumentParser(description="上市5檔股票：10日、10週、10月均線＋水溫分析")
    parser.add_argument("--output", default="water_output", help="輸出資料夾，預設 water_output")
    parser.add_argument("--self-test", action="store_true", help="無網路測試分類規則")
    args = parser.parse_args()
    if args.self_test:
        return self_test()
    out = Path(args.output)
    out.mkdir(exist_ok=True,parents=True)
    summaries = []
    histories = []
    report_date = dt.datetime.now(TZ).strftime("%Y-%m-%d %H:%M 台北時間")
    for code, name in STOCKS:
        print(f"讀取 {code} {name} ...", flush=True)
        try:
            chart, url = yahoo_chart(code)
            history = clean_history(chart)
            summary = analyze(code,name,history,url)
            summaries.append(summary)
            for item in history:
                histories.append([code,name,item["date"],item["raw_close"],item["normalized_close"],url])
            last_date = dt.date.fromisoformat(summary["行情日期"])
            if (dt.datetime.now(TZ).date() - last_date).days > 5:
                print(f"  注意：最新收盤僅至 {last_date}，請確認是否停牌/休市/行情延遲。")
            print(f"  {summary['行情日期']} 收盤{summary['最新收盤價(元)']}，水溫：{summary['水溫分類']}")
        except Exception as exc:
            print(f"  下載/計算失敗：{exc}", file=sys.stderr)
            summaries.append({"股票代號":code,"股票名稱":name,"水溫分類":"資料異常","判定理由":str(exc)[:500]})
        time.sleep(1.0)  # 降低來源站限速機率
    write_csv(out / "五檔水溫總覽.csv", COLUMNS, [[s.get(c,"") for c in COLUMNS] for s in summaries])
    write_csv(out / "五檔歷史價格.csv", HISTORY_COLUMNS, histories)
    write_native_excel(out / "五檔水溫監測.xlsx", summaries, histories, report_date)
    print("\n完成：")
    for name in ("五檔水溫監測.xlsx", "五檔水溫總覽.csv", "五檔歷史價格.csv"):
        print("  ", (out/name).resolve())
    valid = sum(s.get("水溫分類") != "資料異常" for s in summaries)
    print(f"成功 {valid}/{len(STOCKS)} 檔；資料異常 {len(STOCKS)-valid} 檔。")
    return 0 if valid == len(STOCKS) else 1


if __name__ == "__main__":
    sys.exit(main())
