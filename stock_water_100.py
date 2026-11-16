#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""臺灣上市市值前100大｜五水溫每日追蹤 V3 (Python 3.9+, 無額外套件)。

資料：期交所上市市值占比前100名 + Yahoo Finance 2年日價(Adj Close回基準)。
每日輸出 Excel、CSV、HTML，同時維護歷史資料、昨日異動、排行變化。
不產生虛構行情；任何擷取失敗必顯示狀態。
"""
from __future__ import annotations
import argparse
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
import csv
import datetime as dt
import hashlib
from html import escape as html_escape
from html.parser import HTMLParser
import json
from pathlib import Path
import re
import sys
import time
import urllib.request
import zipfile

import stock_water_5 as core

TAIFEX_URL = 'https://www.taifex.com.tw/cht/9/futuresQADetail'
TZ = core.TZ
CATEGORIES = ['滾水區','熱水區','溫水區','冷水區','冰水區','過渡區']
COLORS = {'滾水區':'#DB3647','熱水區':'#ED8B39','溫水區':'#499867',
          '冷水區':'#3988C7','冰水區':'#9BCAE9','過渡區':'#98A2B3',
          '資料不足':'#CBD2DA','資料異常':'#E1C8C9'}
GRADE = {'滾水區':5,'熱水區':4,'溫水區':3,'冷水區':2,'冰水區':1}
DETAIL_FIELDS = ['市值排名','股票代號','公司名稱','市值占大盤比重(%)','行情日期','最新收盤價(元)',
                 '10日均線(元)','10週均線(元)','10月均線(元)','10日方向','10週方向',
                 '10月方向','方向組合(日/週/月)','距10日線(%)','距10週線(%)','距10月線(%)',
                 '今日水溫','前一交易日','前日水溫','水溫變化','資料狀態','判定理由','股價來源URL']
HIST_FIELDS = ['交易日','市值排名','股票代號','公司名稱','市值占大盤比重(%)','收盤價',
               '水溫','10日方向','10週方向','10月方向','排名抓取時間','股價來源URL']
EVENT_FIELDS = ['交易日','股票代號','公司名稱','前一交易日','前日水溫','當日水溫','水溫變化',
                '市值排名','市值占大盤比重(%)']
MARKET_FIELDS = ['交易日','追蹤檔數','可判定檔數','資料不足檔數','滾水區','熱水區','溫水區',
                 '冷水區','冰水區','過渡區','滾水占有效樣本(%)','熱水占有效樣本(%)',
                 '溫水占有效樣本(%)','冷水占有效樣本(%)','冰水占有效樣本(%)',
                 '過渡占有效樣本(%)','滾水占樣本市值權重(%)','熱水占樣本市值權重(%)',
                 '溫水占樣本市值權重(%)','冷水占樣本市值權重(%)',
                 '冰水占樣本市值權重(%)','過渡占樣本市值權重(%)']


class RankingTablesParser(HTMLParser):
    """Read actual HTML <table>/<tr>/<td> structure, keeping unrelated tables apart.

    The TAIFEX page presents two side-by-side securities per row.  A global
    regex across the entire page can accidentally treat unrelated figures as
    the same rank (e.g. two different 'rank 1' entries).
    """
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self.stack = []

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == 'table':
            table = {'rows': [], 'row': None, 'cell': None}
            self.tables.append(table)
            self.stack.append(table)
            return
        if not self.stack:
            return
        table = self.stack[-1]
        if tag == 'tr':
            table['row'] = []
        elif tag in ('td','th') and table['row'] is not None:
            table['cell'] = []
        elif tag == 'br' and table['cell'] is not None:
            table['cell'].append(' ')

    def handle_data(self, data):
        if self.stack and self.stack[-1]['cell'] is not None:
            self.stack[-1]['cell'].append(data)

    def handle_endtag(self, tag):
        tag = tag.lower()
        if not self.stack:
            return
        table = self.stack[-1]
        if tag in ('td','th') and table['cell'] is not None:
            table['row'].append(' '.join(''.join(table['cell']).split()))
            table['cell'] = None
        elif tag == 'tr' and table['row'] is not None:
            if table['row']:
                table['rows'].append(table['row'])
            table['row'] = None
        elif tag == 'table':
            self.stack.pop()


_NUMBER = re.compile(r'\d{1,4}\Z')
_CODE = re.compile(r'\d{4}\Z')
_WEIGHT = re.compile(r'(\d+(?:\.\d+)?)\s*%\Z')
_CODE_NAME = re.compile(r'(\d{4})\s+(.+)\Z')
_FULL_ENTRY = re.compile(r'(\d{1,4})\s+(\d{4})\s+(.+?)\s+(\d+(?:\.\d+)?)\s*%\Z')


def _rank_entry(rank, code, name, weight):
    rank, code, name, weight = str(rank).strip(), str(code).strip(), str(name).strip(), str(weight).strip()
    if not (_NUMBER.fullmatch(rank) and _CODE.fullmatch(code) and _WEIGHT.fullmatch(weight)):
        return None
    if not name or re.fullmatch(r'[\d\s.%]+', name):
        return None
    return {'rank': int(rank), 'code': code, 'name': name,
            'weight': float(_WEIGHT.fullmatch(weight).group(1))}


def _entries_in_row(cells):
    """TAIFEX uses two 4-cell groups; tolerate a 3-cell code/name layout."""
    entries = []
    # 4 cells: [rank, code, company, weight] x 2.
    for i in range(len(cells)-3):
        e = _rank_entry(*cells[i:i+4])
        if e:
            entries.append(e)
    if entries:
        return entries
    # 3 cells: [rank, 'code company', weight] x 2.
    for i in range(len(cells)-2):
        pair = _CODE_NAME.fullmatch(cells[i+1])
        if pair:
            e = _rank_entry(cells[i],pair.group(1),pair.group(2),cells[i+2])
            if e:
                entries.append(e)
    if entries:
        return entries
    # Some responsive templates put each whole security in a single cell.
    for cell in cells:
        m = _FULL_ENTRY.fullmatch(cell)
        if m:
            e = _rank_entry(*m.groups())
            if e:
                entries.append(e)
    return entries


def _ranking_from_entries(entries):
    top = {}
    conflicts = []
    for e in entries:
        r = e['rank']
        if 1 <= r <= 100:
            if r in top and top[r]['code'] != e['code']:
                conflicts.append((r,top[r]['code'],e['code']))
            else:
                top.setdefault(r,e)
    if conflicts:
        return None, '同一表格有重複排名：'+str(conflicts[:3])
    missing = sorted(set(range(1,101))-set(top))
    if missing:
        return None, f'只有 {len(top)}/100 檔，缺少 {missing[:8]}'
    ranking=[top[i] for i in range(1,101)]
    if len({e['code'] for e in ranking}) != 100:
        return None,'股票代號重複'
    if any(e['weight'] <= 0 for e in ranking):
        return None,'市值比重為零或負數'
    if any(ranking[i]['weight'] > ranking[i-1]['weight'] + 1e-8 for i in range(1,100)):
        return None,'市值比重沒有依排名遞減'
    return ranking,''


def parse_taifex(html_text):
    parser = RankingTablesParser()
    parser.feed(html_text)
    candidates=[]
    inspected=[]
    for index, table in enumerate(parser.tables):
        entries=[e for cells in table['rows'] for e in _entries_in_row(cells)]
        if not entries:
            continue
        ranking,why=_ranking_from_entries(entries)
        names=' '.join(' '.join(row) for row in table['rows'][:5])
        header_score=sum(key in names for key in ('排行','證券名稱','市值'))
        if ranking:
            candidates.append((header_score,len(entries),-index,ranking))
        else:
            inspected.append(f'table {index+1}: {why}')
    if candidates:
        # Prefer the table explicitly marked as a market-cap ranking.
        # Never merge unrelated tables; this caused the original rank=1 collision.
        candidates.sort(reverse=True, key=lambda a:(a[0],a[1],a[2]))
        return candidates[0][3]
    detail='；'.join(inspected[:5])
    if not parser.tables:
        detail='找不到HTML表格；可能網站改版或網站返回驗證頁'
    raise ValueError('無法取得完整的期交所市值前100大名單（'+detail+'）。已停止，不會用錯誤名單繼續抓價')


def fetch_ranking(debug_output=None):
    headers = {**core.HEADERS, 'Accept-Language':'zh-TW,zh;q=0.9'}
    request = urllib.request.Request(TAIFEX_URL, headers=headers)
    with urllib.request.urlopen(request, timeout=35) as response:
        payload = response.read()
        encoding = response.headers.get_content_charset() or 'utf-8'
    try:
        page = payload.decode(encoding)
    except (LookupError, UnicodeError):
        page = payload.decode('utf-8', errors='replace')
    try:
        ranking = parse_taifex(page)
    except ValueError as e:
        # Save the actual returned HTML for accurate follow-up fixes.
        if debug_output is not None:
            dst=Path(debug_output)/'data'/'taifex_rank_debug.html'
            dst.parent.mkdir(parents=True,exist_ok=True)
            dst.write_bytes(payload)
            raise ValueError(f'{e}。網頁原始資料已保存：{dst}') from e
        raise
    return ranking, dt.datetime.now(TZ).isoformat(timespec='seconds')


def read_csv(path):
    if not path.exists():
        return []
    with path.open('r',encoding='utf-8-sig',newline='') as f:
        return list(csv.DictReader(f))


def write_csv(path, fields, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix('.partial')
    with temp.open('w',encoding='utf-8-sig',newline='') as f:
        wr = csv.DictWriter(f,fieldnames=fields,extrasaction='ignore')
        wr.writeheader()
        wr.writerows(rows)
    temp.replace(path)


def date_today():
    return dt.datetime.now(TZ).date().isoformat()


def save_ranking(output, ranking, fetched):
    store = output/'data'
    store.mkdir(exist_ok=True,parents=True)
    target = store/'latest_ranking.json'
    previous = {}
    if target.exists():
        try:
            previous = json.loads(target.read_text(encoding='utf-8'))
        except (ValueError, OSError):
            previous = {}
    old = {r['code']:r for r in previous.get('items',[])}
    new = {r['code']:r for r in ranking}
    changes = []
    if old:
        for code, r in new.items():
            if code not in old:
                changes.append({'日期':fetched[:10],'異動':'新進前100','股票代號':code,'公司名稱':r['name'],
                                '原排名':'','新排名':r['rank']})
            elif r['rank'] != old[code]['rank']:
                changes.append({'日期':fetched[:10],'異動':'排名變化','股票代號':code,'公司名稱':r['name'],
                                '原排名':old[code]['rank'],'新排名':r['rank']})
        for code,r in old.items():
            if code not in new:
                changes.append({'日期':fetched[:10],'異動':'退出前100','股票代號':code,'公司名稱':r['name'],
                                '原排名':r['rank'],'新排名':''})
    else:
        changes.append({'日期':fetched[:10],'異動':'首次建立名單','股票代號':'','公司名稱':'',
                        '原排名':'','新排名':''})
    fields = ['日期','異動','股票代號','公司名稱','原排名','新排名']
    changefile=store/'排行榜變動歷史.csv'
    earlier=read_csv(changefile)
    entries={(r['日期'],r['異動'],r['股票代號']):r for r in earlier}
    for r in changes:
        entries[(r['日期'],r['異動'],r['股票代號'])] = r
    write_csv(changefile,fields,sorted(entries.values(),key=lambda r:(r['日期'],r['股票代號'],r['異動'])))
    # 優先保存新清單，當股價下載失敗也保留最新排名來源。
    target.write_text(json.dumps({'url':TAIFEX_URL,'fetched_at':fetched,'items':ranking},
                                  ensure_ascii=False,indent=2),encoding='utf-8')
    return changes


def clean_history_flexible(result):
    """沿用五檔版的Adj Close還原法；新上市不足300日仍標記資料不足而非下載失敗。"""
    ind=result.get('indicators',{})
    quote=(ind.get('quote') or [{}])[0]
    adjusted=(ind.get('adjclose') or [{}])[0]
    timestamps=result.get('timestamp') or []
    closes=quote.get('close') or []
    adjs=adjusted.get('adjclose') or []
    if not timestamps or len(timestamps)!=len(closes) or len(adjs)!=len(timestamps):
        raise ValueError('Adj Close 或原始收盤價不完整（不以未還原價格冒充）')
    records={}
    for epoch,raw,adj in zip(timestamps,closes,adjs):
        if raw is None or adj is None or raw<=0 or adj<=0:
            continue
        day=dt.datetime.fromtimestamp(int(epoch),TZ).date().isoformat()
        records[day]={'date':day,'raw_close':float(raw),'normalized_close':float(adj)}
    sequence=[records[d] for d in sorted(records)]
    if len(sequence)<2:
        raise ValueError('有效日收盤價不足2筆')
    factor=sequence[-1]['raw_close']/sequence[-1]['normalized_close']
    for row in sequence:
        row['normalized_close']=round(row['normalized_close']*factor,6)
        row['raw_close']=round(row['raw_close'],6)
    return sequence


def download_one(item):
    code=item['code']
    try:
        result,url=core.yahoo_chart(code, retries=2)
        hist=clean_history_flexible(result)
        return code,hist,url,''
    except Exception as exc:
        return code,[],f'https://finance.yahoo.com/quote/{code}.TW/history/',str(exc)[:260]


def fetch_all(ranking, workers):
    results={}
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures={executor.submit(download_one, item):item['code'] for item in ranking}
        for index,future in enumerate(as_completed(futures),1):
            code, history, url, error=future.result()
            results[code]={'history':history,'url':url,'error':error}
            print(f'  {index:3d}/{len(ranking)} {code}：'+
                  (f'{len(history)}筆，最新 {history[-1]["date"]}' if history else f'失敗：{error[:90]}'),flush=True)
    return results


def analyze_asof(history, index, near):
    """截至某交易日計算，不使用該日之後股價，回測不偷看未來。"""
    truncated=history[:index+1]
    if not truncated:
        raise ValueError('無行情')
    last=truncated[-1]
    # 原5檔輸出中的normalized_close對齊最後一日；回看歷史時必須重新對齊「當時」收盤價。
    normalize=float(last['raw_close'])/float(last['normalized_close'])
    day=[float(r['normalized_close'])*normalize for r in truncated]
    # period_closes只使用截至該日的資料，不含未來週/月收盤價。
    weekly=core.period_closes([{'date':r['date'],'normalized_close':float(r['normalized_close'])*normalize}
                                for r in truncated],'week')
    monthly=core.period_closes([{'date':r['date'],'normalized_close':float(r['normalized_close'])*normalize}
                                 for r in truncated],'month')
    if len(day)<11 or len(weekly)<11 or len(monthly)<11:
        raise ValueError(f'歷史期間不足（{len(day)}日／{len(weekly)}週／{len(monthly)}月）')
    dm,dp,dd=core.moving_stats(day)
    wm,wp,wd=core.moving_stats(weekly)
    mm,mp,md=core.moving_stats(monthly)
    close=float(last['raw_close'])
    zone,reason=core.classify(close,dm,wm,mm,dd,wd,md,near)
    return {'date':last['date'],'price':close,'d':dm,'w':wm,'m':mm,
            'dd':dd,'wd':wd,'md':md,'zone':zone,'reason':reason,
            'd_pct':100*(close/dm-1),'w_pct':100*(close/wm-1),'m_pct':100*(close/mm-1)}


def blank_detail(r, fetched, latest, state, why, url=''):
    out={field:'' for field in DETAIL_FIELDS}
    out.update({'市值排名':r['rank'],'股票代號':r['code'],'公司名稱':r['name'],
                '市值占大盤比重(%)':r['weight'],'今日水溫':state,'資料狀態':state,
                '判定理由':why,'股價來源URL':url})
    return out


def history_snapshots(ranking, fetched, downloaded, backfill):
    """產出最近 N 個市場交易日的水溫快照，並排除個股未交易日期。"""
    market_calendar=Counter()
    for result in downloaded.values():
        for r in result['history']:
            market_calendar[r['date']]+=1
    cutoff=max(1,min(len(ranking)//3,20))
    dates=sorted(d for d,n in market_calendar.items() if n>=cutoff)[-backfill:]
    date_set=set(dates)
    snaps=[]
    for r in ranking:
        info=downloaded[r['code']]
        for idx,price in enumerate(info['history']):
            if price['date'] not in date_set:
                continue
            try:
                m=analyze_asof(info['history'],idx,core.MA_NEAR_PCT)
                category=m['zone']
                dd,wd,md=m['dd'],m['wd'],m['md']
            except ValueError:
                category='資料不足'
                dd=wd=md=''
            snaps.append({'交易日':price['date'],'市值排名':r['rank'],
                          '股票代號':r['code'],'公司名稱':r['name'],
                          '市值占大盤比重(%)':r['weight'],'收盤價':price['raw_close'],
                          '水溫':category,'10日方向':dd,'10週方向':wd,'10月方向':md,
                          '排名抓取時間':fetched,'股價來源URL':info['url']})
    return dates,snaps


def merge_history(out, newrows, refresh=False):
    path=out/'data'/'每日個股水溫歷史.csv'
    old=read_csv(path)
    d={(r['交易日'],r['股票代號']):r for r in old}
    today=max((r['交易日'] for r in newrows), default='')
    # 同一交易日重新執行時，以這一次完整擷取的可用快照重建該日，
    # 避免把之前成功、這次卻失敗的代號當成這次也成功（亦避免名單更替造成>100檔）。
    if refresh and today:
        d={key:row for key,row in d.items() if key[0]!=today}
    # 曾經產出的其他歷史日期保留當時採用的名單，避免現在的100大回填改寫過去成分。
    for row in newrows:
        key=(row['交易日'],row['股票代號'])
        if key not in d or (refresh and row['交易日']==today):
            d[key]=row
    merged=sorted(d.values(),key=lambda r:(r['交易日'],int(r['市值排名'])))
    write_csv(path,HIST_FIELDS,merged)
    return merged


def change_label(before,after):
    if not before:
        return '首次觀測'
    if before==after:
        return '維持不變'
    if before in GRADE and after in GRADE:
        return '升溫' if GRADE[after]>GRADE[before] else '降溫'
    if after=='過渡區':
        return '轉入過渡區'
    if before=='過渡區':
        return '離開過渡區'
    return '分類改變'


def build_transitions(history):
    by_code=defaultdict(list)
    for r in history:
        if r['水溫'] in CATEGORIES:
            by_code[r['股票代號']].append(r)
    events=[]; prevmap={}
    for code,seq in by_code.items():
        seq.sort(key=lambda a:a['交易日'])
        for i,row in enumerate(seq):
            before=seq[i-1] if i else None
            label=change_label(before['水溫'] if before else '',row['水溫'])
            entry={'交易日':row['交易日'],'股票代號':code,'公司名稱':row['公司名稱'],
                   '前一交易日':before['交易日'] if before else '',
                   '前日水溫':before['水溫'] if before else '',
                   '當日水溫':row['水溫'],'水溫變化':label,
                   '市值排名':row['市值排名'],'市值占大盤比重(%)':row['市值占大盤比重(%)']}
            if i and label != '維持不變':
                events.append(entry)
            prevmap[(row['交易日'],code)]=entry
    return sorted(events,key=lambda r:(r['交易日'],int(r['市值排名']))),prevmap


def build_market(history):
    by_date=defaultdict(list)
    for r in history:
        by_date[r['交易日']].append(r)
    output=[]
    for date,rows in sorted(by_date.items()):
        counts=Counter(r['水溫'] for r in rows)
        valid=[r for r in rows if r['水溫'] in CATEGORIES]
        valid_n=len(valid)
        weight_total=sum(float(r['市值占大盤比重(%)']) for r in valid)
        line={'交易日':date,'追蹤檔數':len(rows),'可判定檔數':valid_n,
              '資料不足檔數':counts['資料不足']}
        for t in CATEGORIES:
            line[t]=counts[t]
            line[t.replace('區','')+'占有效樣本(%)']=round(counts[t]/valid_n*100,2) if valid_n else 0
            line[t.replace('區','')+'占樣本市值權重(%)']=round(sum(float(r['市值占大盤比重(%)'])
                for r in valid if r['水溫']==t)/weight_total*100,2) if weight_total else 0
        output.append(line)
    return output


def latest_details(ranking, downloaded, fetched, market_latest, prevmap):
    details=[]
    for r in ranking:
        info=downloaded[r['code']]
        hist=info['history']
        if not hist:
            details.append(blank_detail(r,fetched,market_latest,'資料異常',info['error'],info['url']))
            continue
        latest_day=hist[-1]['date']
        if latest_day!=market_latest:
            details.append(blank_detail(r,fetched,market_latest,'資料異常',
                     f'行情較舊：個股 {latest_day}，本次市場基準日 {market_latest}',info['url']))
            continue
        try:
            m=analyze_asof(hist,len(hist)-1,core.MA_NEAR_PCT)
        except ValueError as exc:
            x=blank_detail(r,fetched,market_latest,'資料不足',str(exc),info['url'])
            x['行情日期']=latest_day;x['最新收盤價(元)']=hist[-1]['raw_close']
            details.append(x);continue
        prev=prevmap.get((latest_day,r['code']),{})
        details.append({'市值排名':r['rank'],'股票代號':r['code'],'公司名稱':r['name'],
             '市值占大盤比重(%)':r['weight'],'行情日期':m['date'],
             '最新收盤價(元)':round(m['price'],4),'10日均線(元)':round(m['d'],4),
             '10週均線(元)':round(m['w'],4),'10月均線(元)':round(m['m'],4),
             '10日方向':m['dd'],'10週方向':m['wd'],'10月方向':m['md'],
             '方向組合(日/週/月)':f"{m['dd']}/{m['wd']}/{m['md']}",
             '距10日線(%)':round(m['d_pct'],2),'距10週線(%)':round(m['w_pct'],2),
             '距10月線(%)':round(m['m_pct'],2),'今日水溫':m['zone'],
             '前一交易日':prev.get('前一交易日',''), '前日水溫':prev.get('前日水溫',''),
             '水溫變化':prev.get('水溫變化','首次觀測'), '資料狀態':'正常',
             '判定理由':m['reason'],'股價來源URL':info['url']})
    return details


def xlsx_write(path, sheets):
    """延續已在五檔版成功執行的標準函式庫 Excel OOXML 產檔方式。"""
    esc=core.escape
    content='''<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'''
    for i in range(1,len(sheets)+1):
        content+=f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    content+='</Types>'
    rels=('<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    work=('<?xml version="1.0" encoding="utf-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+
           ''.join(f'<sheet name="{esc(a[0])}" sheetId="{i}" r:id="rId{i}"/>' for i,a in enumerate(sheets,1))+'</sheets></workbook>')
    bookrels=('<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
      ''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,len(sheets)+1))+
      f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
    styles='''<?xml version="1.0" encoding="utf-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
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
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''
    path.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(path,'w',compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',content)
        z.writestr('_rels/.rels',rels)
        z.writestr('xl/workbook.xml',work)
        z.writestr('xl/_rels/workbook.xml.rels',bookrels)
        z.writestr('xl/styles.xml',styles)
        for i,(_,rows,widths,freeze,filter_row) in enumerate(sheets,1):
            z.writestr(f'xl/worksheets/sheet{i}.xml',core.excel_sheet_xml(rows,widths,freeze,filter_row,title_rows=(1,)))


def export_html(path,details,market,latest,fetched):
    cards=[]
    today=next((r for r in market if r['交易日']==latest),None)
    for zone in CATEGORIES:
        num=int(today.get(zone,0)) if today else 0
        cards.append(f'<div class="card"><span style="background:{COLORS[zone]}" class="dot"></span>'
                     f'<span>{zone}</span><strong>{num}</strong></div>')
    cols=['市值排名','股票代號','公司名稱','最新收盤價(元)','10日方向','10週方向','10月方向','方向組合(日/週/月)','今日水溫','前日水溫','水溫變化','資料狀態']
    def cell(text,k):
        v=html_escape(str(text))
        if k=='今日水溫' and str(text) in COLORS:
            return f'<td><b class="tag" style="background:{COLORS[str(text)]}">{v}</b></td>'
        return f'<td>{v}</td>'
    trs=''.join('<tr>'+''.join(cell(r.get(k,''),k) for k in cols)+'</tr>' for r in details)
    options=''.join(f'<option value="{s}">{s}</option>' for s in CATEGORIES+['資料異常','資料不足'])
    title='上市市值前100大' if len(details)==100 else '五檔離線驗證【非即時】'
    html=f'''<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>前100大五水溫儀表板</title><style>
body{{font-family:"Microsoft JhengHei","Noto Sans TC",sans-serif;background:#f4f7fb;color:#172b3f;margin:0}}main{{max-width:1540px;margin:auto;padding:24px}}h1{{margin:0;font-size:25px}}.muted{{color:#6c7b8b;font-size:13px;line-height:1.7}}.cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(175px,1fr));gap:10px;margin:18px 0}}.card{{background:white;border:1px solid #e1e8ef;border-radius:12px;padding:17px;display:flex;align-items:center;gap:10px}}.card strong{{margin-left:auto;font-size:25px}}.dot{{height:11px;width:11px;border-radius:20px}}.panel{{background:white;border-radius:12px;padding:17px;border:1px solid #e1e8ef}}input,select{{padding:10px;border:1px solid #d2dce8;border-radius:7px;margin:0 12px 15px 0}}.scroll{{max-height:75vh;overflow:auto}}table{{border-collapse:collapse;width:100%;min-width:1250px;font-size:13px}}th,td{{padding:11px;border-bottom:1px solid #edf0f4;text-align:left;white-space:nowrap}}th{{position:sticky;top:0;background:#18384f;color:white;cursor:pointer}}tbody tr:hover{{background:#f1f7ff}}.tag{{border-radius:15px;padding:4px 9px;color:#172b3f}}.meta{{margin-top:12px}}</style></head><body><main>
<h1>{title}｜五水溫趨勢追蹤</h1><div class="muted">行情基準日：{latest}｜排名抓取：{fetched}｜未取得之行情會顯示資料異常｜貼近均線門檻 ±2%</div>
<div class="cards">{''.join(cards)}</div><div class="panel"><input id="q" placeholder="搜尋股票名稱／代號" oninput="filter()"><select id="temp" onchange="filter()"><option value="">全部水溫</option>{options}</select><span id="counter"></span><div class="scroll"><table id="stocks"><thead><tr>{''.join('<th>'+x+'</th>' for x in cols)}</tr></thead><tbody>{trs}</tbody></table></div></div>
<p class="muted meta">市值排行榜：{TAIFEX_URL}<br>股價來源：Yahoo Finance；日線、當期週線及月線以當時收盤價計算，歷史價格採調整收盤價重基準。市場統計僅含可判定且與市場基準日一致之股票；不是投資建議。</p>
<script>let idx=-1,desc=false;function filter(){{let q=document.getElementById('q').value.toLowerCase(),t=document.getElementById('temp').value,n=0;for(let r of document.querySelectorAll('#stocks tbody tr')){{let show=r.textContent.toLowerCase().includes(q)&&(!t||r.cells[8].textContent===t);r.style.display=show?'':'none';if(show)n++}}document.getElementById('counter').textContent=n+' / {len(details)} 檔';}}
for(let [i,h] of [...document.querySelectorAll('#stocks th')].entries()){{h.onclick=()=>{{desc=idx===i?!desc:false;idx=i;let t=document.querySelector('#stocks tbody'),rs=[...t.rows];rs.sort((a,b)=>{{let x=a.cells[i].textContent,y=b.cells[i].textContent,nx=Number(x.replace(/,/g,'')),ny=Number(y.replace(/,/g,''));return (x!==''&&y!==''&&!isNaN(nx)&&!isNaN(ny)?nx-ny:x.localeCompare(y,'zh-Hant'))*(desc?-1:1)}});rs.forEach(r=>t.appendChild(r));}}}}filter();</script></main></body></html>'''
    path.write_text(html,encoding='utf-8')


def export_all(out,ranking,fetched,downloaded,backfill,rank_changes):
    out.mkdir(exist_ok=True,parents=True)
    valid_dates=[v['history'][-1]['date'] for v in downloaded.values() if v['history']]
    if not valid_dates:
        raise RuntimeError('100檔股價全部擷取失敗，停止產生報表；名單已留存，請檢查網路')
    frequency=Counter(valid_dates)
    latest=max(frequency, key=lambda d:(frequency[d],d))
    days,snapshots=history_snapshots(ranking,fetched,downloaded,backfill)
    merged=merge_history(out,snapshots,refresh=True)
    transitions,prevmap=build_transitions(merged)
    market=build_market(merged)
    details=latest_details(ranking,downloaded,fetched,latest,prevmap)
    today_events=[r for r in transitions if r['交易日']==latest]
    counter=Counter(r['今日水溫'] for r in details)
    denom=sum(counter[t] for t in CATEGORIES)
    goodweights=sum(float(r['市值占大盤比重(%)']) for r in details if r['今日水溫'] in CATEGORIES)
    display_title="市值前100大" if len(ranking)==100 else "五檔離線驗證【非100檔即時行情】"
    summary=[[f'{display_title}｜五水溫總覽',f'股價基準日：{latest}',f'排名抓取：{fetched}'],
             ['10日／10週／10月均線＋方向；週/月線含當期；貼近 ±2%；交易日之間比較；非即時報價'],
             ['水溫','當日檔數','有效樣本占比(%)','樣本市值加權占比(%)','定義/註解']]
    for t in CATEGORIES:
        n=counter[t]
        share=round(n/denom*100,2) if denom else 0
        weight=round(sum(float(r['市值占大盤比重(%)']) for r in details if r['今日水溫']==t)/goodweights*100,2) if goodweights else 0
        summary.append([t,n,share,weight,'市值占比以當次有效樣本權重正規化；非整體上市大盤占比'])
    summary.extend([['資料不足',counter['資料不足'],'','', '未滿11個月樣本不強制判定'],
                    ['資料異常',counter['資料異常'],'','', '股價缺失、延遲或API失敗'],
                    ['合計',len(details),'','',f'有效 {denom} / {len(details)} 檔'],
                    ['今日有變化',len(today_events),'','', '只計最近一個市場交易日'],
                    ['前100大合計市值占大盤比重(%)',round(sum(r['weight'] for r in ranking),4)],
                    ['市值統計說明','加權占比：當水溫類別公司之大盤市值比重合計 ÷ 所有有效樣本公司之大盤市值比重合計 × 100'],
                    ['歷史統計說明','首次回填用「當次100大成分」推算過去各日；之後保留每次執行時的當期成分，避免未來名單改寫舊紀錄']])
    detailrows=[[f'{display_title}｜股票明細'],[f'行情基準 {latest}｜每次執行時重抓排名及股價；來源追溯欄位在最右側'],DETAIL_FIELDS]
    detailrows.extend([[r.get(f,'') for f in DETAIL_FIELDS] for r in details])
    events_rows=[['最近交易日異動｜只顯示水溫變化'],[f'最近市場交易日 {latest}；「過渡區」沒有固定強弱排序'],EVENT_FIELDS]
    events_rows.extend([[r.get(f,'') for f in EVENT_FIELDS] for r in today_events])
    if not today_events:events_rows.append(['沒有已觀測到的水溫異動'])
    trendrows=[['市場水溫日序列｜保留歷次執行所得歷史'],['統計母體以當日追蹤檔數為準；初次回填為本次100大樣本'],MARKET_FIELDS]
    trendrows.extend([[r.get(f,'') for f in MARKET_FIELDS] for r in market[-120:]])
    historyrows=[['每日個股水溫歷史｜最近交易日'],['完整歷史另存 data/每日個股水溫歷史.csv'],HIST_FIELDS]
    latest_20=set(sorted({r['交易日'] for r in merged})[-20:])
    historyrows.extend([[r.get(f,'') for f in HIST_FIELDS] for r in merged if r['交易日'] in latest_20])
    rankfields=['日期','異動','股票代號','公司名稱','原排名','新排名']
    rankrows=[[f'{display_title}｜名單異動'],['比較前次成功擷取之排名；初次執行僅顯示首次建立名單'],rankfields]
    rankrows.extend([[r.get(f,'') for f in rankfields] for r in rank_changes])
    rules=[['規則與資料來源｜V2'],
      ['資料母體','期交所公布之臺灣上市公司市值占比排名前100名'],
      ['名單資料','每次執行時重抓100大；來源網頁的發布日期不一定等於擷取日期'],
      ['股價來源','Yahoo Finance Chart API，依個股 .TW 代號抓最近2年歷史價'],
      ['股價整理','Adj Close 重基準至當日原始收盤價；與證券商採原始K線比較可能不同'],
      ['10日線','最近10個交易日調整價平均'],
      ['10週線','最近10個週收盤調整價平均，含當週截至目前交易日'],
      ['10月線','最近10個月收盤調整價平均，含當月截至目前交易日'],
      ['方向','當期MA較上一期MA高=上揚，較低=下彎（各自以日、週、月為一期）'],
      ['滾水區','10日/10週/10月均線全部上揚'],
      ['冰水區','10日/10週/10月均線全部下彎'],
      ['溫水區','股價在10月均線 ±2%以內且10日/10週線上揚'],
      ['熱水區','股價高於10月線超過2%；10週線上揚；距10日線 ±2%內'],
      ['冷水區','股價低於10月線超過2%；10週線上揚；距10日線 ±2%內'],
      ['過渡區','以上均不符合，不強制歸類'],
      ['優先順序','滾水 → 冰水 → 溫水 → 熱水 → 冷水 → 過渡區'],
      ['比較日','每一檔股票本身最近兩個有效交易日；若缺資料會標註'],
      ['資料不足','短期上市等，無法取得至少11期月均線所需歷史'],
      ['資料異常','Yahoo取價失敗或最後成交日不是本次股價基準日'],
      ['歷史留存','每天個股水溫、各區數量及異動記錄會保留於 water_output/data'],
      ['指數權重','各區成分市值比重合計/有效樣本市值比重合計，非加權指數日報酬'],
      ['期交所',TAIFEX_URL],['Yahoo','https://finance.yahoo.com/'],
      ['重要限制','本工具僅作趨勢觀測，非估值模型或買賣建議；Yahoo可能限流/改版']]
    xlsx_write(out/'100大水溫監測.xlsx',[
        ('市場水溫儀表板',summary,[24,26,24,31,60],3,3),
        ('100檔股票明細',detailrows,[11,13,20,23,16,17,17,17,17,13,13,13,24,17,17,17,13,18,13,17,16,45,83],3,3),
        ('最近水溫異動',events_rows,[16,15,22,18,16,16,18,13,22],3,3),
        ('市場水溫日序列',trendrows,[17]+[20]*(len(MARKET_FIELDS)-1),3,3),
        ('個股水溫歷史',historyrows,[18,13,13,23,24,16,15,13,13,13,27,79],3,3),
        ('前100大名單變化',rankrows,[17,18,15,20,13,13],3,3),
        ('規則與資料來源',rules,[27,105],1,None)])
    write_csv(out/'100大水溫總覽.csv',DETAIL_FIELDS,details)
    write_csv(out/'最近水溫異動.csv',EVENT_FIELDS,today_events)
    write_csv(out/'市場水溫日序列.csv',MARKET_FIELDS,market)
    write_csv(out/'排行榜變化.csv',['日期','異動','股票代號','公司名稱','原排名','新排名'],rank_changes)
    write_csv(out/'data'/'水溫異動歷史.csv',EVENT_FIELDS,transitions)
    export_html(out/'100大水溫儀表板.html',details,market,latest,fetched)
    print(f'\n完成：最新行情 {latest}｜有效 {denom}/{len(details)}｜今日水溫異動 {len(today_events)} 筆')
    print('  '+str((out/'100大水溫監測.xlsx').resolve()))
    print('  '+str((out/'100大水溫儀表板.html').resolve()))
    print('  歷史紀錄：'+str((out/'data').resolve()))
    if len(ranking) != 100:
        print('本次為五檔歷史資料的離線驗證，權重各設1.0僅為檢驗計算，不代表真實市值分布。')
    if counter['資料異常'] or counter['資料不足']:
        print('注意：資料不足 %d 檔，資料異常 %d 檔；不納入有效水溫分布。' %
              (counter['資料不足'],counter['資料異常']))
    return 0 if denom>=0.9*len(ranking) else 2


def offline_fixture(path):
    """只用使用者先前5檔有效歷史資料驗證輸出；不冒充100檔或最新盤價。"""
    rows=read_csv(path)
    if not rows:
        raise RuntimeError('沒有離線測試用歷史資料')
    groups=defaultdict(list)
    names={}
    for r in rows:
        code=r['股票代號']; names[code]=r['股票名稱']
        groups[code].append({'date':r['日期'],'raw_close':float(r['實際收盤價(元)']),
                            'normalized_close':float(r['還原且對齊最新價(元)'])})
    priority=['2330','2454','2308','2317','3711']
    ranking=[{'rank':i,'code':code,'name':names[code],'weight':1.0} for i,code in enumerate(priority,1)]
    downloaded={c:{'history':sorted(groups[c],key=lambda r:r['date']),
                   'url':'離線檢驗：使用者上傳之2026-11-16五檔水溫監測.xlsx','error':''} for c in priority}
    return ranking,downloaded


def self_test():
    # Test actual side-by-side layout and ensure a separate table with another
    # rank=1 cannot corrupt the market-cap list (the user's reported bug).
    inner=''.join('<tr><td>%d</td><td>%04d</td><td>測試甲</td><td>%.4f%%</td>'
                  '<td>%d</td><td>%04d</td><td>測試乙</td><td>0.001%%</td></tr>'
                  % (i,2000+i,10-i*0.01,i+550,4000+i) for i in range(1,101))
    table='<table><tr><th>排行</th><th>證券名稱</th><th>市值</th></tr>'+inner+'</table>'
    other='<table><tr><td>1</td><td>8888</td><td>另一張表</td><td>8.9%</td></tr></table>'
    parsed=parse_taifex('<html>'+other+table+other+'</html>')
    assert len(parsed)==100 and parsed[0]['code']=='2001' and parsed[-1]['code']=='2100'
    # Variant where the security code and name are one cell.
    threes='<table>'+''.join('<tr><td>%d</td><td>%04d 測試甲</td><td>%.4f%%</td>'
                '<td>%d</td><td>%04d 測試乙</td><td>0.001%%</td></tr>'
                % (i,2000+i,10-i*0.01,i+550,4000+i) for i in range(1,101))+'</table>'
    assert len(parse_taifex(threes))==100
    try:
        parse_taifex('<table><tr><td>1</td><td>2330</td><td>台積電</td><td>41.6%</td></tr></table>')
        raise AssertionError('應拒絕只有1檔的排行榜')
    except ValueError as exc:
        assert '只有 1/100' in str(exc)
    core.self_test()
    assert change_label('冰水區','冷水區')=='升溫'
    assert change_label('熱水區','過渡區')=='轉入過渡區'
    print('PASS: TAIFEX table-isolated parsing, duplicate-rank protection, 3/4-cell layouts, water temperature rules.')


def main():
    parser=argparse.ArgumentParser(description='上市前100大五水溫趨勢儀表板與每日異動追蹤')
    parser.add_argument('--output',default='water_output',help='預設於程式所在資料夾建立water_output')
    parser.add_argument('--workers',type=int,default=3,choices=range(1,6),metavar='{1..5}')
    parser.add_argument('--backfill',type=int,default=25,help='首次回填最近N個交易日，預設25')
    parser.add_argument('--offline-5',action='store_true',help='用五檔已驗證的歷史CSV離線檢查報表；非100檔實時版')
    parser.add_argument('--self-test',action='store_true')
    args=parser.parse_args()
    if args.self_test:return self_test()
    if not 2<=args.backfill<=120:parser.error('--backfill 須為2至120')
    output=Path(args.output)
    if not output.is_absolute():output=Path(__file__).resolve().parent/output
    if args.offline_5:
        file=Path(__file__).resolve().parent/'reference_5_history.csv'
        ranking,downloaded=offline_fixture(file)
        out=output/'離線五檔驗證'
        print('離線五檔檢驗模式（不是市值前100大即時報告）')
        return export_all(out,ranking,'2026-11-16T16:38:00+08:00',downloaded,args.backfill,
                          [{'日期':'2026-11-16','異動':'五檔離線驗證','股票代號':'','公司名稱':'',
                            '原排名':'','新排名':''}])
    print('讀取期交所市值前100大排名...',flush=True)
    ranking,fetched=fetch_ranking(output)
    print('排行100檔下載完成，正在批次讀取 Yahoo Finance 歷史價格...',flush=True)
    downloaded=fetch_all(ranking,args.workers)
    rank_changes=save_ranking(output,ranking,fetched)
    return export_all(output,ranking,fetched,downloaded,args.backfill,rank_changes)


if __name__=='__main__':
    try:
        sys.exit(main())
    except Exception as ex:
        print('\n執行失敗：'+str(ex),file=sys.stderr)
        sys.exit(1)
