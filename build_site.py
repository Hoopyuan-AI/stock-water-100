#!/usr/bin/env python3
# Publish only the generated HTML and downloadable Excel, not debugging/source histories.
from pathlib import Path
import shutil, re, sys, zipfile
src=Path('water_output')
out=Path('public_site')
html=src/'100大水溫儀表板.html'
xlsx=src/'100大水溫監測.xlsx'
if not html.is_file() or not xlsx.is_file():
    raise SystemExit('Missing generated dashboard or workbook; site not published.')
text=html.read_text(encoding='utf-8')
if '<html' not in text.lower() or '100' not in text:
    raise SystemExit('HTML verification failed; site not published.')
with zipfile.ZipFile(xlsx) as z:
    if 'xl/workbook.xml' not in z.namelist():
        raise SystemExit('Excel verification failed; site not published.')
out.mkdir(parents=True,exist_ok=True)
shutil.copy2(html,out/'index.html')
shutil.copy2(xlsx,out/'water-report.xlsx')
(out/'.nojekyll').write_text('',encoding='utf-8')
print('PASS: static website generated, index.html + water-report.xlsx')
