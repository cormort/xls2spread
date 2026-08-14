#!/usr/bin/env python3
"""把舊版 .xls 轉成 .xlsx，好丟進 index.html。

瀏覽器讀不了 BIFF 格式的 .xls，這支只做轉檔（用 LibreOffice），排版一律在 index.html 做。
手邊有 Excel 的話，直接「另存新檔 → Excel 活頁簿 (.xlsx)」效果一樣，用不到這支。

    python3 xls2spread.py 固資成本效益分析表.xls [-o 輸出目錄]
"""
import argparse, os, shutil, subprocess, sys

SOFFICE = shutil.which('soffice') or '/Applications/LibreOffice.app/Contents/MacOS/soffice'


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xls')
    ap.add_argument('-o', '--outdir', default=None)
    a = ap.parse_args()

    if not os.path.exists(SOFFICE):
        sys.exit('找不到 LibreOffice（soffice）。請安裝，或直接用 Excel 另存為 .xlsx。')

    outdir = a.outdir or os.path.dirname(os.path.abspath(a.xls))
    subprocess.run([SOFFICE, '--headless', '--convert-to', 'xlsx', '--outdir', outdir, a.xls],
                   check=True, capture_output=True)
    out = os.path.join(outdir, os.path.splitext(os.path.basename(a.xls))[0] + '.xlsx')
    print(out)
    print('把這個檔案拖進 index.html（或 https://cormort.github.io/xls2spread/）就可以排版。')


if __name__ == '__main__':
    main()
