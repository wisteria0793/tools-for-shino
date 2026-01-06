"""
このスクリプトは、Excelファイル (.xlsx) 内の特定セル（送金日・請求月）を
指定された日（または当日）の和暦に更新します。

更新対象セル（先頭シート）:
  - 請求月 (2行目): AN2(元号), AO2(年), AQ2(月)
  - 送金日 (15行目): F15(元号), G15(年), I15(月), K15(日)

使い方:
  python scripts/update_excel_date.py documents/manage
  python scripts/update_excel_date.py documents/manage --date 2026-02-01
"""

import openpyxl
import os
import argparse
import glob
from datetime import date, datetime

# japaneraライブラリをインポート
try:
    from japanera import EraDate
except ImportError:
    print("エラー: 'japanera' ライブラリが見つかりません。")
    print("pip install japanera を実行してください。")
    exit()

def update_excel_file(file_path, target_date):
    print(f"--- Processing: {os.path.basename(file_path)} ---")
    
    try:
        # data_only=False で読み込まないと式が消える可能性があるが、
        # 今回は値を上書きするのでデフォルト(数式保持)でOK
        wb = openpyxl.load_workbook(file_path)
        
        # 先頭のシートのみ対象
        if not wb.worksheets:
            print("  Error: No worksheets found.")
            return

        sheet = wb.worksheets[0]
        modified = False
        
        # 和暦情報の準備
        era_date = EraDate.from_date(target_date)
        era_name = era_date.era.kanji
        # 年を計算
        era_year = era_date.year - era_date.era.since.year + 1
        
        target_month = era_date.month
        target_day = era_date.day
        
        # 更新するセルと値のマッピング
        # (Cell Address, Value, Description)
        updates = [
            # 請求月 (2行目)
            ('AN2', era_name, "Era (Header)"),
            ('AO2', era_year, "Year (Header)"),
            ('AQ2', target_month, "Month (Header)"),
            
            # 送金日 (15行目)
            ('F15', era_name, "Era (Transfer Date)"),
            ('G15', era_year, "Year (Transfer Date)"),
            ('I15', target_month, "Month (Transfer Date)"),
            ('K15', target_day, "Day (Transfer Date)"),
        ]
        
        for cell_addr, value, desc in updates:
            try:
                sheet[cell_addr].value = value
                print(f"  Updated {cell_addr} ({desc}): {value}")
                modified = True
            except Exception as e:
                print(f"  Warning: Could not update {cell_addr}: {e}")
        
        if modified:
            wb.save(file_path)
            print("  Saved changes.")
            
    except Exception as e:
        print(f"  Error: {e}")

def main():
    parser = argparse.ArgumentParser(description='Excelファイル内の送金日・請求月（固定セル）を更新します。')
    parser.add_argument('path', help='対象のディレクトリまたはファイルパス')
    parser.add_argument('--date', help='設定する日付 (YYYY-MM-DD)。省略時は当日。', default=None)
    
    args = parser.parse_args()
    
    # 日付設定
    if args.date:
        try:
            target_date = datetime.strptime(args.date, '%Y-%m-%d').date()
        except ValueError:
            print("エラー: 日付フォーマットが不正です。YYYY-MM-DD で指定してください。")
            return
    else:
        target_date = date.today()
        
    print(f"Setting date to: {target_date.strftime('%Y年%m月%d日')}")

    # ファイルリスト取得
    files = []
    if os.path.isdir(args.path):
        # xlsxのみ対象
        files = glob.glob(os.path.join(args.path, "*.xlsx"))
        # xlsファイルへの警告
        xls_files = glob.glob(os.path.join(args.path, "*.xls"))
        if xls_files:
            print("\n[Warning] 以下の .xls ファイルはサポート外のためスキップされます。.xlsx に変換してください:")
            for f in xls_files:
                print(f"  - {os.path.basename(f)}")
            print("")
    elif os.path.isfile(args.path):
        if args.path.endswith(".xlsx"):
            files = [args.path]
        else:
            print("エラー: .xlsx ファイルのみサポートしています。")
            return
            
    if not files:
        print("処理対象の .xlsx ファイルが見つかりませんでした。")
        return
        
    for file_path in files:
        if not os.path.basename(file_path).startswith("~"):
            update_excel_file(file_path, target_date)
            
    print("\n完了しました。")

if __name__ == "__main__":
    main()
