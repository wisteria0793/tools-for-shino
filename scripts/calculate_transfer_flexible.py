import xlwings as xw
import glob
import os
import time

def get_transfer_amount_xlwings(file_path):
    # Appを非表示で起動（これがないとウィンドウが大量に開きます）
    # ※既にExcelが開いている場合、それを利用するか新規インスタンスにするかは設定次第ですが、
    #  安全のため個別に開閉します。
    app = xw.App(visible=False)
    try:
        wb = app.books.open(file_path)
        
        # 1. まず D18 を確認
        # 全シートをループ
        for sheet in wb.sheets:
            try:
                # xlwingsでは range('D18').value で値が取れる（数式も計算後の値になる）
                val = sheet.range('D18').value
                if isinstance(val, (int, float)) and val > 0:
                    wb.close()
                    app.quit()
                    return val
            except:
                pass
        
        # 2. D18で見つからない場合、「送金対象額」を検索
        # xlwingsにはfindメソッドがあるが、API範囲内で行うため全セル探索は重い。
        # used_range を使って効率化する
        for sheet in wb.sheets:
            used_range = sheet.used_range
            # 値を2次元リストとして一括取得（高速化）
            values = used_range.value 
            
            if not values: continue

            # 行番号・列番号のオフセット（used_rangeの開始位置）
            start_row = used_range.row
            start_col = used_range.column
            
            for r_idx, row in enumerate(values):
                for c_idx, cell_val in enumerate(row):
                    if cell_val and isinstance(cell_val, str) and ("送金対象額" in cell_val or "送金額" in cell_val):
                        # 見つかった座標 (1-based index for xlwings cells)
                        # r_idx, c_idx は 0-based なので +1 して、さらに start_row/col を考慮
                        target_row = start_row + r_idx
                        target_col = start_col + c_idx
                        
                        # 周辺セルをチェック (下, 右, 5つ右など)
                        # シートオブジェクトから直接値を取る
                        candidates_coords = [
                            (target_row + 1, target_col),     # 下
                            (target_row, target_col + 1),     # 右
                            (target_row, target_col + 5),     # 5つ右
                            (target_row + 1, target_col + 5), # 下の5つ右
                        ]
                        
                        for r, c in candidates_coords:
                            try:
                                val = sheet.cells(r, c).value
                                if isinstance(val, (int, float)) and val > 0:
                                    wb.close()
                                    app.quit()
                                    return val
                            except:
                                pass
        
        wb.close()
    except Exception as e:
        print(f"Error processing {os.path.basename(file_path)}: {e}")
    finally:
        # プロセスを確実に終了させる
        try:
            app.quit()
        except:
            pass
            
    return 0

def main():
    target_dir = os.path.abspath('documents/manage')
    files = glob.glob(os.path.join(target_dir, "*.xlsx"))
    files.sort()

    bank_totals = {}
    print(f"{ 'ファイル名':<45} | { '銀行':<6} | { '金額':<10}")
    print("-" * 75)

    for f in files:
        filename = os.path.basename(f)
        if filename.startswith("~"):
            continue
        
        # 銀行名の判定
        normalized_name = filename.replace("　", "_").replace(" ", "_")
        if "_" in normalized_name:
            bank_name = normalized_name.split("_")[0]
        else:
            bank_name = "その他"
            
        amount = get_transfer_amount_xlwings(f)
        
        print(f"{filename[:43]:<45} | {bank_name:<6} | {int(amount):>10,}")
        
        if bank_name in bank_totals:
            bank_totals[bank_name] += amount
        else:
            bank_totals[bank_name] = amount

    print("\n" + "="*40)
    print("銀行別送金合計金額")
    print("="*40)
    for bank, total in bank_totals.items():
        print(f"{bank}: {int(total):,} 円")
    print("="*40)

if __name__ == "__main__":
    main()
