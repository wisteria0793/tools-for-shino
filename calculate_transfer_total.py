import pandas as pd
import os
import glob
import re

# 対象ディレクトリ
target_dir = 'documents/manage'
# 銀行ごとの合計を格納する辞書
bank_totals = {}
# エラーファイルリスト
error_files = []

def get_transfer_amount(file_path):
    try:
        # ヘッダーなしで読み込む
        # engine='openpyxl' を明示（xlsxの場合）
        if file_path.lower().endswith('.xlsx'):
            df = pd.read_excel(file_path, header=None, engine='openpyxl')
        else:
            df = pd.read_excel(file_path, header=None)
        
        # "送金額" を含むセルを探す
        for r_idx, row in df.iterrows():
            for c_idx, value in row.items():
                if isinstance(value, str) and "送金対象額" in value:
                    # 見つかったら、その下のセル (r+1, c) の値を取得
                    if r_idx + 1 < len(df):
                        amount_val = df.iat[r_idx + 1, c_idx]
                        
                        # カンマ除去と数値変換
                        try:
                            if isinstance(amount_val, str):
                                amount_val = amount_val.replace(',', '').strip()
                            return float(amount_val)
                        except (ValueError, TypeError):
                            # 数値変換できない場合
                            pass
        
        # 見つからない場合
        return 0
        
    except Exception as e:
        error_files.append((os.path.basename(file_path), str(e)))
        return 0

def main():
    print(f"Scanning directory: {target_dir} ...")
    
    files = glob.glob(os.path.join(target_dir, "*"))
    files.sort()
    
    for file_path in files:
        filename = os.path.basename(file_path)
        
        if filename.startswith("~") or filename.startswith("."):
            continue
            
        if not (filename.lower().endswith(".xlsx") or filename.lower().endswith(".xls")):
            continue
            
        # 銀行名の抽出
        # 全角/半角スペースをアンダースコアに置換してから分割など、揺れを吸収
        normalized_name = filename.replace("　", "_").replace(" ", "_")
        if "_" in normalized_name:
            bank_name = normalized_name.split("_")[0]
        else:
            bank_name = "Others"
        
        bank_name = bank_name.strip()
        
        amount = get_transfer_amount(file_path)
        
        # エラーで0が返ってきた場合は集計しない（0円として扱う）
        # ただしエラーリストに入っているものは別途報告される
        
        if bank_name in bank_totals:
            bank_totals[bank_name] += amount
        else:
            bank_totals[bank_name] = amount
            
        # エラーファイルでない場合のみログ出力
        is_error = any(f[0] == filename for f in error_files)
        if not is_error:
            print(f"Processed: {filename} -> Bank: {bank_name}, Amount: {int(amount):,}")

    print("\n" + "="*40)
    print("銀行別送金合計金額")
    print("="*40)
    for bank, total in bank_totals.items():
        print(f"{bank}: {int(total):,} 円")
    print("="*40)
    
    if error_files:
        print("\n[!] 以下のファイルは読み込みエラー等のため集計に含まれていないか、0円として計算されました:")
        for fname, err in error_files:
            print(f"- {fname}: {err}")

if __name__ == "__main__":
    main()
