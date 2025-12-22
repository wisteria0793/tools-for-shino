"""
このスクリプトは、指定されたファイルまたはディレクトリ内のファイルをmacOSのデフォルトプリンタで印刷します。

使い方:

1.  単一のファイルを印刷する場合:
    python scripts/print_word_document.py /path/to/your/file.docx

2.  ディレクトリ内のすべての.docxファイルを印刷する場合 (デフォルト):
    python scripts/print_word_document.py /path/to/your/folder

3.  ディレクトリ内の特定のパターンのファイルを印刷する場合 (例: .pdfファイル):
    python scripts/print_word_document.py /path/to/your/folder --pattern "*.pdf"

注意: このスクリプトはmacOS専用です。
"""
import subprocess
import sys
import argparse
import os
import glob

def print_file(file_path):
    """
    指定されたファイルをmacOSのデフォルトプリンタで印刷します。
    """
    if sys.platform != 'darwin':
        print("エラー: この印刷機能はmacOSでのみサポートされています。")
        return False

    if not os.path.isfile(file_path):
        print(f"エラー: パス '{file_path}' は有効なファイルではありません。")
        return False

    try:
        print(f"--- '{os.path.basename(file_path)}' を印刷します ---")
        # macOSの'lp'コマンドを使用して印刷
        subprocess.run(['lp', file_path], check=True)
        print("印刷コマンドをプリンタに送信しました。")
        print("-" * (len(os.path.basename(file_path)) + 15))
        return True
    except FileNotFoundError:
        print("エラー: 'lp'コマンドが見つかりません。macOSのシステムが正しく構成されているか確認してください。")
        return False
    except subprocess.CalledProcessError as e:
        print(f"印刷実行中にエラーが発生しました: {e}")
        return False
    except Exception as e:
        print(f"印刷中に予期せぬエラーが発生しました: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(
        description='指定されたファイルまたはディレクトリ内のファイルをmacOSのデフォルトプリンタで印刷します。',
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('path', help='処理対象のファイルまたはディレクトリのパス。')
    parser.add_argument('--pattern', default='*.docx', help='ディレクトリ処理時のファイル名パターン。例: "*.pdf" デフォルトは全ての.docxファイル ("*.docx")。')
    args = parser.parse_args()

    target_path = args.path
    files_to_process = []

    if os.path.isdir(target_path):
        search_pattern = os.path.join(target_path, args.pattern)
        files_to_process = glob.glob(search_pattern)
        print(f"ディレクトリ '{target_path}' で、パターン '{args.pattern}' に一致する {len(files_to_process)} 個のファイルを処理します。")

    elif os.path.isfile(target_path):
        files_to_process.append(target_path)

    else:
        print(f"エラー: '{target_path}' は有効なファイルまたはディレクトリではありません。")
        return

    if not files_to_process:
        print("処理対象のファイルが見つかりませんでした。")
        return
    
    # 各ファイルを処理
    for file_path in files_to_process:
        # Wordなどが作成する一時ファイル (~$で始まるファイル) を無視する
        if not os.path.basename(file_path).startswith('~$'):
            print_file(file_path)

    print("\nすべての処理が完了しました。")

if __name__ == '__main__':
    main()
