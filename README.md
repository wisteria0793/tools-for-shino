# tools-for-shino

これは、業務に関連する日常的な作業を自動化するためのPythonスクリプト（ツール）を集めたリポジトリです。

## セットアップ

一部のスクリプトは外部ライブラリに依存しています。以下のコマンドでインストールしてください。

```bash
pip install japanera python-docx pandas openpyxl xlrd
```

## スクリプト一覧

### 1. `update_date_wareki.py`

Word文書内の和暦の日付を、スクリプトを実行した当日の日付に自動で更新します。

#### 使い方

-   **フォルダ内のすべての`.docxファイル`に記載されている日付を更新する**

    ```bash
    python scripts/update_date_wareki.py /path/to/folder --pattern "*.docx"
    ```


### 2. `print_word_document.py`

指定したWord文書やその他のファイル（PDFなど）を、macOSに設定されたデフォルトのプリンタで印刷します。

#### 使い方

-   **単一のファイルを印刷する**
    ```bash
    python scripts/print_word_document.py /path/to/your/file.docx
    ```

-   **フォルダ内のすべての`.docx`ファイルを印刷する**
    ```bash
    python scripts/print_word_document.py /path/to/your/folder
    ```

-   **フォルダ内のPDFファイルをすべて印刷する**
    ```bash
    python scripts/print_word_document.py /path/to/your/folder --pattern "*.pdf"
    ```

### 3. `calculate_transfer_total.py`

`documents/manage` 内にあるExcelファイルから銀行別の送金合計金額を算出します。ファイル名の先頭（「_」まで）を銀行名として集計します。

#### 使い方

-   **銀行別の送金合計を表示する**
    ```bash
    python scripts/calculate_transfer_total.py
    ```
    ※ 内部で `documents/manage` フォルダを参照します。