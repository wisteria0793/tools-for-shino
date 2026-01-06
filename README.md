# tools-for-shino

これは、業務に関連する日常的な作業を自動化するためのPythonスクリプト（ツール）を集めたリポジトリです。

## セットアップ

一部のスクリプトは外部ライブラリに依存しています。以下のコマンドでインストールしてください。

```bash
pip install japanera python-docx pandas openpyxl xlrd
```

## スクリプト一覧

### 1. `update_date_wareki.py` (Word用)

Word文書内の和暦の日付を、実行当日または指定した日付に自動で更新します。

#### 使い方

-   **「今日」に更新する**
    ```bash
    python scripts/update_date_wareki.py /path/to/folder
    ```
-   **「指定した日」に更新する**
    ```bash
    python scripts/update_date_wareki.py /path/to/folder --date 2025-01-30
    ```

### 2. `update_excel_date.py` (Excel用)

`documents/manage` 内のExcelファイル（.xlsx）の特定セル（請求月、送金日）を、実行当日または指定した日付に更新します。

#### 使い方

-   **「今日」に更新する**
    ```bash
    python scripts/update_excel_date.py documents/manage
    ```
-   **「指定した日」に更新する**
    ```bash
    python scripts/update_excel_date.py documents/manage --date 2026-02-01
    ```

### 3. `calculate_transfer_total.py`

`documents/manage` 内にあるExcelファイルから銀行別の送金合計金額を算出します。

#### 使い方

-   **銀行別の送金合計を表示する**
    ```bash
    python scripts/calculate_transfer_total.py
    ```

### 4. `print_word_document.py`

指定したWord文書やその他のファイルを、macOSに設定されたデフォルトのプリンタで印刷します。