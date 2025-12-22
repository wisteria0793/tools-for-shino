# tools-for-shino

これは、Word文書に関連する日常的な作業を自動化するためのPythonスクリプト（ツール）を集めたリポジトリです。

## セットアップ

一部のスクリプトは外部ライブラリに依存しています。以下のコマンドでインストールしてください。

```bash
pip install japanera python-docx
```

## スクリプト一覧

### 1. `update_date_wareki.py`

Word文書内の和暦の日付を、スクリプトを実行した当日の日付に自動で更新します。

#### 使い方

-   **単一のファイルの日付を更新する**
    ```bash
    python scripts/update_date_wareki.py /path/to/your/file.docx
    ```

-   **日付を更新して、そのまま印刷する (macOSのみ)**
    ```bash
    python scripts/update_date_wareki.py /path/to/your/file.docx --print
    ```

-   **フォルダ内のすべての`.docx`ファイルの日付を更新する**
    ```bash
    python scripts/update_date_wareki.py /path/to/your/folder
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
