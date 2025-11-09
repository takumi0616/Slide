# Cline によるスライド自動生成システム 利用ガイド

---

## cline に投げるプロンプト例

```
スライドの元となる資料パス（以降は「元資料」と呼ぶ）
/Users/takumi0616/Slide/infomation/main_v8.py
/Users/takumi0616/Slide/infomation/minisom_v8.py
/Users/takumi0616/Slide/infomation/search_results_v8.txt
/Users/takumi0616/Slide/infomation/SITA2025_presentation_タカスカ.md
/Users/takumi0616/Slide/infomation/SITA2025_高須賀 (16).pdf

使用するCSSファイル
/Users/takumi0616/Slide/marp-theme-academic/themes/academic-tml.css

このREADMEを参考に、上記の元資料からスライドを作成してください。
基本情報や構成はSITA2025_presentation_タカスカ.mdを参考にしてください（これは量が少ないだけの完成版に近いです）
スライドは30枚にしてください（タイトル込み）
日本語で作成してください
タイトル：SITA 学会発表
著者：Takasuka Takumi（長岡技術科学大学）
おこなったことを日本語で解説
```

---

## システム概要

このシステムは、Cline の MCP ツール（markdownify-mcp 等）と Marp を活用し、  
PDF や Word などの資料ファイルから美しいスライド資料（PDF および 編集可能な PPTX）を 3 段階のステージで自動生成する仕組みです。

---

## スライド作成の 3 ステージプロセス

### Stage 1：情報整理とマークダウン作成

- 元資料から情報を抽出し、整理・加筆・修正
- 綺麗なマークダウン形式で保存（スライド用ではない通常のマークダウン）
- 保存先：`/Users/takumi0616/Slide/document/stage1/`
- ファイル命名規則：`title_stage1_YYYYMMDD_HHMMSS.md`

### Stage 2：スライド用マークダウン作成

- Stage 1 のファイルを読み取り、Marp 用のスライドマークダウンに変換
- スライド区切り、見出し、frontmatter などを自動付与
- 保存先：`/Users/takumi0616/Slide/document/stage2/`
- ファイル命名規則：`title_stage2_YYYYMMDD_HHMMSS.md`

### Stage 3：スライドファイル生成

- Stage 2 のマークダウンと CSS テーマを使用してスライドを生成
- PDF 形式および編集可能な PPTX 形式を作成
- 保存先：`/Users/takumi0616/Slide/document/stage3/`
- ファイル命名規則：
  - `title_stage3_YYYYMMDD_HHMMSS.pdf`
  - `title_stage3_editable_YYYYMMDD_HHMMSS.pptx`

---

## ディレクトリ構成

```
/Users/takumi0616/Slide/
├── document/
│   ├── README.md
│   ├── stage1/ (情報整理済みマークダウン)
│   ├── stage2/ (スライド用マークダウン)
│   └── stage3/ (生成されたPDF・PPTXファイル)
└── marp-theme-academic/
    ├── themes/ (CSSテーマファイル)
    └── images/ (背景画像等)
```

---

## テーマ（デザイン）の種類と切り替え

- 利用可能なカスタムテーマ:
  - `academic-tcue` (緑色基調)
  - `academic-tml` (青色基調)
  - `academic-wsd` (赤色基調)
- テーマは Markdown 先頭の frontmatter で指定できます。
  ```
  ---
  marp: true
  theme: academic-tcue
  ---
  ```
- サンプル比較用 PDF:
  - `marp_theme_sample_tcue.pdf`
  - `marp_theme_sample_tml.pdf`
  - `marp_theme_sample_wsd.pdf`

---

## 各ステージでの処理内容

### Stage 1 の処理詳細

1. MCP ツールで資料を Markdown に変換
2. 内容を整理、加筆、修正
3. 読みやすく構造化されたマークダウンに
4. タイトルを最初に統一（以降のファイル名に使用）

### Stage 2 の処理詳細

1. Marp の frontmatter を追加
2. スライド区切り（`---`）を適切に配置
3. 見出しレベルを調整（h1, h2, h3 など）
4. 画像サイズや表の最適化
5. スライド特有の要素（ヘッダー、フッターなど）を設定

### Stage 3 の処理詳細

1. Marp CLI で PDF スライドを生成
2. 編集可能な PPTX ファイルを生成
   - LibreOffice を使用して生成され、PowerPoint で各要素を直接編集できます
3. 両方のファイルを stage3 ディレクトリに保存
4. ファイル命名規則：
   - PDF: `title_stage3_YYYYMMDD_HHMMSS.pdf`
   - 編集可能 PPTX: `title_stage3_editable_YYYYMMDD_HHMMSS.pptx`

---

## 依頼時のテンプレート例

```
[基本依頼例]
/Users/takumi0616/path/to/your/document.pdf からスライドを作成してください。
タイトル：研究発表資料

[テーマ指定例]
/Users/takumi0616/path/to/your/document.docx を academic-wsd テーマでスライド化してください。
タイトル：プロジェクト進捗報告

[スライド生成例]
/Users/takumi0616/path/to/your/document.pdf からスライドを作成してください。
タイトル：研究発表資料
```

---

## 注意点・FAQ

- **すべてのステージは自動化されます**  
  Cline が 3 つのステージをすべて自動で処理します。特別な指示がない限り、すべてのステージの出力が生成されます。

- **各ステージの結果の確認**  
  各ステージの結果を確認したい場合は、対応するディレクトリ内のファイルを開いてください。

- **特定のステージからの再開**  
  特定のステージから再開したい場合は、前のステージの出力ファイルパスを指定してください。例：

  ```
  Stage 2から再開：/Users/takumi0616/Slide/document/stage1/mytitle_stage1_20250508_123000.md からスライドを作成してください。
  ```

- **テーマのカスタマイズ**  
  `marp-theme-academic/themes/` ディレクトリの CSS ファイルを編集することで、独自テーマを作成できます。

- **エラー時の対応**  
  生成に失敗した場合は、エラーメッセージを確認し、Cline に伝えてください。

- **PPTX ファイルについて**  
  本システムでは編集可能な PPTX ファイルのみを生成します。

  1. **前提条件**：

     - LibreOffice がインストールされていること（`brew install --cask libreoffice`でインストール可能）

  2. **手動での生成（もし必要な場合）**:

     ```
     # PDF生成
     npx @marp-team/marp-cli document/stage2/タイトル_stage2_YYYYMMDD_HHMMSS.md --output document/stage3/タイトル_stage3_YYYYMMDD_HHMMSS.pdf

     # 編集可能なPPTX生成
     npx @marp-team/marp-cli document/stage2/タイトル_stage2_YYYYMMDD_HHMMSS.md --output document/stage3/タイトル_stage3_editable_YYYYMMDD_HHMMSS.pptx --pptx --pptx-editable --allow-local-files
     ```

  **注意事項**：

  - 編集可能な PPTX は LibreOffice を経由して生成するため、一部のデザイン要素が完全には保持されない場合があります
  - `--pptx-editable`オプションは実験的機能であり、スライドの完全な再現性は保証されていません
  - ローカルの画像ファイルを使用する場合は`--allow-local-files`オプションが必要です

---

## まとめ

- 3 段階のステージで、元資料から高品質なスライドを自動生成します。
- 統一されたタイトルとファイル命名規則で、管理が容易になります。
- 各ステージの出力を確認・編集することで、細かな調整も可能です。
- PDF と編集可能な PPTX 形式で出力されるため、様々な環境で利用できます。
- 編集可能な PPTX は LibreOffice を使用して生成され、特別なコマンドオプション（`--pptx-editable --allow-local-files`）を自動的に適用します。
