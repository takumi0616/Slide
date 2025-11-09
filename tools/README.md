# PDF をそのまま画像として貼り付けた PPTX を作る手順（安全版）

本手順は、Marp の `--pptx-editable` で発生しがちな「数式の欠落」や「レイアウト崩れ」を回避するため、PDF の各ページをそのまま画像として PowerPoint スライドに貼り付けた PPTX を生成します。  
ステージ3の PDF（例: `document/stage3/SITA学会発表_stage3_20251104_175700.pdf`）が存在する前提で記載しています。

- 方式: PyMuPDF で PDF を高解像度でラスタライズ → python-pptx で各ページ画像を全面貼付け
- メリット: PDF の見た目（数式/図含む）を 100% 保持。PPTX 破損リスクが極小。
- 依存: Python 3（推奨は仮想環境）、`pymupdf`、`pillow`、`python-pptx`
- 使うスクリプト: `tools/pdf_to_pptx.py`（外部バイナリ不要）

---

## 1. 推奨: 仮想環境の作成（衝突回避）

既存の他ツールと Python ライブラリのバージョンが衝突しないよう、仮想環境で実行することを推奨します。

```bash
# プロジェクト直下で
python3 -m venv .venv
source .venv/bin/activate    # macOS / Linux
# Windows の場合: .venv\Scripts\activate

# 必要ライブラリをインストール
pip install --upgrade pip
# Pillow は 11 系の方が他ツールと衝突しにくい
pip install "pillow==11.3.0" python-pptx pymupdf
```

> 既存のグローバル環境を使う場合は `--user` で入れても動きますが、他ツール（例: streamlit / open-webui）と Pillow などのバージョンが競合することがあります。仮想環境を推奨します。

---

## 2. 変換の実行

以下は、`document/stage3` に PDF がある前提での例です。ファイル名に日本語・スペースが含まれるため、必ずクォートしてください。

```bash
# 例: 入力 PDF
# document/stage3/SITA学会発表_stage3_20251104_175700.pdf

# 変換（300dpi 推奨）
python tools/pdf_to_pptx.py \
  --input 'document/stage3/SITA学会発表_stage3_20251104_175700.pdf' \
  --output 'document/stage3/SITA学会発表_stage3_20251104_175700.pptx' \
  --dpi 300
```

- `--dpi` は 300 で十分なことが多いです。より鮮明さが欲しい場合は 400–450 に上げてください（ファイルサイズは増えます）。
- 出力先に同名ファイルがある場合は上書き保存されます。

---

## 3. スライドのアスペクト比について

`tools/pdf_to_pptx.py` は「1枚目の画像の縦横比」にスライドサイズを合わせます（余白を最小化するため）。  
16:9 固定にしたい場合の簡単な変更案:

1. `tools/pdf_to_pptx.py` の `images_to_pptx()` 内、スライドサイズ設定を次のように固定
   ```python
   prs.slide_width = Inches(13.333)  # 16:9 横幅
   prs.slide_height = Inches(7.5)    # 16:9 高さ
   ```
2. 画像貼付け時に「中央寄せ & 縦横比維持」の計算を加える（必要に応じて拡張）。現状は全面フィットです。

---

## 4. 生成物の確認

```bash
open 'document/stage3/SITA学会発表_stage3_20251104_175700.pptx'   # macOS の場合
# Windows: エクスプローラからダブルクリック
```

各スライドが 1 枚の画像として貼り付けられていれば成功です。  
PowerPoint の「修復」ダイアログは出ないはずです（python-pptx が形成した標準的な PPTX のため）。

---

## 5. トラブルシューティング

- 「モジュールが見つかりません」
  - 仮想環境が有効化されているか確認（`source .venv/bin/activate`）。
  - `pip install "pillow==11.3.0" python-pptx pymupdf` を再実行。

- 「他ツールが要求する Pillow のバージョンと衝突する」
  - 仮想環境での実行に切替（推奨）。
  - どうしてもグローバルで入れる必要がある場合は、他ツールの要件に合わせて Pillow を固定してください。

- 解像度（画質）
  - `--dpi` を 300 → 400–450 に上げると鮮明になります（ファイルサイズは増えます）。
  - PDF 側が非ベクタ（埋め込み画像主体）の場合、元画像の品質に依存します。

- 文字化け・フォント
  - 本方式は PDF を画像化するため、フォント問題は基本的に発生しません。PDF の見た目がそのまま PPTX に乗ります。

---

## 6. 参考: 代替コマンド（外部バイナリ利用）

外部バイナリ（Poppler）を使うアプローチもあります。  
インストール: `brew install poppler`

```bash
# 1) PDF → PNG (300dpi)
mkdir -p 'document/stage3/_tmp_pdf2pptx'
pdftocairo -png -r 300 \
  'document/stage3/SITA学会発表_stage3_20251104_175700.pdf' \
  'document/stage3/_tmp_pdf2pptx/slide'

# 2) 画像をPPTXへ（tools/pdf_to_pptx.py を使わない場合の例）
python - <<'PY'
from pptx import Presentation
from pptx.util import Inches
import glob
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
imgs = sorted(glob.glob('document/stage3/_tmp_pdf2pptx/slide-*.png'))
for im in imgs:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(im, 0, 0, width=prs.slide_width, height=prs.slide_height)
prs.save('document/stage3/SITA学会発表_stage3_20251104_175700.pptx')
PY
```

> なお、ImageMagick の `magick input.pdf output.pptx` は PPTX 破損を生むことがあり、推奨しません。

---

## 7. まとめ（最小コマンド）

仮想環境なしで最短実行（衝突リスクはあり）
```bash
python3 -m pip install --user pillow python-pptx pymupdf
python3 tools/pdf_to_pptx.py \
  --input 'document/stage3/SITA学会発表_stage3_20251104_175700.pdf' \
  --output 'document/stage3/SITA学会発表_stage3_20251104_175700.pptx' \
  --dpi 300
```

推奨（仮想環境＋固定バージョン）
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install "pillow==11.3.0" python-pptx pymupdf
python tools/pdf_to_pptx.py \
  --input 'document/stage3/SITA学会発表_stage3_20251104_175700.pdf' \
  --output 'document/stage3/SITA学会発表_stage3_20251104_175700.pptx' \
  --dpi 300
```

これで PDF の見た目をそのまま保った PPTX を安定して作成できます。
