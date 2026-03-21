# P&ID PDF → Excel 変換プロジェクト 進捗記録

## 目的
PDFのP&ID（配管計装図）を読み取り、Excelのテキスト・図形へ変換してxlsxファイルへ出力するPythonスクリプト

## サンプルPDF
- `System_test_PID_PII_R1_01.pdf` - LZ R&D P&ID (1224x792pt, 17x11インチ横向き)
- 参照画像: `System_test_PID_PII_R1_01_pdf.png` (2550x1650px, 150dpi)
- **ページ回転**: 270°（mediabox: 792x1224、表示: 1224x792）

## PDFの内容分析
- **ページ数**: 1
- **テキストスパン**: 243個（フォント: ArialMT, サイズ: 3.1〜13.4pt）
  - dir=(0,1): 192テキスト → 表示空間で水平
  - dir=(-1,0): 32テキスト → 表示空間で90°回転（垂直）
  - dir=(1,0): 15テキスト → 表示空間で270°回転（垂直）
- **描画パス**: 1051個（全て変換対象）
  - 直線: 868
  - ベジェ曲線: 81 (円・弧・記号)
  - 矩形/quad: 102
- **画像**: 0（全てベクター描画）
- **色**: 黒だけでなく赤・青・シアン・オレンジ・緑など多色
- **線の太さ**: 0.24pt〜2.28pt（6段階）
- **破線パターン**: なし（全て実線）

## 現在のスクリプト
- `pid2xlsx.py` - メイン変換スクリプト（1863図形を生成）
- `verify.py` - PDF vs Excel 比較検証スクリプト（Excel COM対応）

## 変換品質
- **ピクセル差分**: 平均19.5/255、差分>10: 13.9%
- **視覚的一致度**: 容器形状、配管線、バルブ記号、テキスト、表、凡例が正しく変換
- **残る差異**: フォントレンダリング、アンチエイリアシング（構造的限界）

## 完了済み

### Phase 1: 検証環境の確立 ✅
- Excel COM (win32com) を使用してxlsx→PDF→PNG変換
- verify.py でPDF/Excel並列比較画像を自動生成
- LibreOfficeは不要（Excel COM方式で代替）

### Phase 2: 座標系の修正 ✅
- **EMU値実測**: Excel COMで列幅2.14chars = 12.75pt = 161925 EMU を確認
- **ページ回転対応**: mediabox座標→表示座標の変換（0°/90°/180°/270°対応）
  - 270°変換式: display_x = mediabox_y, display_y = mediabox_width - mediabox_x
- **ページ設定**: 横向き(landscape)、Tabloidサイズ、1ページに収める、余白最小

### Phase 3: 図形変換の改善 ✅
- **カスタムジオメトリ(freeform)サポート**: ベジェ曲線→custGeomで正確なパス描画
- **複数直線パス**: 連続した直線をフリーフォームとして正確に変換
- **楕円**: 4+ベジェ曲線→アスペクト比制限を撤廃
- **単一直線**: 実際の始点・終点座標を使用（bounding box → actual points）
- **flipH/flipV対応**: 右上→左下方向の直線を正しく描画
- **バルブ記号**: 3直線X型パターン（ボウタイ）→ カスタムジオメトリで描画
  - 検出条件: 3直線、4端点、サイズ5-30pt、アスペクト比<2.5、対角線2本以上
  - 14個のバルブを正確に検出（偽陽性の3辺矩形を排除）
  - バルブの辺にあたる単一直線パスを自動抑制

### Phase 4: テキストの改善 ✅
- **回転テキスト**: ページ回転を考慮した表示空間角度を計算
  - 90°回転: `vert="vert"`（上→下）
  - 270°回転: `vert="vert270"`（下→上）
- **テキスト色**: PDFからspan.colorを取得してExcelに反映
- **フォントサイズ**: PDFのspan.sizeをそのまま使用

### Phase 5: 線スタイルの改善 ✅
- **線の太さ**: PDF線幅をEMU変換、最小0.25pt（3175 EMU）
- **破線パターン**: PDFに破線なし（全て実線）→対応不要
- **グリッド線非表示**: showGridLines = False

### Phase 6: 品質改善と安全対策 ✅
- **バルブ誤検出修正**: 対角線チェック追加で3辺矩形の偽陽性を排除（27→14個）
- **水平/垂直線の精度**: 線のアンカー最小サイズ確保を線以外に限定（斜め化を防止）
- **斜め線flipH/flipV修正**: `spPr`ではなく`a:xfrm`にflipH/flipV属性を設定（OOXML仕様準拠）
  - 修正前: Excelがflip属性を無視し全斜め線がデフォルト方向で描画
  - 修正後: flipH/flipVがa:xfrm要素に正しく設定されExcelが認識
- **直線矩形検出**: 3+本のH/V直線で構成されたclosePathの矩形を正しくrectと分類
- **ゼロ長線（ドット）**: 始点=終点の線を塗りつぶし円（枠なし）として変換（201個）
- **バルブ検出ロジック共通化**: `_is_valve_pattern()`関数で重複コード削除
- **一時ファイル安全対策**: inject_drawingsにtry/finally追加、失敗時に.tmpを削除
- **三角形プリセット**: 3直線/3頂点の矢印マーカー（82個）をExcel `triangle`プリセットで描画
  - 頂点方向の自動判定（上/下/左/右）→ 適切な回転値を設定
  - 色付き塗りつぶし対応
- **ホームベースプリセット**: 5直線/5頂点の五角形（10個）をExcel `homePlate`プリセットで描画
  - 尖端方向の自動判定 → 適切な回転値を設定

## 今後の改善候補（優先度低）
1. **テキスト位置の微調整** - セルグリッド丸めによる1-2ptの誤差
2. **oneCellAnchor方式** - twoCellAnchor依存をなくしてEMU精度向上
3. **矢印シンボル** - 三角形の矢印の検出と変換
4. **複数ページ対応** の検証
5. **異なるPDFでの汎用性テスト**

## 依存パッケージ
- PyMuPDF (fitz) 1.26.4 - PDF解析
- openpyxl 3.1.5 - Excel作成
- pywin32 - Excel COM自動化（verify.py用）
- Pillow - 画像比較

## 技術メモ
- openpyxlの`SpreadsheetDrawing`は`_write`メソッドで`charts`と`images`のみ処理 → Drawing XMLをzipに直接注入
- Excel Drawing XMLの名前空間: `http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing`
- 図形名前空間: `http://schemas.openxmlformats.org/drawingml/2006/main`
- ページ回転270°: get_text()/get_drawings()はmediabox座標を返す → make_coord_transform()で変換
- Excel列幅: openpyxl width=2.14 → Excel実測 12.75pt = 161925 EMU
- カスタムジオメトリ: GEOM_SCALE=100000でパス座標を正規化
- テキスト方向: PyMuPDF dir=(dx,dy) → ページ回転後の表示角度を計算 → Excel bodyPr.vert で設定
