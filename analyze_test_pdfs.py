"""
テスト用P&ID PDFの特性を分析するスクリプト
テキスト抽出・描画パス数・テキストアウトライン検出などを調査
"""

import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import fitz
from pathlib import Path
from collections import Counter

TEST_DIR = Path("test_pdfs")
# 元のテストファイルも含める
ORIGINAL = Path("System_test_PID_PII_R1_01.pdf")


def analyze_pdf(pdf_path):
    """PDFの各ページを分析"""
    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        print(f"  ERROR: {e}")
        return

    print(f"\n{'='*80}")
    print(f"FILE: {pdf_path.name} ({pdf_path.stat().st_size / 1024:.0f} KB)")
    print(f"  Pages: {len(doc)}")

    for page_idx in range(min(len(doc), 5)):  # 最大5ページまで分析
        page = doc[page_idx]
        print(f"\n  --- Page {page_idx+1} ---")
        print(f"  Size: {page.rect.width:.0f} x {page.rect.height:.0f} pt "
              f"({page.rect.width/72:.1f} x {page.rect.height/72:.1f} in)")
        print(f"  Rotation: {page.rotation}")

        # テキスト分析
        text_dict = page.get_text('dict')
        text_spans = []
        for block in text_dict.get('blocks', []):
            if block['type'] == 0:  # text block
                for line in block.get('lines', []):
                    for span in line.get('spans', []):
                        if span['text'].strip():
                            text_spans.append(span)

        print(f"  Text spans: {len(text_spans)}")
        if text_spans:
            sizes = [s['size'] for s in text_spans]
            fonts = Counter(s['font'] for s in text_spans)
            print(f"    Font sizes: {min(sizes):.1f} - {max(sizes):.1f} pt")
            print(f"    Fonts: {dict(fonts.most_common(5))}")
            # サンプルテキスト
            sample = text_spans[:3]
            for s in sample:
                print(f"    Sample: '{s['text'][:50]}' (size={s['size']:.1f}, font={s['font']})")

        # 描画パス分析
        drawings = page.get_drawings()
        print(f"  Drawings: {len(drawings)}")

        if not drawings:
            continue

        # アイテムタイプの統計
        item_counts = Counter()
        total_items = 0
        small_path_count = 0  # 5pt以下の小さいパス
        tiny_path_count = 0   # 2pt以下の非常に小さいパス
        closed_small_count = 0  # 閉じた小さいパス（テキストアウトラインの可能性）

        sizes = []
        for d in drawings:
            items = d['items']
            for item in items:
                item_counts[item[0]] += 1
                total_items += 1

            rect = d['rect']
            w = rect.x1 - rect.x0
            h = rect.y1 - rect.y0
            max_dim = max(w, h)
            sizes.append(max_dim)

            if max_dim < 5:
                small_path_count += 1
            if max_dim < 2:
                tiny_path_count += 1
            if max_dim < 8 and d.get('closePath', False):
                closed_small_count += 1

        print(f"  Total items: {total_items}")
        print(f"    Item types: {dict(item_counts.most_common())}")
        print(f"  Path sizes: min={min(sizes):.1f}, max={max(sizes):.1f}, "
              f"median={sorted(sizes)[len(sizes)//2]:.1f} pt")
        print(f"  Small paths (<5pt): {small_path_count} ({small_path_count/len(drawings)*100:.1f}%)")
        print(f"  Tiny paths (<2pt): {tiny_path_count} ({tiny_path_count/len(drawings)*100:.1f}%)")
        print(f"  Closed small (<8pt): {closed_small_count} ({closed_small_count/len(drawings)*100:.1f}%)")

        # テキストアウトライン検出のヒューリスティック
        # 多数の小さな閉じたベジェ曲線パス = テキストがアウトラインとして描画されている可能性
        curve_small = 0
        for d in drawings:
            items = d['items']
            rect = d['rect']
            w = rect.x1 - rect.x0
            h = rect.y1 - rect.y0
            has_curves = any(i[0] == 'c' for i in items)
            if has_curves and max(w, h) < 15 and d.get('closePath', False):
                curve_small += 1

        if curve_small > 50:
            print(f"  ⚠ TEXT OUTLINES LIKELY: {curve_small} small closed curved paths")
        elif curve_small > 10:
            print(f"  ⚡ Possible text outlines: {curve_small} small closed curved paths")

        # 予想されるシェイプ数（現在のロジックで）
        # multi_lineを展開した場合の推定
        est_shapes = 0
        for d in drawings:
            items = d['items']
            line_items = [i for i in items if i[0] == 'l']
            if any(i[0] in ('re', 'qu') for i in items):
                est_shapes += 1
            elif any(i[0] == 'c' for i in items):
                est_shapes += 1
            elif line_items:
                if len(line_items) == 1:
                    est_shapes += 1
                else:
                    est_shapes += len(line_items)  # multi_line展開

        est_shapes += len(text_spans)
        print(f"  Estimated shapes (current logic): ~{est_shapes}")

        # ページサイズ分類
        w_in = page.rect.width / 72
        h_in = page.rect.height / 72
        if abs(w_in - 11) < 1 and abs(h_in - 8.5) < 1:
            paper = "Letter landscape"
        elif abs(w_in - 8.5) < 1 and abs(h_in - 11) < 1:
            paper = "Letter portrait"
        elif abs(w_in - 17) < 1 and abs(h_in - 11) < 1:
            paper = "Tabloid landscape"
        elif abs(w_in - 11) < 1 and abs(h_in - 17) < 1:
            paper = "Tabloid portrait"
        elif abs(w_in - 11.69) < 1 and abs(h_in - 8.27) < 1:
            paper = "A4 landscape"
        elif abs(w_in - 8.27) < 1 and abs(h_in - 11.69) < 1:
            paper = "A4 portrait"
        elif abs(w_in - 16.54) < 1 and abs(h_in - 11.69) < 1:
            paper = "A3 landscape"
        elif abs(w_in - 11.69) < 1 and abs(h_in - 16.54) < 1:
            paper = "A3 portrait"
        elif abs(w_in - 33.11) < 1 and abs(h_in - 23.39) < 1:
            paper = "A1 landscape"
        elif abs(w_in - 23.39) < 1 and abs(h_in - 33.11) < 1:
            paper = "A1 portrait"
        else:
            paper = f"Custom ({w_in:.1f}x{h_in:.1f} in)"
        print(f"  Paper: {paper}")

    doc.close()


def main():
    print("P&ID PDF Test File Analysis")
    print("=" * 80)

    # 元のテストファイル
    if ORIGINAL.exists():
        analyze_pdf(ORIGINAL)

    # テストPDF
    for pdf_path in sorted(TEST_DIR.glob("*.pdf")):
        analyze_pdf(pdf_path)

    print(f"\n{'='*80}")
    print("ANALYSIS COMPLETE")


if __name__ == '__main__':
    main()
