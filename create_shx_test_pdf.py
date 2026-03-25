"""
AutoCAD SHXフォント風のテスト用PDFを生成するスクリプト

SHXフォントはシングルストロークフォント（各文字が線分で構成）。
このスクリプトはテキストを通常のフォントで描画するのではなく、
ベクターパス（線分・曲線）として描画し、AutoCAD出力を模倣する。
"""

import fitz  # PyMuPDF

# シンプルなシングルストロークフォント定義
# 各文字は [(x1,y1,x2,y2), ...] の線分リスト（0-1正規化座標）
# w: 文字幅（相対）
STROKE_FONT = {
    'A': {'w': 0.7, 'lines': [(0,1,0.35,0), (0.35,0,0.7,1), (0.1,0.6,0.6,0.6)]},
    'B': {'w': 0.6, 'lines': [(0,0,0,1), (0,0,0.4,0), (0.4,0,0.55,0.1), (0.55,0.1,0.55,0.4), (0.55,0.4,0.4,0.5), (0,0.5,0.4,0.5), (0.4,0.5,0.6,0.6), (0.6,0.6,0.6,0.9), (0.6,0.9,0.4,1), (0,1,0.4,1)]},
    'C': {'w': 0.6, 'lines': [(0.6,0.15,0.45,0), (0.45,0,0.15,0), (0.15,0,0,0.15), (0,0.15,0,0.85), (0,0.85,0.15,1), (0.15,1,0.45,1), (0.45,1,0.6,0.85)]},
    'D': {'w': 0.6, 'lines': [(0,0,0,1), (0,0,0.4,0), (0.4,0,0.6,0.2), (0.6,0.2,0.6,0.8), (0.6,0.8,0.4,1), (0.4,1,0,1)]},
    'E': {'w': 0.55, 'lines': [(0,0,0,1), (0,0,0.55,0), (0,0.5,0.4,0.5), (0,1,0.55,1)]},
    'F': {'w': 0.55, 'lines': [(0,0,0,1), (0,0,0.55,0), (0,0.5,0.4,0.5)]},
    'G': {'w': 0.65, 'lines': [(0.6,0.15,0.45,0), (0.45,0,0.15,0), (0.15,0,0,0.15), (0,0.15,0,0.85), (0,0.85,0.15,1), (0.15,1,0.45,1), (0.45,1,0.65,0.85), (0.65,0.85,0.65,0.5), (0.35,0.5,0.65,0.5)]},
    'H': {'w': 0.6, 'lines': [(0,0,0,1), (0.6,0,0.6,1), (0,0.5,0.6,0.5)]},
    'I': {'w': 0.3, 'lines': [(0.15,0,0.15,1), (0,0,0.3,0), (0,1,0.3,1)]},
    'K': {'w': 0.6, 'lines': [(0,0,0,1), (0.6,0,0,0.5), (0,0.5,0.6,1)]},
    'L': {'w': 0.55, 'lines': [(0,0,0,1), (0,1,0.55,1)]},
    'M': {'w': 0.8, 'lines': [(0,1,0,0), (0,0,0.4,0.5), (0.4,0.5,0.8,0), (0.8,0,0.8,1)]},
    'N': {'w': 0.65, 'lines': [(0,1,0,0), (0,0,0.65,1), (0.65,1,0.65,0)]},
    'O': {'w': 0.65, 'lines': [(0.15,0,0,0.15), (0,0.15,0,0.85), (0,0.85,0.15,1), (0.15,1,0.5,1), (0.5,1,0.65,0.85), (0.65,0.85,0.65,0.15), (0.65,0.15,0.5,0), (0.5,0,0.15,0)]},
    'P': {'w': 0.6, 'lines': [(0,0,0,1), (0,0,0.45,0), (0.45,0,0.6,0.1), (0.6,0.1,0.6,0.4), (0.6,0.4,0.45,0.5), (0,0.5,0.45,0.5)]},
    'R': {'w': 0.6, 'lines': [(0,0,0,1), (0,0,0.45,0), (0.45,0,0.6,0.1), (0.6,0.1,0.6,0.4), (0.6,0.4,0.45,0.5), (0,0.5,0.45,0.5), (0.3,0.5,0.6,1)]},
    'S': {'w': 0.6, 'lines': [(0.6,0.15,0.45,0), (0.45,0,0.15,0), (0.15,0,0,0.15), (0,0.15,0,0.4), (0,0.4,0.15,0.5), (0.15,0.5,0.45,0.5), (0.45,0.5,0.6,0.6), (0.6,0.6,0.6,0.85), (0.6,0.85,0.45,1), (0.45,1,0.15,1), (0.15,1,0,0.85)]},
    'T': {'w': 0.6, 'lines': [(0,0,0.6,0), (0.3,0,0.3,1)]},
    'U': {'w': 0.6, 'lines': [(0,0,0,0.85), (0,0.85,0.15,1), (0.15,1,0.45,1), (0.45,1,0.6,0.85), (0.6,0.85,0.6,0)]},
    'V': {'w': 0.7, 'lines': [(0,0,0.35,1), (0.35,1,0.7,0)]},
    'W': {'w': 0.9, 'lines': [(0,0,0.2,1), (0.2,1,0.45,0.5), (0.45,0.5,0.7,1), (0.7,1,0.9,0)]},
    'X': {'w': 0.6, 'lines': [(0,0,0.6,1), (0.6,0,0,1)]},
    'Y': {'w': 0.6, 'lines': [(0,0,0.3,0.5), (0.6,0,0.3,0.5), (0.3,0.5,0.3,1)]},
    'Z': {'w': 0.6, 'lines': [(0,0,0.6,0), (0.6,0,0,1), (0,1,0.6,1)]},
    '0': {'w': 0.6, 'lines': [(0.15,0,0,0.15), (0,0.15,0,0.85), (0,0.85,0.15,1), (0.15,1,0.45,1), (0.45,1,0.6,0.85), (0.6,0.85,0.6,0.15), (0.6,0.15,0.45,0), (0.45,0,0.15,0)]},
    '1': {'w': 0.35, 'lines': [(0.05,0.2,0.2,0), (0.2,0,0.2,1), (0,1,0.35,1)]},
    '2': {'w': 0.6, 'lines': [(0,0.15,0.15,0), (0.15,0,0.45,0), (0.45,0,0.6,0.15), (0.6,0.15,0.6,0.35), (0.6,0.35,0,1), (0,1,0.6,1)]},
    '3': {'w': 0.6, 'lines': [(0,0.15,0.15,0), (0.15,0,0.45,0), (0.45,0,0.6,0.15), (0.6,0.15,0.6,0.4), (0.6,0.4,0.45,0.5), (0.2,0.5,0.45,0.5), (0.45,0.5,0.6,0.6), (0.6,0.6,0.6,0.85), (0.6,0.85,0.45,1), (0.45,1,0.15,1), (0.15,1,0,0.85)]},
    '4': {'w': 0.6, 'lines': [(0.45,1,0.45,0), (0.45,0,0,0.65), (0,0.65,0.6,0.65)]},
    '5': {'w': 0.6, 'lines': [(0.6,0,0,0), (0,0,0,0.45), (0,0.45,0.45,0.45), (0.45,0.45,0.6,0.6), (0.6,0.6,0.6,0.85), (0.6,0.85,0.45,1), (0.45,1,0.15,1), (0.15,1,0,0.85)]},
    '6': {'w': 0.6, 'lines': [(0.5,0,0.15,0), (0.15,0,0,0.15), (0,0.15,0,0.85), (0,0.85,0.15,1), (0.15,1,0.45,1), (0.45,1,0.6,0.85), (0.6,0.85,0.6,0.6), (0.6,0.6,0.45,0.5), (0.45,0.5,0,0.5)]},
    '7': {'w': 0.6, 'lines': [(0,0,0.6,0), (0.6,0,0.2,1)]},
    '8': {'w': 0.6, 'lines': [(0.15,0,0,0.15), (0,0.15,0,0.4), (0,0.4,0.15,0.5), (0.15,0.5,0.45,0.5), (0.45,0.5,0.6,0.6), (0.6,0.6,0.6,0.85), (0.6,0.85,0.45,1), (0.45,1,0.15,1), (0.15,1,0,0.85), (0,0.85,0,0.6), (0,0.6,0.15,0.5), (0.45,0,0.6,0.15), (0.6,0.15,0.6,0.4), (0.6,0.4,0.45,0.5), (0.15,0,0.45,0)]},
    '9': {'w': 0.6, 'lines': [(0.6,0.5,0.15,0.5), (0.15,0.5,0,0.35), (0,0.35,0,0.15), (0,0.15,0.15,0), (0.15,0,0.45,0), (0.45,0,0.6,0.15), (0.6,0.15,0.6,0.85), (0.6,0.85,0.45,1), (0.45,1,0.1,1)]},
    '-': {'w': 0.4, 'lines': [(0,0.5,0.4,0.5)]},
    '/': {'w': 0.4, 'lines': [(0,1,0.4,0)]},
    '"': {'w': 0.3, 'lines': [(0.05,0,0.05,0.25), (0.2,0,0.2,0.25)]},
    '.': {'w': 0.15, 'lines': [(0.05,0.9,0.1,1), (0.1,1,0.05,1), (0.05,1,0.05,0.9)]},
    ' ': {'w': 0.3, 'lines': []},
}


def draw_shx_text(page, text, x, y, font_size, color=(0, 0, 0), line_width=0.4):
    """SHXフォント風にテキストをベクターパスとして描画"""
    char_height = font_size
    spacing = font_size * 0.15  # 文字間スペース
    cursor_x = x

    shape = page.new_shape()
    for ch in text.upper():
        glyph = STROKE_FONT.get(ch)
        if glyph is None:
            cursor_x += font_size * 0.3 + spacing
            continue

        char_width = glyph['w'] * char_height
        for x1, y1, x2, y2 in glyph['lines']:
            px1 = cursor_x + x1 * char_height
            py1 = y + y1 * char_height
            px2 = cursor_x + x2 * char_height
            py2 = y + y2 * char_height
            shape.draw_line(fitz.Point(px1, py1), fitz.Point(px2, py2))

        cursor_x += char_width + spacing

    shape.finish(color=color, width=line_width)
    shape.commit()
    return cursor_x - x  # 描画幅を返す


def create_shx_test_pdf(output_path="test_pdfs/shx_test_pid.pdf"):
    """SHXフォント風のP&IDテスト用PDFを生成"""
    doc = fitz.open()

    # ページ1: P&ID風の図面（Tabloid横向き）
    page = doc.new_page(width=17*72, height=11*72)  # 17x11 inches
    shape = page.new_shape()

    # === タイトルブロック ===
    # 外枠
    shape.draw_rect(fitz.Rect(20, 20, 17*72-20, 11*72-20))
    shape.finish(color=(0, 0, 0), width=1.5)
    shape.commit()

    # タイトルブロック右下
    tx, ty = 900, 680
    shape = page.new_shape()
    shape.draw_rect(fitz.Rect(tx, ty, tx+300, ty+90))
    shape.draw_line(fitz.Point(tx, ty+30), fitz.Point(tx+300, ty+30))
    shape.draw_line(fitz.Point(tx, ty+60), fitz.Point(tx+300, ty+60))
    shape.draw_line(fitz.Point(tx+150, ty), fitz.Point(tx+150, ty+90))
    shape.finish(color=(0, 0, 0), width=0.5)
    shape.commit()

    draw_shx_text(page, "PROJECT TITLE", tx+10, ty+5, 8)
    draw_shx_text(page, "TEST P AND ID", tx+155, ty+5, 8)
    draw_shx_text(page, "DRAWING NO", tx+10, ty+35, 7)
    draw_shx_text(page, "PID-001-A", tx+155, ty+35, 8)
    draw_shx_text(page, "REV", tx+10, ty+65, 7)
    draw_shx_text(page, "R1", tx+155, ty+65, 8)

    # === プロセス機器 ===
    # タンク T-101
    shape = page.new_shape()
    shape.draw_rect(fitz.Rect(100, 200, 220, 450))
    shape.finish(color=(0, 0, 0), width=1.0)
    shape.commit()
    draw_shx_text(page, "T-101", 120, 310, 12)
    draw_shx_text(page, "FEED TANK", 105, 340, 8)
    draw_shx_text(page, "VOLUME", 115, 360, 7)
    draw_shx_text(page, "5000 GAL", 110, 380, 7)

    # ポンプ P-101（円で表現）
    shape = page.new_shape()
    shape.draw_circle(fitz.Point(370, 420), 25)
    shape.finish(color=(0, 0, 0), width=1.0)
    shape.commit()
    draw_shx_text(page, "P-101", 345, 455, 8)

    # 熱交換器 E-101（縦長矩形 + 内部線）
    shape = page.new_shape()
    shape.draw_rect(fitz.Rect(500, 250, 560, 450))
    shape.draw_line(fitz.Point(500, 300), fitz.Point(560, 300))
    shape.draw_line(fitz.Point(500, 400), fitz.Point(560, 400))
    shape.finish(color=(0, 0, 0), width=1.0)
    shape.commit()
    draw_shx_text(page, "E-101", 505, 340, 8)
    draw_shx_text(page, "HEAT", 510, 355, 7)
    draw_shx_text(page, "EXCHANGER", 503, 370, 6)

    # リアクター R-101（円）
    shape = page.new_shape()
    shape.draw_circle(fitz.Point(720, 350), 60)
    shape.finish(color=(0, 0, 0), width=1.5)
    shape.commit()
    draw_shx_text(page, "R-101", 690, 340, 10)

    # === 配管ライン ===
    shape = page.new_shape()
    # T-101 → P-101
    shape.draw_line(fitz.Point(220, 420), fitz.Point(345, 420))
    # P-101 → E-101
    shape.draw_line(fitz.Point(395, 420), fitz.Point(500, 420))
    shape.draw_line(fitz.Point(500, 420), fitz.Point(500, 400))
    # E-101 → R-101
    shape.draw_line(fitz.Point(560, 350), fitz.Point(660, 350))
    # R-101出口
    shape.draw_line(fitz.Point(780, 350), fitz.Point(880, 350))
    shape.finish(color=(0, 0, 0), width=1.0)
    shape.commit()

    # ライン番号
    draw_shx_text(page, "6\"-PL-101", 240, 405, 7, color=(0, 0, 0.8))
    draw_shx_text(page, "4\"-PL-102", 400, 405, 7, color=(0, 0, 0.8))
    draw_shx_text(page, "8\"-PL-103", 590, 335, 7, color=(0, 0, 0.8))
    draw_shx_text(page, "8\"-PL-104", 800, 335, 7, color=(0, 0, 0.8))

    # === バルブ記号 ===
    # ゲートバルブ（ボウタイ）
    bx, by = 290, 410
    shape = page.new_shape()
    shape.draw_line(fitz.Point(bx, by), fitz.Point(bx+10, by+10))
    shape.draw_line(fitz.Point(bx+10, by+10), fitz.Point(bx, by+20))
    shape.draw_line(fitz.Point(bx+10, by), fitz.Point(bx, by+10))
    shape.draw_line(fitz.Point(bx, by+10), fitz.Point(bx+10, by+20))
    shape.finish(color=(0, 0, 0), width=0.8)
    shape.commit()
    draw_shx_text(page, "HV-101", 272, 435, 6)

    # === 計器 ===
    # 温度計 TI-101
    shape = page.new_shape()
    ix, iy = 620, 280
    shape.draw_circle(fitz.Point(ix, iy), 12)
    shape.draw_line(fitz.Point(ix, iy+12), fitz.Point(ix, iy+30))
    shape.finish(color=(0, 0, 0), width=0.5)
    shape.commit()
    draw_shx_text(page, "TI", ix-6, iy-6, 7)
    draw_shx_text(page, "101", ix-9, iy+2, 6)

    # 圧力計 PI-101
    shape = page.new_shape()
    ix, iy = 620, 480
    shape.draw_circle(fitz.Point(ix, iy), 12)
    shape.draw_line(fitz.Point(ix, iy-12), fitz.Point(ix, iy-30))
    shape.draw_line(fitz.Point(ix, iy-30), fitz.Point(560, iy-30))
    shape.finish(color=(0, 0, 0), width=0.5)
    shape.commit()
    draw_shx_text(page, "PI", ix-6, iy-6, 7)
    draw_shx_text(page, "101", ix-9, iy+2, 6)

    # フロー計 FI-101
    shape = page.new_shape()
    ix, iy = 820, 280
    shape.draw_circle(fitz.Point(ix, iy), 12)
    shape.draw_line(fitz.Point(ix, iy+12), fitz.Point(ix, iy+30))
    shape.draw_line(fitz.Point(ix, iy+30), fitz.Point(820, 350))
    shape.finish(color=(0, 0, 0), width=0.5)
    shape.commit()
    draw_shx_text(page, "FI", ix-6, iy-6, 7)
    draw_shx_text(page, "101", ix-9, iy+2, 6)

    # === 凡例エリア ===
    lx, ly = 50, 530
    draw_shx_text(page, "LEGEND", lx, ly, 10)
    draw_shx_text(page, "HV - HAND VALVE", lx+10, ly+25, 7)
    draw_shx_text(page, "TI - TEMPERATURE INDICATOR", lx+10, ly+45, 7)
    draw_shx_text(page, "PI - PRESSURE INDICATOR", lx+10, ly+65, 7)
    draw_shx_text(page, "FI - FLOW INDICATOR", lx+10, ly+85, 7)
    draw_shx_text(page, "PL - PROCESS LINE", lx+10, ly+105, 7)

    # ノート
    draw_shx_text(page, "NOTES", 500, ly, 10)
    draw_shx_text(page, "1. ALL DIMENSIONS IN INCHES", 510, ly+25, 7)
    draw_shx_text(page, "2. DESIGN PRESSURE 150 PSI", 510, ly+45, 7)
    draw_shx_text(page, "3. DESIGN TEMPERATURE 300 F", 510, ly+65, 7)

    doc.save(output_path)
    doc.close()
    print(f"SHXテスト用PDF生成完了: {output_path}")

    # 生成したPDFの分析
    doc = fitz.open(output_path)
    page = doc[0]
    text_spans = len(page.get_text('words'))
    drawings = len(page.get_drawings())
    print(f"  テキストスパン: {text_spans} (SHX→0であるべき)")
    print(f"  描画パス: {drawings}")
    doc.close()


if __name__ == '__main__':
    create_shx_test_pdf()
