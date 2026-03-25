"""
日本語テキストを含むSHXフォント風テスト用PDFを生成

- 機器名称・ライン番号: SHXストローク（英数字）
- 注記・凡例: 通常フォント日本語テキスト
- P&ID図形: 通常の線・円・矩形

easyocrが日本語を認識できるかのテスト用。
"""

import fitz
from create_shx_test_pdf import draw_shx_text, STROKE_FONT


def create_ja_test_pdf(output_path="test_pdfs/shx_test_pid_ja.pdf"):
    doc = fitz.open()
    # A3横向き
    page = doc.new_page(width=16.54*72, height=11.69*72)
    shape = page.new_shape()

    # 外枠
    shape.draw_rect(fitz.Rect(20, 20, 16.54*72-20, 11.69*72-20))
    shape.finish(color=(0, 0, 0), width=1.5)
    shape.commit()

    # === 機器（図形はベクター描画）===

    # タンク T-101
    shape = page.new_shape()
    shape.draw_rect(fitz.Rect(100, 150, 240, 420))
    shape.finish(color=(0, 0, 0), width=1.2)
    shape.commit()
    # タンク名（SHXストローク）
    draw_shx_text(page, "T-101", 130, 260, 12)
    # 日本語名（通常フォント）
    page.insert_text(fitz.Point(115, 300), "原料タンク",
                     fontname="japan", fontsize=10, color=(0, 0, 0))
    page.insert_text(fitz.Point(110, 320), "容量: 5000 L",
                     fontname="japan", fontsize=8, color=(0, 0, 0))

    # ポンプ P-101
    shape = page.new_shape()
    shape.draw_circle(fitz.Point(370, 390), 28)
    shape.finish(color=(0, 0, 0), width=1.0)
    shape.commit()
    draw_shx_text(page, "P-101", 343, 430, 8)
    page.insert_text(fitz.Point(340, 450), "送液ポンプ",
                     fontname="japan", fontsize=8, color=(0, 0, 0))

    # 熱交換器 E-101
    shape = page.new_shape()
    shape.draw_rect(fitz.Rect(520, 200, 590, 430))
    shape.draw_line(fitz.Point(520, 260), fitz.Point(590, 260))
    shape.draw_line(fitz.Point(520, 370), fitz.Point(590, 370))
    shape.finish(color=(0, 0, 0), width=1.0)
    shape.commit()
    draw_shx_text(page, "E-101", 528, 310, 8)
    page.insert_text(fitz.Point(525, 340), "熱交換器",
                     fontname="japan", fontsize=8, color=(0, 0, 0))

    # 反応器 R-101
    shape = page.new_shape()
    shape.draw_circle(fitz.Point(760, 320), 65)
    shape.finish(color=(0, 0, 0), width=1.5)
    shape.commit()
    draw_shx_text(page, "R-101", 725, 310, 11)
    page.insert_text(fitz.Point(725, 340), "反応器",
                     fontname="japan", fontsize=10, color=(0, 0, 0))

    # 蒸留塔 C-101
    shape = page.new_shape()
    shape.draw_rect(fitz.Rect(920, 120, 980, 450))
    # 段トレイ
    for ty in range(160, 440, 30):
        shape.draw_line(fitz.Point(920, ty), fitz.Point(980, ty))
    shape.finish(color=(0, 0, 0), width=0.8)
    shape.commit()
    draw_shx_text(page, "C-101", 925, 470, 8)
    page.insert_text(fitz.Point(920, 490), "蒸留塔",
                     fontname="japan", fontsize=8, color=(0, 0, 0))

    # === 配管ライン ===
    shape = page.new_shape()
    shape.draw_line(fitz.Point(240, 390), fitz.Point(342, 390))
    shape.draw_line(fitz.Point(398, 390), fitz.Point(520, 390))
    shape.draw_line(fitz.Point(520, 390), fitz.Point(520, 370))
    shape.draw_line(fitz.Point(590, 315), fitz.Point(695, 315))
    shape.draw_line(fitz.Point(825, 315), fitz.Point(920, 315))
    shape.finish(color=(0, 0, 0), width=1.0)
    shape.commit()

    # ライン番号（SHX）
    draw_shx_text(page, "6\"-PL-101", 250, 375, 7, color=(0, 0, 0.8))
    draw_shx_text(page, "4\"-PL-102", 420, 375, 7, color=(0, 0, 0.8))
    draw_shx_text(page, "8\"-PL-103", 610, 300, 7, color=(0, 0, 0.8))

    # === バルブ ===
    bx, by = 300, 380
    shape = page.new_shape()
    shape.draw_line(fitz.Point(bx, by), fitz.Point(bx+12, by+12))
    shape.draw_line(fitz.Point(bx+12, by+12), fitz.Point(bx, by+24))
    shape.draw_line(fitz.Point(bx+12, by), fitz.Point(bx, by+12))
    shape.draw_line(fitz.Point(bx, by+12), fitz.Point(bx+12, by+24))
    shape.finish(color=(0, 0, 0), width=0.8)
    shape.commit()
    draw_shx_text(page, "HV-101", 280, 410, 6)

    # === 計器 ===
    # 温度指示計
    shape = page.new_shape()
    ix, iy = 650, 230
    shape.draw_circle(fitz.Point(ix, iy), 14)
    shape.draw_line(fitz.Point(ix, iy+14), fitz.Point(ix, iy+35))
    shape.finish(color=(0, 0, 0), width=0.5)
    shape.commit()
    draw_shx_text(page, "TI", ix-6, iy-6, 7)
    draw_shx_text(page, "101", ix-10, iy+3, 6)

    # 圧力指示計
    shape = page.new_shape()
    ix, iy = 650, 460
    shape.draw_circle(fitz.Point(ix, iy), 14)
    shape.draw_line(fitz.Point(ix, iy-14), fitz.Point(ix, iy-35))
    shape.finish(color=(0, 0, 0), width=0.5)
    shape.commit()
    draw_shx_text(page, "PI", ix-6, iy-6, 7)
    draw_shx_text(page, "102", ix-10, iy+3, 6)

    # === 日本語テキストエリア（凡例・注記）===
    lx, ly = 50, 520

    page.insert_text(fitz.Point(lx, ly), "凡例",
                     fontname="japan", fontsize=12, color=(0, 0, 0))
    page.insert_text(fitz.Point(lx+10, ly+25), "HV - 手動弁",
                     fontname="japan", fontsize=8, color=(0, 0, 0))
    page.insert_text(fitz.Point(lx+10, ly+45), "TI - 温度指示計",
                     fontname="japan", fontsize=8, color=(0, 0, 0))
    page.insert_text(fitz.Point(lx+10, ly+65), "PI - 圧力指示計",
                     fontname="japan", fontsize=8, color=(0, 0, 0))
    page.insert_text(fitz.Point(lx+10, ly+85), "PL - プロセスライン",
                     fontname="japan", fontsize=8, color=(0, 0, 0))

    # 注記
    nx, ny = 500, 520
    page.insert_text(fitz.Point(nx, ny), "注記",
                     fontname="japan", fontsize=12, color=(0, 0, 0))
    page.insert_text(fitz.Point(nx+10, ny+25), "1. 設計圧力: 1.0 MPaG",
                     fontname="japan", fontsize=8, color=(0, 0, 0))
    page.insert_text(fitz.Point(nx+10, ny+45), "2. 設計温度: 150°C",
                     fontname="japan", fontsize=8, color=(0, 0, 0))
    page.insert_text(fitz.Point(nx+10, ny+65), "3. 材質: SUS304",
                     fontname="japan", fontsize=8, color=(0, 0, 0))
    page.insert_text(fitz.Point(nx+10, ny+85), "4. 配管仕様: JIS 10K",
                     fontname="japan", fontsize=8, color=(0, 0, 0))

    # タイトルブロック
    tx, ty = 850, 680
    shape = page.new_shape()
    shape.draw_rect(fitz.Rect(tx, ty, tx+300, ty+100))
    shape.draw_line(fitz.Point(tx, ty+35), fitz.Point(tx+300, ty+35))
    shape.draw_line(fitz.Point(tx, ty+65), fitz.Point(tx+300, ty+65))
    shape.draw_line(fitz.Point(tx+150, ty), fitz.Point(tx+150, ty+100))
    shape.finish(color=(0, 0, 0), width=0.5)
    shape.commit()

    page.insert_text(fitz.Point(tx+10, ty+20), "工事名称",
                     fontname="japan", fontsize=8, color=(0, 0, 0))
    page.insert_text(fitz.Point(tx+155, ty+20), "テストプラント P&ID",
                     fontname="japan", fontsize=8, color=(0, 0, 0))
    page.insert_text(fitz.Point(tx+10, ty+55), "図面番号",
                     fontname="japan", fontsize=8, color=(0, 0, 0))
    draw_shx_text(page, "PID-001-A", tx+155, ty+40, 8)
    page.insert_text(fitz.Point(tx+10, ty+85), "改訂",
                     fontname="japan", fontsize=8, color=(0, 0, 0))
    draw_shx_text(page, "R1", tx+155, ty+70, 8)

    doc.save(output_path)
    doc.close()
    print(f"日本語テストPDF生成: {output_path}")

    # 分析
    doc = fitz.open(output_path)
    page = doc[0]
    text_words = page.get_text('words')
    drawings = page.get_drawings()
    print(f"  テキストスパン: {len(text_words)}")
    print(f"  描画パス: {len(drawings)}")
    doc.close()


if __name__ == '__main__':
    create_ja_test_pdf()
