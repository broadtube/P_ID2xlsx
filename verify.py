"""
PDF vs Excel 比較検証スクリプト

PDFを画像化し、Excelも画像化して並べて比較する。
"""

import sys
from pathlib import Path
import fitz  # PyMuPDF


def pdf_to_png(pdf_path: str, output_path: str = None, dpi: int = 150, page_num: int = 0):
    """PDFをPNG画像に変換（page_num: 0-based）"""
    pdf_path = Path(pdf_path)
    if output_path is None:
        output_path = pdf_path.with_name(pdf_path.stem + '_pdf.png')

    doc = fitz.open(str(pdf_path))
    page = doc[page_num]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat)
    pix.save(str(output_path))
    doc.close()
    print(f"PDF画像: {output_path} ({pix.width}x{pix.height}px)")
    return str(output_path)


def xlsx_to_png(xlsx_path: str, output_path: str = None):
    """xlsxをPNG画像に変換（Excel COM使用）"""
    xlsx_path = Path(xlsx_path).resolve()
    if output_path is None:
        output_path = xlsx_path.with_name(xlsx_path.stem + '_xlsx.png')

    temp_pdf = xlsx_path.with_name(xlsx_path.stem + '_temp.pdf')

    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False

        try:
            try:
                excel.DisplayAlerts = False
            except Exception:
                pass  # 一部環境ではDisplayAlertsが設定できない
            wb = excel.Workbooks.Open(str(xlsx_path))
            # xlTypePDF = 0
            wb.ExportAsFixedFormat(0, str(temp_pdf))
            wb.Close(False)
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()

        if temp_pdf.exists():
            png_path = pdf_to_png(str(temp_pdf), str(output_path))
            temp_pdf.unlink()
            return png_path
        else:
            print("Excel COM: PDF出力に失敗しました")
            return None

    except ImportError:
        print("win32comが必要です: pip install pywin32")
        return None
    except Exception as e:
        print(f"Excel COM エラー: {e}")
        if temp_pdf.exists():
            temp_pdf.unlink()
        return None


def compare_side_by_side(pdf_png: str, xlsx_png: str, output_path: str = None):
    """2つの画像を横に並べて比較画像を生成"""
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        print("Pillow未インストール: pip install Pillow")
        return None

    img1 = Image.open(pdf_png)
    img2 = Image.open(xlsx_png)

    # サイズを揃える
    h = max(img1.height, img2.height)
    w1 = img1.width
    w2 = img2.width

    # 横に並べる（間に区切り線）
    gap = 20
    combined = Image.new('RGB', (w1 + gap + w2, h + 40), 'white')
    combined.paste(img1, (0, 40))
    combined.paste(img2, (w1 + gap, 40))

    # ラベル追加
    draw = ImageDraw.Draw(combined)
    draw.text((w1 // 2 - 30, 10), "PDF (原本)", fill='red')
    draw.text((w1 + gap + w2 // 2 - 30, 10), "Excel (変換結果)", fill='blue')
    # 区切り線
    draw.line([(w1 + gap // 2, 0), (w1 + gap // 2, h + 40)], fill='gray', width=2)

    if output_path is None:
        output_path = Path(pdf_png).with_name('comparison.png')
    combined.save(str(output_path))
    print(f"比較画像: {output_path}")
    return str(output_path)


def main():
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else None
    xlsx_path = sys.argv[2] if len(sys.argv) > 2 else None
    page_num = int(sys.argv[3]) - 1 if len(sys.argv) > 3 else 0  # 1-based → 0-based

    if not pdf_path:
        pdfs = list(Path('.').glob('*.pdf'))
        if pdfs:
            pdf_path = str(pdfs[0])

    if not xlsx_path:
        xlsxs = list(Path('.').glob('*.xlsx'))
        if xlsxs:
            xlsx_path = str(xlsxs[0])

    if not pdf_path or not xlsx_path:
        print("使用法: python verify.py <input.pdf> <output.xlsx> [page_num]")
        sys.exit(1)

    print("=== PDF → PNG ===")
    pdf_png = pdf_to_png(pdf_path, page_num=page_num)

    print("\n=== Excel → PNG ===")
    xlsx_png = xlsx_to_png(xlsx_path)

    if pdf_png and xlsx_png:
        print("\n=== 比較画像生成 ===")
        compare_side_by_side(pdf_png, xlsx_png)


if __name__ == '__main__':
    main()
