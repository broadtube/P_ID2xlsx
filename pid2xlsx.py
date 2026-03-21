"""
P&ID PDF → Excel (.xlsx) 変換スクリプト

PDFからテキスト・図形（線・円・矩形）を抽出し、
Excelのテキストボックス・シェイプとして再現する。
"""

import sys
import zipfile
import shutil
import tempfile
from pathlib import Path
from xml.etree.ElementTree import Element, SubElement, tostring, register_namespace

import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

# --- 名前空間 ---
NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_SHEET = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

register_namespace('', NS_XDR)
register_namespace('a', NS_A)
register_namespace('r', NS_R)

# --- 定数 ---
PDF_TO_EMU = 914400 / 72  # 12700 EMU per PDF point

# Excel デフォルトセルサイズ（EMU）- Excel COMで実測済み
COL_WIDTH_CHARS = 2.14    # 列幅（文字数）→ Excel実測: 12.75pt
ROW_HEIGHT_PT = 7.5       # 行高さ（pt）
# Excel実測: width=2.14 → 12.75pt → 161925 EMU
COL_WIDTH_EMU = 161925
# 7.5pt × 12700 = 95250 EMU
ROW_HEIGHT_EMU = 95250


def pdf_pt_to_emu(pt: float) -> int:
    return int(pt * PDF_TO_EMU)


def make_coord_transform(page):
    """ページの回転を考慮した座標変換関数を返す"""
    rotation = page.rotation
    mbox = page.mediabox
    mw, mh = mbox.width, mbox.height

    if rotation == 0:
        def transform(x, y):
            return x, y
    elif rotation == 90:
        def transform(x, y):
            return mh - y, x
    elif rotation == 180:
        def transform(x, y):
            return mw - x, mh - y
    elif rotation == 270:
        def transform(x, y):
            return y, mw - x
    else:
        def transform(x, y):
            return x, y

    return transform


def coord_to_anchor(x_pt: float, y_pt: float) -> tuple:
    """表示空間座標 → (col, colOff, row, rowOff)"""
    x_emu = pdf_pt_to_emu(x_pt)
    y_emu = pdf_pt_to_emu(y_pt)
    col = int(x_emu // COL_WIDTH_EMU)
    col_off = int(x_emu % COL_WIDTH_EMU)
    row = int(y_emu // ROW_HEIGHT_EMU)
    row_off = int(y_emu % ROW_HEIGHT_EMU)
    return col, col_off, row, row_off


def color_tuple_to_hex(color) -> str | None:
    if color is None:
        return None
    return '{:02X}{:02X}{:02X}'.format(
        int(color[0] * 255), int(color[1] * 255), int(color[2] * 255)
    )


# --- XML要素生成 ---
def make_marker(tag: str, col: int, col_off: int, row: int, row_off: int) -> Element:
    """xdr:from / xdr:to マーカー要素を生成"""
    m = Element(f'{{{NS_XDR}}}{tag}')
    SubElement(m, f'{{{NS_XDR}}}col').text = str(col)
    SubElement(m, f'{{{NS_XDR}}}colOff').text = str(col_off)
    SubElement(m, f'{{{NS_XDR}}}row').text = str(row)
    SubElement(m, f'{{{NS_XDR}}}rowOff').text = str(row_off)
    return m


def make_freeform_geom(items: list, x1: float, y1: float, x2: float, y2: float,
                       closePath: bool = False) -> Element:
    """カスタムジオメトリ（フリーフォーム）要素を生成"""
    GEOM_SCALE = 100000  # EMU内部座標のスケール
    w = max(x2 - x1, 0.1)
    h = max(y2 - y1, 0.1)

    cust_geom = Element(f'{{{NS_A}}}custGeom')
    SubElement(cust_geom, f'{{{NS_A}}}avLst')
    SubElement(cust_geom, f'{{{NS_A}}}gdLst')
    SubElement(cust_geom, f'{{{NS_A}}}ahLst')
    SubElement(cust_geom, f'{{{NS_A}}}cxnLst')
    a_rect = SubElement(cust_geom, f'{{{NS_A}}}rect')
    a_rect.set('l', '0')
    a_rect.set('t', '0')
    a_rect.set('r', str(GEOM_SCALE))
    a_rect.set('b', str(GEOM_SCALE))

    path_lst = SubElement(cust_geom, f'{{{NS_A}}}pathLst')
    path = SubElement(path_lst, f'{{{NS_A}}}path')
    path.set('w', str(GEOM_SCALE))
    path.set('h', str(GEOM_SCALE))

    def pt_to_local(px, py):
        """PDF座標 → ジオメトリ内ローカル座標"""
        lx = int((px - x1) / w * GEOM_SCALE)
        ly = int((py - y1) / h * GEOM_SCALE)
        return max(0, min(GEOM_SCALE, lx)), max(0, min(GEOM_SCALE, ly))

    def add_pt(parent, px, py):
        pt = SubElement(parent, f'{{{NS_A}}}pt')
        lx, ly = pt_to_local(px, py)
        pt.set('x', str(lx))
        pt.set('y', str(ly))

    first_move = True
    for item in items:
        if item[0] == 'l':
            # 直線: ('l', Point, Point)
            p1, p2 = item[1], item[2]
            if first_move:
                move_to = SubElement(path, f'{{{NS_A}}}moveTo')
                add_pt(move_to, p1.x, p1.y)
                first_move = False
            ln_to = SubElement(path, f'{{{NS_A}}}lnTo')
            add_pt(ln_to, p2.x, p2.y)
        elif item[0] == 'c':
            # ベジェ曲線: ('c', Point, Point, Point, Point)
            p1, c1, c2, p2 = item[1], item[2], item[3], item[4]
            if first_move:
                move_to = SubElement(path, f'{{{NS_A}}}moveTo')
                add_pt(move_to, p1.x, p1.y)
                first_move = False
            cubic = SubElement(path, f'{{{NS_A}}}cubicBezTo')
            add_pt(cubic, c1.x, c1.y)
            add_pt(cubic, c2.x, c2.y)
            add_pt(cubic, p2.x, p2.y)

    if closePath:
        SubElement(path, f'{{{NS_A}}}close')

    return cust_geom


def make_valve_geom(vertical: bool) -> Element:
    """バルブ（ボウタイ/砂時計）のカスタムジオメトリを生成
    vertical=True: 三角が上下に向き合う（⧖）
    vertical=False: 三角が左右に向き合う（⧗）
    """
    S = 100000
    cust_geom = Element(f'{{{NS_A}}}custGeom')
    SubElement(cust_geom, f'{{{NS_A}}}avLst')
    SubElement(cust_geom, f'{{{NS_A}}}gdLst')
    SubElement(cust_geom, f'{{{NS_A}}}ahLst')
    SubElement(cust_geom, f'{{{NS_A}}}cxnLst')
    a_rect = SubElement(cust_geom, f'{{{NS_A}}}rect')
    a_rect.set('l', '0')
    a_rect.set('t', '0')
    a_rect.set('r', str(S))
    a_rect.set('b', str(S))

    path_lst = SubElement(cust_geom, f'{{{NS_A}}}pathLst')

    def add_pt(parent, x, y):
        pt = SubElement(parent, f'{{{NS_A}}}pt')
        pt.set('x', str(x))
        pt.set('y', str(y))

    if vertical:
        # 上三角（頂点が中央下向き）+ 下三角（頂点が中央上向き）
        # Triangle 1: (0,0) → (S,0) → (S/2,S/2) → close
        p1 = SubElement(path_lst, f'{{{NS_A}}}path')
        p1.set('w', str(S))
        p1.set('h', str(S))
        m = SubElement(p1, f'{{{NS_A}}}moveTo')
        add_pt(m, 0, 0)
        l = SubElement(p1, f'{{{NS_A}}}lnTo')
        add_pt(l, S, 0)
        l = SubElement(p1, f'{{{NS_A}}}lnTo')
        add_pt(l, S // 2, S // 2)
        SubElement(p1, f'{{{NS_A}}}close')
        # Triangle 2: (0,S) → (S,S) → (S/2,S/2) → close
        p2 = SubElement(path_lst, f'{{{NS_A}}}path')
        p2.set('w', str(S))
        p2.set('h', str(S))
        m = SubElement(p2, f'{{{NS_A}}}moveTo')
        add_pt(m, 0, S)
        l = SubElement(p2, f'{{{NS_A}}}lnTo')
        add_pt(l, S, S)
        l = SubElement(p2, f'{{{NS_A}}}lnTo')
        add_pt(l, S // 2, S // 2)
        SubElement(p2, f'{{{NS_A}}}close')
    else:
        # 左三角（頂点が中央右向き）+ 右三角（頂点が中央左向き）
        # Triangle 1: (0,0) → (0,S) → (S/2,S/2) → close
        p1 = SubElement(path_lst, f'{{{NS_A}}}path')
        p1.set('w', str(S))
        p1.set('h', str(S))
        m = SubElement(p1, f'{{{NS_A}}}moveTo')
        add_pt(m, 0, 0)
        l = SubElement(p1, f'{{{NS_A}}}lnTo')
        add_pt(l, 0, S)
        l = SubElement(p1, f'{{{NS_A}}}lnTo')
        add_pt(l, S // 2, S // 2)
        SubElement(p1, f'{{{NS_A}}}close')
        # Triangle 2: (S,0) → (S,S) → (S/2,S/2) → close
        p2 = SubElement(path_lst, f'{{{NS_A}}}path')
        p2.set('w', str(S))
        p2.set('h', str(S))
        m = SubElement(p2, f'{{{NS_A}}}moveTo')
        add_pt(m, S, 0)
        l = SubElement(p2, f'{{{NS_A}}}lnTo')
        add_pt(l, S, S)
        l = SubElement(p2, f'{{{NS_A}}}lnTo')
        add_pt(l, S // 2, S // 2)
        SubElement(p2, f'{{{NS_A}}}close')

    return cust_geom


def make_shape_xml(shape_id: int, name: str, prst: str,
                   x1: float, y1: float, x2: float, y2: float,
                   line_width_emu: int = 12700, line_color: str = '000000',
                   fill_color: str = None, text: str = None,
                   font_size: float = None, no_line: bool = False,
                   path_items: list = None, closePath: bool = False,
                   text_rotation: int = 0, shape_rot: int = 0) -> Element:
    """TwoCellAnchor + Shape XML要素を生成"""
    anchor = Element(f'{{{NS_XDR}}}twoCellAnchor')

    # 直線の場合、描画方向を判定してflipを設定
    need_flipH = False
    need_flipV = False
    if prst == 'line':
        # lineプリセットはfrom→toの対角線（左上→右下）を描画
        # 実際の方向が異なる場合はflipで対応
        if x1 > x2:
            need_flipH = True
        if y1 > y2:
            need_flipV = True

    # from / to マーカー（常に左上→右下の順に正規化）
    ax1, ay1 = min(x1, x2), min(y1, y2)
    ax2, ay2 = max(x1, x2), max(y1, y2)
    c1, co1, r1, ro1 = coord_to_anchor(ax1, ay1)
    c2, co2, r2, ro2 = coord_to_anchor(ax2, ay2)

    # 最小サイズ確保（線以外の図形のみ）
    # 線の場合、垂直方向にpadを追加すると斜めになるため除外
    if prst != 'line':
        if c1 == c2 and co1 >= co2:
            co2 = co1 + 9525  # 0.75pt
        if r1 == r2 and ro1 >= ro2:
            ro2 = ro1 + 9525

    anchor.append(make_marker('from', c1, co1, r1, ro1))
    anchor.append(make_marker('to', c2, co2, r2, ro2))

    # sp 要素
    sp = SubElement(anchor, f'{{{NS_XDR}}}sp')

    # nvSpPr
    nv = SubElement(sp, f'{{{NS_XDR}}}nvSpPr')
    cnv_pr = SubElement(nv, f'{{{NS_XDR}}}cNvPr')
    cnv_pr.set('id', str(shape_id))
    cnv_pr.set('name', name)
    SubElement(nv, f'{{{NS_XDR}}}cNvSpPr')

    # spPr
    sp_pr = SubElement(sp, f'{{{NS_XDR}}}spPr')

    # xfrm: flipH/flipV/rotation はa:xfrm要素に設定（spPrではない）
    if need_flipH or need_flipV or shape_rot:
        xfrm = SubElement(sp_pr, f'{{{NS_A}}}xfrm')
        if need_flipH:
            xfrm.set('flipH', '1')
        if need_flipV:
            xfrm.set('flipV', '1')
        if shape_rot:
            xfrm.set('rot', str(shape_rot))

    # ジオメトリ：カスタムまたはプリセット
    if prst == 'flowChartCollate':
        # バルブ: カスタムジオメトリで向きを制御
        vertical = (ay2 - ay1) > (ax2 - ax1)  # display bboxが縦長なら上下向き
        sp_pr.append(make_valve_geom(vertical))
    elif path_items is not None:
        sp_pr.append(make_freeform_geom(path_items, ax1, ay1, ax2, ay2, closePath))
    else:
        prst_geom = SubElement(sp_pr, f'{{{NS_A}}}prstGeom')
        prst_geom.set('prst', prst)
        SubElement(prst_geom, f'{{{NS_A}}}avLst')

    # 塗りつぶし
    if fill_color:
        solid_fill = SubElement(sp_pr, f'{{{NS_A}}}solidFill')
        srgb = SubElement(solid_fill, f'{{{NS_A}}}srgbClr')
        srgb.set('val', fill_color)
    else:
        SubElement(sp_pr, f'{{{NS_A}}}noFill')

    # 線
    if no_line:
        ln = SubElement(sp_pr, f'{{{NS_A}}}ln')
        SubElement(ln, f'{{{NS_A}}}noFill')
    else:
        ln = SubElement(sp_pr, f'{{{NS_A}}}ln')
        ln.set('w', str(line_width_emu))
        sf = SubElement(ln, f'{{{NS_A}}}solidFill')
        srgb = SubElement(sf, f'{{{NS_A}}}srgbClr')
        srgb.set('val', line_color)

    # テキスト
    if text:
        tx_body = SubElement(sp, f'{{{NS_XDR}}}txBody')
        body_pr = SubElement(tx_body, f'{{{NS_A}}}bodyPr')
        body_pr.set('wrap', 'none')
        body_pr.set('lIns', '0')
        body_pr.set('tIns', '0')
        body_pr.set('rIns', '0')
        body_pr.set('bIns', '0')
        # テキスト回転
        if text_rotation == 90:
            body_pr.set('vert', 'vert')  # 90° CW (top to bottom)
        elif text_rotation == -90 or text_rotation == 270:
            body_pr.set('vert', 'vert270')  # 270° CW (bottom to top)
        SubElement(tx_body, f'{{{NS_A}}}lstStyle')
        p = SubElement(tx_body, f'{{{NS_A}}}p')
        r = SubElement(p, f'{{{NS_A}}}r')
        rp = SubElement(r, f'{{{NS_A}}}rPr')
        rp.set('lang', 'en-US')
        sz = int((font_size or 6.0) * 100)
        rp.set('sz', str(sz))
        solid = SubElement(rp, f'{{{NS_A}}}solidFill')
        srgb = SubElement(solid, f'{{{NS_A}}}srgbClr')
        srgb.set('val', line_color)
        latin = SubElement(rp, f'{{{NS_A}}}latin')
        latin.set('typeface', 'Arial')
        t = SubElement(r, f'{{{NS_A}}}t')
        t.text = text

    # clientData
    SubElement(anchor, f'{{{NS_XDR}}}clientData')

    return anchor


# --- PDF解析 ---
def _snap_line(x1, y1, x2, y2, threshold=1.5):
    """ほぼ垂直/水平な線を正確にスナップ"""
    if abs(x2 - x1) < threshold and abs(y2 - y1) > threshold:
        avg = (x1 + x2) / 2
        return avg, y1, avg, y2
    elif abs(y2 - y1) < threshold and abs(x2 - x1) > threshold:
        avg = (y1 + y2) / 2
        return x1, avg, x2, avg
    return x1, y1, x2, y2


def _is_line_diagonal(p1, p2, threshold=1.5):
    """線が対角線（水平でも垂直でもない）かどうか判定"""
    dx = abs(p2.x - p1.x)
    dy = abs(p2.y - p1.y)
    return dx > threshold and dy > threshold


def _triangle_rotation(tpts):
    """三角形の3頂点から、頂点(apex)が指す方向に応じたExcel回転値を返す。
    Excelのtriangleプリセットはデフォルトで頂点が上を向く。
    回転値: 0=上, 5400000=右, 10800000=下, 16200000=左"""
    import math
    # 最長辺（底辺）を見つけ、その対頂点がapex
    edges = []
    for i in range(3):
        j = (i + 1) % 3
        dist = (tpts[i][0] - tpts[j][0]) ** 2 + (tpts[i][1] - tpts[j][1]) ** 2
        edges.append((dist, i, j))
    edges.sort(reverse=True)
    # 最長辺の2頂点のインデックス
    _, bi, bj = edges[0]
    apex_idx = 3 - bi - bj  # 残りの頂点
    apex = tpts[apex_idx]
    base_mid = ((tpts[bi][0] + tpts[bj][0]) / 2, (tpts[bi][1] + tpts[bj][1]) / 2)
    # base_midからapexへの方向 = apexが指す方向
    dx = apex[0] - base_mid[0]
    dy = apex[1] - base_mid[1]
    angle = math.degrees(math.atan2(dy, dx))
    # Y-down座標系: angle 0=右, 90=下, -90=上, 180=左
    # Excel triangle: apex上=0, 右=90, 下=180, 左=270
    if abs(angle + 90) < 45:  # 上 (angle ~ -90)
        return 0
    elif abs(angle) < 45:  # 右 (angle ~ 0)
        return 5400000
    elif abs(angle - 90) < 45:  # 下 (angle ~ 90)
        return 10800000
    else:  # 左
        return 16200000


def _homeplate_rotation(tpts):
    """5頂点のホームベース形状の向き（尖った方向）を判定。
    ExcelのhomePlateプリセットはデフォルトで尖端が右を向く。
    回転値: 0=右, 5400000=下, 10800000=左, 16200000=上
    ホームベースでない場合はNoneを返す。"""
    # 5頂点から矩形部分と尖端を特定
    xs = [p[0] for p in tpts]
    ys = [p[1] for p in tpts]
    # 各座標値の出現頻度を調べる
    from collections import Counter
    x_counts = Counter(round(x, 0) for x in xs)
    y_counts = Counter(round(y, 0) for y in ys)
    # ホームベースの特徴: 矩形の4隅のうち2つが同じx（またはy）を共有し、
    # 尖端はその軸の中間にある
    # 尖端 = 他の頂点と同じx/yを共有しない頂点
    tip = None
    for p in tpts:
        rx = round(p[0], 0)
        ry = round(p[1], 0)
        if x_counts[rx] == 1 and y_counts[ry] == 1:
            tip = p
            break
    if tip is None:
        # 尖端がx方向に一意な場合
        for p in tpts:
            rx = round(p[0], 0)
            if x_counts[rx] == 1:
                tip = p
                break
    if tip is None:
        for p in tpts:
            ry = round(p[1], 0)
            if y_counts[ry] == 1:
                tip = p
                break
    if tip is None:
        return None
    # 重心から尖端への方向
    cx = sum(p[0] for p in tpts) / 5
    cy = sum(p[1] for p in tpts) / 5
    dx = tip[0] - cx
    dy = tip[1] - cy
    import math
    angle = math.degrees(math.atan2(dy, dx))
    # Excel homePlate: 尖端が右=0°
    # angle: 0=右, 90=下, -90=上, 180=左
    if abs(angle) < 45:  # 右
        return 0
    elif abs(angle - 90) < 45:  # 下
        return 5400000
    elif abs(angle + 90) < 45:  # 上
        return 16200000
    else:  # 左
        return 10800000


def _is_valve_pattern(items, rect):
    """3直線がバルブ（ボウタイ/X型）パターンかどうか判定
    条件: 3直線、4端点、適切なサイズ、少なくとも2本が対角線"""
    line_items = [i for i in items if i[0] == 'l']
    if len(line_items) != 3 or len(items) != 3:
        return False

    mw = rect.x1 - rect.x0
    mh = rect.y1 - rect.y0
    if min(mw, mh) <= 5 or max(mw, mh) >= 30:
        return False
    if max(mw, mh) / max(min(mw, mh), 0.1) >= 2.5:
        return False

    # 端点が4つの角に集まるかチェック
    pts = set()
    for li in line_items:
        pts.add((round(li[1].x, 1), round(li[1].y, 1)))
        pts.add((round(li[2].x, 1), round(li[2].y, 1)))
    if len(pts) != 4:
        return False

    # 少なくとも2本が対角線であること（3辺矩形との区別）
    diag_count = sum(1 for li in line_items if _is_line_diagonal(li[1], li[2]))
    if diag_count < 2:
        return False

    return True


def transform_items(items, transform):
    """描画パスのアイテム座標を変換し、角度スナップを適用する"""
    import fitz
    new_items = []
    for item in items:
        if item[0] == 'l':
            tx1, ty1 = transform(item[1].x, item[1].y)
            tx2, ty2 = transform(item[2].x, item[2].y)
            # 角度スナップ
            tx1, ty1, tx2, ty2 = _snap_line(tx1, ty1, tx2, ty2)
            p1 = fitz.Point(tx1, ty1)
            p2 = fitz.Point(tx2, ty2)
            new_items.append(('l', p1, p2))
        elif item[0] == 'c':
            p1 = fitz.Point(*transform(item[1].x, item[1].y))
            c1 = fitz.Point(*transform(item[2].x, item[2].y))
            c2 = fitz.Point(*transform(item[3].x, item[3].y))
            p2 = fitz.Point(*transform(item[4].x, item[4].y))
            new_items.append(('c', p1, c1, c2, p2))
        else:
            new_items.append(item)
    return new_items


def classify_drawing(drawing: dict, transform=None) -> dict | None:
    """PDF描画パスを分類し、図形情報を返す"""
    items = drawing['items']
    rect = drawing['rect']
    color = drawing.get('color', (0, 0, 0))
    fill = drawing.get('fill')
    width = drawing.get('width', 1.0)
    closePath = drawing.get('closePath', False)

    if not items:
        return None

    line_color = color_tuple_to_hex(color) or '000000'
    fill_color = color_tuple_to_hex(fill)
    line_width_emu = max(int(width * PDF_TO_EMU), 3175)  # 最小0.25pt

    # 座標変換
    if transform:
        tx1, ty1 = transform(rect.x0, rect.y0)
        tx2, ty2 = transform(rect.x1, rect.y1)
        # 変換後のbounding boxを正規化
        x1, x2 = min(tx1, tx2), max(tx1, tx2)
        y1, y2 = min(ty1, ty2), max(ty1, ty2)
    else:
        x1, y1, x2, y2 = rect.x0, rect.y0, rect.x1, rect.y1

    base = dict(x1=x1, y1=y1, x2=x2, y2=y2,
                line_color=line_color, fill_color=fill_color, line_width=line_width_emu)

    # 矩形
    if any(i[0] in ('re', 'qu') for i in items):
        return dict(type='rect', **base)

    # 円/楕円（4+ ベジェ曲線、アスペクト比不問）
    curve_items = [i for i in items if i[0] == 'c']
    if len(curve_items) >= 4:
        w, h = x2 - x1, y2 - y1
        if w > 0 and h > 0:
            return dict(type='ellipse', **base)

    # ベジェ曲線を含むパス → カスタムジオメトリ
    if curve_items:
        t_items = transform_items(items, transform) if transform else items
        return dict(type='freeform', items=t_items, closePath=closePath, **base)

    # 直線パス
    line_items = [i for i in items if i[0] == 'l']
    if line_items:
        # 直線のみで構成された矩形の検出（3辺または4辺、全てH/V）
        if len(line_items) >= 3 and len(line_items) == len(items):
            all_hv = all(
                abs(li[1].x - li[2].x) < 1.5 or abs(li[1].y - li[2].y) < 1.5
                for li in line_items
            )
            if all_hv and closePath:
                return dict(type='rect', **base)

        # バルブ検出: 3直線でボウタイ（X型）を形成
        # 向きはmake_valve_geom(vertical)で制御（display bboxの縦横比で判定）
        if _is_valve_pattern(items, rect):
            return dict(type='valve', **base)

        # 三角形検出: 3直線、3頂点（バルブは4頂点なので除外済み）
        if len(line_items) == 3 and len(items) == 3:
            pts_raw = []
            for li in line_items:
                pts_raw.append((li[1].x, li[1].y))
                pts_raw.append((li[2].x, li[2].y))
            pts = list(set((round(px, 1), round(py, 1)) for px, py in pts_raw))
            if len(pts) == 3:
                # 三角形の頂点を表示座標に変換
                if transform:
                    tpts = [transform(px, py) for px, py in pts]
                else:
                    tpts = list(pts)
                # 頂点の向きを判定（頂点から対辺の中点への方向）
                tri_rot = _triangle_rotation(tpts)
                # 矢印三角形は塗りつぶしで描画
                return dict(type='triangle', shape_rot=tri_rot,
                            x1=x1, y1=y1, x2=x2, y2=y2,
                            line_color=line_color, fill_color=line_color,
                            line_width=line_width_emu)

        # ホームベース検出: 5直線、5頂点（矩形+一辺が三角形）
        # 条件: 少なくとも2本の対角線を含む（全H/VのL字型パスを除外）
        if len(line_items) == 5 and len(items) == 5:
            diag_count = sum(1 for li in line_items if _is_line_diagonal(li[1], li[2]))
            if diag_count >= 2:
                pts_raw = []
                for li in line_items:
                    pts_raw.append((li[1].x, li[1].y))
                    pts_raw.append((li[2].x, li[2].y))
                pts = list(set((round(px, 1), round(py, 1)) for px, py in pts_raw))
                if len(pts) == 5:
                    if transform:
                        tpts = [transform(px, py) for px, py in pts]
                    else:
                        tpts = list(pts)
                    hp_rot = _homeplate_rotation(tpts)
                    if hp_rot is not None:
                        return dict(type='homePlate', shape_rot=hp_rot, **base)

        # 単一直線 - 実際の始点・終点を使用
        if len(line_items) == 1 and len(items) == 1:
            p1, p2 = items[0][1], items[0][2]
            if transform:
                lx1, ly1 = transform(p1.x, p1.y)
                lx2, ly2 = transform(p2.x, p2.y)
            else:
                lx1, ly1 = p1.x, p1.y
                lx2, ly2 = p2.x, p2.y
            # ゼロ長の線（ドット）→ 小さい塗りつぶし円で表現
            if abs(lx2 - lx1) < 0.1 and abs(ly2 - ly1) < 0.1:
                dot_r = max(width * 0.5, 0.5)  # ドットの半径
                return dict(type='dot',
                            x1=lx1 - dot_r, y1=ly1 - dot_r,
                            x2=lx1 + dot_r, y2=ly1 + dot_r,
                            line_color=line_color, fill_color=line_color,
                            line_width=line_width_emu)
            # 角度スナップ
            lx1, ly1, lx2, ly2 = _snap_line(lx1, ly1, lx2, ly2)
            return dict(type='line', x1=lx1, y1=ly1, x2=lx2, y2=ly2,
                        line_color=line_color, fill_color=None, line_width=line_width_emu)
        # 複数直線 → 個別の線に分解（freeformのbbox丸め問題を回避）
        lines = []
        for li in line_items:
            p1, p2 = li[1], li[2]
            if transform:
                lx1, ly1 = transform(p1.x, p1.y)
                lx2, ly2 = transform(p2.x, p2.y)
            else:
                lx1, ly1 = p1.x, p1.y
                lx2, ly2 = p2.x, p2.y
            lx1, ly1, lx2, ly2 = _snap_line(lx1, ly1, lx2, ly2)
            lines.append(dict(type='line', x1=lx1, y1=ly1, x2=lx2, y2=ly2,
                              line_color=line_color, fill_color=None, line_width=line_width_emu))
        return dict(type='multi_line', lines=lines)

    return None


def extract_text_spans(page, transform=None) -> list:
    import math
    rotation = page.rotation
    spans = []
    blocks = page.get_text('dict')['blocks']
    for block in blocks:
        if block['type'] != 0:
            continue
        for line in block['lines']:
            dir_ = line.get('dir', (1, 0))
            for span in line['spans']:
                text = span['text'].strip()
                if not text:
                    continue
                bbox = span['bbox']
                if transform:
                    tx1, ty1 = transform(bbox[0], bbox[1])
                    tx2, ty2 = transform(bbox[2], bbox[3])
                    sx1, sx2 = min(tx1, tx2), max(tx1, tx2)
                    sy1, sy2 = min(ty1, ty2), max(ty1, ty2)
                else:
                    sx1, sy1, sx2, sy2 = bbox[0], bbox[1], bbox[2], bbox[3]

                # テキスト回転角度（表示空間）
                # 方向ベクトルをページ回転で変換
                dx, dy = dir_
                if rotation == 270:
                    ddx, ddy = dy, -dx
                elif rotation == 90:
                    ddx, ddy = -dy, dx
                elif rotation == 180:
                    ddx, ddy = -dx, -dy
                else:
                    ddx, ddy = dx, dy
                # 表示空間での角度（度）
                angle_deg = math.degrees(math.atan2(ddy, ddx))
                # 0°に近い場合は回転なし
                if abs(angle_deg) < 1:
                    text_rot = 0
                else:
                    text_rot = round(angle_deg)

                # テキスト色
                color_int = span.get('color', 0)
                text_color = '{:02X}{:02X}{:02X}'.format(
                    (color_int >> 16) & 0xFF,
                    (color_int >> 8) & 0xFF,
                    color_int & 0xFF)

                spans.append(dict(text=text, x1=sx1, y1=sy1,
                                  x2=sx2, y2=sy2,
                                  size=span['size'], font=span['font'],
                                  rotation=text_rot, color=text_color))
    return spans


def _is_valve_edge_line(drawing, valve_rects):
    """単一直線がバルブの辺（三角形の一辺）かどうか判定
    バルブbboxの一辺に沿った垂直/水平線のみ抑制する"""
    items = drawing['items']
    if len(items) != 1 or items[0][0] != 'l':
        return False
    p1, p2 = items[0][1], items[0][2]
    tol = 0.5
    for vx0, vy0, vx1, vy1 in valve_rects:
        # 垂直線: x座標がバルブの左辺or右辺と一致し、y範囲がバルブ内
        if abs(p1.x - p2.x) < tol:
            if (abs(p1.x - vx0) < tol or abs(p1.x - vx1) < tol):
                y_min, y_max = min(p1.y, p2.y), max(p1.y, p2.y)
                if abs(y_min - vy0) < tol and abs(y_max - vy1) < tol:
                    return True
        # 水平線: y座標がバルブの上辺or下辺と一致し、x範囲がバルブ内
        if abs(p1.y - p2.y) < tol:
            if (abs(p1.y - vy0) < tol or abs(p1.y - vy1) < tol):
                x_min, x_max = min(p1.x, p2.x), max(p1.x, p2.x)
                if abs(x_min - vx0) < tol and abs(x_max - vx1) < tol:
                    return True
    return False


def build_drawing_xml(page) -> tuple:
    """ページからDrawing XMLを生成。(xml_bytes, shape_count) を返す"""
    root = Element(f'{{{NS_XDR}}}wsDr')
    transform = make_coord_transform(page)
    rotation = page.rotation
    drawings = page.get_drawings()

    # Pass 1: バルブを検出してbounding box（mediabox座標）を記録
    valve_rects = []
    for d in drawings:
        if _is_valve_pattern(d['items'], d['rect']):
            rect = d['rect']
            valve_rects.append((rect.x0, rect.y0, rect.x1, rect.y1))

    shape_id = 2
    count = 0

    # Pass 2: 図形変換（バルブ辺の単一直線を抑制）
    for d in drawings:
        # バルブの辺（三角形の一辺）を抑制
        if _is_valve_edge_line(d, valve_rects):
            continue

        info = classify_drawing(d, transform)
        if info is None:
            continue

        # 複数直線 → 個別の線として展開
        if info.get('type') == 'multi_line':
            for line_info in info['lines']:
                elem = make_shape_xml(
                    shape_id=shape_id,
                    name=f'line_{shape_id}',
                    prst='line',
                    x1=line_info['x1'], y1=line_info['y1'],
                    x2=line_info['x2'], y2=line_info['y2'],
                    line_width_emu=line_info['line_width'],
                    line_color=line_info['line_color'],
                    fill_color=None,
                )
                root.append(elem)
                shape_id += 1
                count += 1
            continue

        dx = abs(info['x2'] - info['x1'])
        dy = abs(info['y2'] - info['y1'])
        if info['type'] not in ('line', 'freeform', 'dot') and dx < 2 and dy < 2:
            continue

        # 図形タイプ → Excelプリセットジオメトリ
        shape_type = info['type']
        if shape_type == 'freeform':
            prst = 'rect'
        elif shape_type == 'valve':
            prst = 'flowChartCollate'
        elif shape_type == 'dot':
            prst = 'ellipse'
        elif shape_type in ('triangle', 'homePlate'):
            prst = shape_type
        else:
            prst = shape_type
        path_items = info.get('items') if shape_type == 'freeform' else None
        closePath = info.get('closePath', False)
        is_dot = (shape_type == 'dot')

        elem = make_shape_xml(
            shape_id=shape_id,
            name=f'{info["type"]}_{shape_id}',
            prst=prst,
            x1=info['x1'], y1=info['y1'],
            x2=info['x2'], y2=info['y2'],
            line_width_emu=info['line_width'],
            line_color=info['line_color'],
            fill_color=info['fill_color'],
            no_line=is_dot,
            path_items=path_items,
            closePath=closePath,
            shape_rot=info.get('shape_rot', 0),
        )
        root.append(elem)
        shape_id += 1
        count += 1

    # テキスト
    for span in extract_text_spans(page, transform):
        elem = make_shape_xml(
            shape_id=shape_id,
            name=f'text_{shape_id}',
            prst='rect',
            x1=span['x1'], y1=span['y1'],
            x2=span['x2'], y2=span['y2'],
            line_color=span.get('color', '000000'),
            fill_color=None,
            text=span['text'],
            font_size=span['size'],
            no_line=True,
            text_rotation=span.get('rotation', 0),
        )
        root.append(elem)
        shape_id += 1
        count += 1

    xml_bytes = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    xml_bytes += tostring(root, encoding='unicode').encode('utf-8')
    return xml_bytes, count


def convert_pid_to_xlsx(pdf_path: str, xlsx_path: str = None):
    """P&ID PDFをExcelに変換"""
    pdf_path = Path(pdf_path)
    if xlsx_path is None:
        xlsx_path = pdf_path.with_suffix('.xlsx')
    else:
        xlsx_path = Path(xlsx_path)

    print(f"入力: {pdf_path}")
    print(f"出力: {xlsx_path}")

    doc = fitz.open(str(pdf_path))

    # Step 1: openpyxlでベースのxlsxを作成（セルサイズ設定のみ）
    wb = Workbook()
    for page_idx in range(len(doc)):
        page = doc[page_idx]
        if page_idx == 0:
            ws = wb.active
            ws.title = f"P&ID_{page_idx + 1}"
        else:
            ws = wb.create_sheet(title=f"P&ID_{page_idx + 1}")

        # グリッド線を非表示
        ws.sheet_view.showGridLines = False

        # ページ設定（横向き、余白最小、1ページに収める）
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID  # 11x17インチ
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins = PageMargins(left=0.1, right=0.1, top=0.1, bottom=0.1,
                                       header=0, footer=0)

        # 細かいグリッド（PDF 1224x792pt → 96列x106行で十分）
        num_cols = int(page.rect.width / 12.75) + 5
        num_rows = int(page.rect.height / 7.5) + 5
        for col_idx in range(1, num_cols + 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = COL_WIDTH_CHARS
        for row_idx in range(1, num_rows + 1):
            ws.row_dimensions[row_idx].height = ROW_HEIGHT_PT

    wb.save(str(xlsx_path))
    doc_page_count = len(doc)

    # Step 2: Drawing XMLを生成してxlsxに注入
    drawing_data = []
    for page_idx in range(doc_page_count):
        page = doc[page_idx]
        print(f"\nページ {page_idx + 1}/{doc_page_count} "
              f"(サイズ: {page.rect.width:.0f}x{page.rect.height:.0f} pt)")
        xml_bytes, count = build_drawing_xml(page)
        drawing_data.append((xml_bytes, count))
        print(f"  図形数: {count}")

    doc.close()

    # Step 3: xlsxファイルにDrawing XMLを注入
    inject_drawings(str(xlsx_path), drawing_data)

    total = sum(c for _, c in drawing_data)
    print(f"\n変換完了: {xlsx_path}")
    print(f"合計図形数: {total}")


def inject_drawings(xlsx_path: str, drawing_data: list):
    """xlsxファイルにDrawing XMLを直接注入"""
    tmp_path = xlsx_path + '.tmp'

    try:
        _inject_drawings_impl(xlsx_path, tmp_path, drawing_data)
        # 元のファイルを置き換え
        shutil.move(tmp_path, xlsx_path)
    except Exception:
        # 失敗時に一時ファイルを削除
        if Path(tmp_path).exists():
            Path(tmp_path).unlink()
        raise


def _inject_drawings_impl(xlsx_path: str, tmp_path: str, drawing_data: list):
    """inject_drawingsの内部実装"""
    with zipfile.ZipFile(xlsx_path, 'r') as zin, \
         zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:

        # 既存ファイルをコピー（変更が必要なものは除く）
        modified_files = set()
        for page_idx in range(len(drawing_data)):
            modified_files.add(f'xl/worksheets/sheet{page_idx + 1}.xml')
        modified_files.add('[Content_Types].xml')

        rels_path = 'xl/_rels/workbook.xml.rels'

        for item in zin.namelist():
            if item in modified_files:
                continue
            data = zin.read(item)
            zout.writestr(item, data)

        # Drawing XMLを追加
        for page_idx, (xml_bytes, count) in enumerate(drawing_data):
            if count == 0:
                continue

            drawing_path = f'xl/drawings/drawing{page_idx + 1}.xml'
            zout.writestr(drawing_path, xml_bytes)

            # シートのRelationshipsにDrawingを追加
            sheet_rels_path = f'xl/worksheets/_rels/sheet{page_idx + 1}.xml.rels'
            rels_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<Relationships xmlns="{NS_REL}">'
                f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
                f'Target="../drawings/drawing{page_idx + 1}.xml"/>'
                '</Relationships>'
            )
            zout.writestr(sheet_rels_path, rels_xml.encode('utf-8'))

        # シートXMLにdrawing参照を追加
        for page_idx in range(len(drawing_data)):
            sheet_path = f'xl/worksheets/sheet{page_idx + 1}.xml'
            sheet_xml = zin.read(sheet_path).decode('utf-8')

            if drawing_data[page_idx][1] > 0:
                # </worksheet> の前に <drawing r:id="rId1"/> を挿入
                drawing_ref = f'<drawing r:id="rId1" xmlns:r="{NS_R}"/>'
                sheet_xml = sheet_xml.replace('</worksheet>', f'{drawing_ref}</worksheet>')

            zout.writestr(sheet_path, sheet_xml.encode('utf-8'))

        # Content_Types にDrawingのコンテンツタイプを追加
        ct_xml = zin.read('[Content_Types].xml').decode('utf-8')
        for page_idx, (_, count) in enumerate(drawing_data):
            if count > 0:
                override = (
                    f'<Override PartName="/xl/drawings/drawing{page_idx + 1}.xml" '
                    f'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
                )
                ct_xml = ct_xml.replace('</Types>', f'{override}</Types>')
        zout.writestr('[Content_Types].xml', ct_xml.encode('utf-8'))


def main():
    if len(sys.argv) < 2:
        pdf_files = list(Path('.').glob('*.pdf'))
        if not pdf_files:
            print("使用法: python pid2xlsx.py <input.pdf> [output.xlsx]")
            sys.exit(1)
        pdf_path = pdf_files[0]
        xlsx_path = None
    else:
        pdf_path = sys.argv[1]
        xlsx_path = sys.argv[2] if len(sys.argv) > 2 else None

    convert_pid_to_xlsx(pdf_path, xlsx_path)


if __name__ == '__main__':
    main()
