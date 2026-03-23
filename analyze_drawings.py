"""
PDF Drawing Data Analysis Script
Analyzes drawing data from System_test_PID_PII_R1_01.pdf
"""

import fitz
from collections import Counter

PDF_PATH = "System_test_PID_PII_R1_01.pdf"


def item_type_label(items):
    """Return a human-readable label for the item composition."""
    types = [i[0] for i in items]
    counter = Counter(types)
    parts = []
    for t, n in sorted(counter.items()):
        parts.append(f"{n} {t}")
    return ", ".join(parts)


def check_valve_criteria(drawing):
    """Replicate the valve detection logic from pid2xlsx.py."""
    items = drawing['items']
    rect = drawing['rect']
    line_items = [i for i in items if i[0] == 'l']

    if len(line_items) != 3 or len(items) != 3:
        return False, "not 3 lines"

    mw = rect.x1 - rect.x0
    mh = rect.y1 - rect.y0

    if min(mw, mh) <= 5:
        return False, f"min(mw,mh)={min(mw,mh):.2f} <= 5"
    if max(mw, mh) >= 30:
        return False, f"max(mw,mh)={max(mw,mh):.2f} >= 30"
    if max(mw, mh) / max(min(mw, mh), 0.1) >= 2.5:
        return False, f"aspect ratio={max(mw,mh)/max(min(mw,mh),0.1):.2f} >= 2.5"

    pts = set()
    for li in line_items:
        pts.add((round(li[1].x, 1), round(li[1].y, 1)))
        pts.add((round(li[2].x, 1), round(li[2].y, 1)))

    if len(pts) != 4:
        return False, f"unique points={len(pts)}, need 4"

    return True, "PASS"


def main():
    doc = fitz.open(PDF_PATH)
    page = doc[0]

    print("=" * 80)
    print("PDF DRAWING DATA ANALYSIS")
    print("=" * 80)

    # --- Page info ---
    print(f"\n--- Page Info ---")
    print(f"Rotation: {page.rotation}")
    print(f"MediaBox: {page.mediabox}")
    print(f"MediaBox width x height: {page.mediabox.width} x {page.mediabox.height}")
    print(f"Page rect: {page.rect}")
    print(f"Page rect width x height: {page.rect.width} x {page.rect.height}")

    # --- Coordinate transform verification ---
    print(f"\n--- Coordinate Transform Verification (rotation={page.rotation}) ---")
    mbox = page.mediabox
    mw, mh = mbox.width, mbox.height
    print(f"MediaBox: mw={mw}, mh={mh}")

    if page.rotation == 270:
        print("Transform formula: (x, y) -> (y, mw - x)")
        samples = [(0, 0), (100, 0), (0, 100), (100, 200), (mh, mw)]
        for sx, sy in samples:
            tx, ty = sy, mw - sx
            print(f"  ({sx:>7.1f}, {sy:>7.1f}) -> ({tx:>7.1f}, {ty:>7.1f})")
    else:
        print(f"Rotation is {page.rotation}, not 270 - showing identity")

    # --- Get all drawings ---
    drawings = page.get_drawings()
    total = len(drawings)
    print(f"\n--- Total Drawings: {total} ---")

    # --- Composition statistics ---
    composition_counter = Counter()
    type_counter = Counter()  # classified type
    valve_pass = 0
    valve_candidates = []
    rect_drawings = []
    single_lines = []
    mixed_type_drawings = []

    for idx, d in enumerate(drawings):
        items = d['items']
        rect = d['rect']
        color = d.get('color')
        fill = d.get('fill')
        width = d.get('width', 1.0)
        closePath = d.get('closePath', False)

        if not items:
            composition_counter["empty"] += 1
            continue

        label = item_type_label(items)
        composition_counter[label] += 1

        # Check item types present
        item_types = set(i[0] for i in items)

        # Mixed type detection
        if len(item_types) > 1:
            mixed_type_drawings.append((idx, d, label))

        # 3-line valve candidates
        line_items = [i for i in items if i[0] == 'l']
        if len(line_items) == 3 and len(items) == 3:
            passed, reason = check_valve_criteria(d)
            valve_candidates.append((idx, d, passed, reason))
            if passed:
                valve_pass += 1

        # Rectangle drawings
        if any(i[0] in ('re', 'qu') for i in items):
            rect_drawings.append((idx, d, label))

        # Single line items
        if len(items) == 1 and items[0][0] == 'l':
            single_lines.append((idx, d))

    # --- Print composition stats ---
    print(f"\n--- Item Composition Statistics ---")
    for comp, cnt in composition_counter.most_common():
        print(f"  {comp:40s} : {cnt}")

    # --- Valve candidates (3 lines) ---
    print(f"\n{'=' * 80}")
    print(f"VALVE CANDIDATES (exactly 3 line items): {len(valve_candidates)}")
    print(f"{'=' * 80}")
    for idx, d, passed, reason in valve_candidates:
        items = d['items']
        rect = d['rect']
        mw_d = rect.x1 - rect.x0
        mh_d = rect.y1 - rect.y0
        print(f"\n  Drawing #{idx}:")
        print(f"    rect: ({rect.x0:.1f}, {rect.y0:.1f}) - ({rect.x1:.1f}, {rect.y1:.1f})")
        print(f"    rect size: {mw_d:.2f} x {mh_d:.2f}")
        print(f"    color: {d.get('color')}, fill: {d.get('fill')}, width: {d.get('width', 1.0)}")
        print(f"    closePath: {d.get('closePath', False)}")
        print(f"    Valve test: {'PASS' if passed else 'FAIL'} - {reason}")
        # Show line endpoints
        pts = set()
        for li_idx, li in enumerate(items):
            p1, p2 = li[1], li[2]
            pts.add((round(p1.x, 1), round(p1.y, 1)))
            pts.add((round(p2.x, 1), round(p2.y, 1)))
            print(f"    Line {li_idx}: ({p1.x:.1f}, {p1.y:.1f}) -> ({p2.x:.1f}, {p2.y:.1f})")
        print(f"    Unique points (rounded): {len(pts)} -> {sorted(pts)}")

    # --- Rectangle drawings ---
    print(f"\n{'=' * 80}")
    print(f"RECTANGLE DRAWINGS ('re' or 'qu' items): {len(rect_drawings)}")
    print(f"{'=' * 80}")
    for idx, d, label in rect_drawings[:30]:  # limit output
        rect = d['rect']
        print(f"\n  Drawing #{idx}: [{label}]")
        print(f"    rect: ({rect.x0:.1f}, {rect.y0:.1f}) - ({rect.x1:.1f}, {rect.y1:.1f})")
        print(f"    size: {rect.x1-rect.x0:.2f} x {rect.y1-rect.y0:.2f}")
        print(f"    color: {d.get('color')}, fill: {d.get('fill')}, width: {d.get('width', 1.0)}")
        print(f"    closePath: {d.get('closePath', False)}")
        for i, item in enumerate(d['items']):
            if item[0] in ('re', 'qu'):
                print(f"    Item {i}: type={item[0]}, data={item[1:]}")
            else:
                print(f"    Item {i}: type={item[0]}")
    if len(rect_drawings) > 30:
        print(f"\n  ... and {len(rect_drawings) - 30} more rectangle drawings")

    # --- Single lines ---
    print(f"\n{'=' * 80}")
    print(f"SINGLE LINE ITEMS: {len(single_lines)}")
    print(f"{'=' * 80}")
    horiz_count = 0
    vert_count = 0
    angled_count = 0
    angled_lines = []
    for idx, d in single_lines:
        p1, p2 = d['items'][0][1], d['items'][0][2]
        dx = abs(p2.x - p1.x)
        dy = abs(p2.y - p1.y)
        if dx < 1.5 and dy > 1.5:
            vert_count += 1
        elif dy < 1.5 and dx > 1.5:
            horiz_count += 1
        else:
            angled_count += 1
            angled_lines.append((idx, d, dx, dy))

    print(f"  Horizontal (dy < 1.5): {horiz_count}")
    print(f"  Vertical (dx < 1.5): {vert_count}")
    print(f"  Angled (or very short): {angled_count}")

    if angled_lines:
        print(f"\n  Angled/short lines detail (first 30):")
        for idx, d, dx, dy in angled_lines[:30]:
            p1, p2 = d['items'][0][1], d['items'][0][2]
            import math
            length = math.sqrt(dx**2 + dy**2)
            if length > 0.001:
                angle = math.degrees(math.atan2(dy, dx))
            else:
                angle = 0
            print(f"    Drawing #{idx}: ({p1.x:.2f}, {p1.y:.2f}) -> ({p2.x:.2f}, {p2.y:.2f})  "
                  f"dx={dx:.2f} dy={dy:.2f} len={length:.2f} angle={angle:.1f}deg "
                  f"color={d.get('color')} width={d.get('width', 1.0)}")
        if len(angled_lines) > 30:
            print(f"    ... and {len(angled_lines) - 30} more")

    # --- Mixed type drawings ---
    print(f"\n{'=' * 80}")
    print(f"MIXED TYPE DRAWINGS: {len(mixed_type_drawings)}")
    print(f"{'=' * 80}")
    for idx, d, label in mixed_type_drawings[:20]:
        rect = d['rect']
        print(f"\n  Drawing #{idx}: [{label}]")
        print(f"    rect: ({rect.x0:.1f}, {rect.y0:.1f}) - ({rect.x1:.1f}, {rect.y1:.1f})")
        print(f"    color: {d.get('color')}, fill: {d.get('fill')}, width: {d.get('width', 1.0)}")
        print(f"    closePath: {d.get('closePath', False)}")
        for i, item in enumerate(d['items']):
            if item[0] == 'l':
                print(f"    Item {i}: l ({item[1].x:.1f},{item[1].y:.1f})->({item[2].x:.1f},{item[2].y:.1f})")
            elif item[0] == 'c':
                print(f"    Item {i}: c ({item[1].x:.1f},{item[1].y:.1f})->...->({item[4].x:.1f},{item[4].y:.1f})")
            elif item[0] in ('re', 'qu'):
                print(f"    Item {i}: {item[0]} {item[1:]}")
            else:
                print(f"    Item {i}: {item[0]}")
    if len(mixed_type_drawings) > 20:
        print(f"\n  ... and {len(mixed_type_drawings) - 20} more mixed drawings")

    # --- Classification summary ---
    print(f"\n{'=' * 80}")
    print(f"CLASSIFICATION SUMMARY")
    print(f"{'=' * 80}")

    # Run classify_drawing on each to see how pid2xlsx classifies them
    from pid2xlsx import classify_drawing, make_coord_transform
    transform = make_coord_transform(page)

    classified = Counter()
    null_count = 0
    for d in drawings:
        info = classify_drawing(d, transform)
        if info is None:
            null_count += 1
        else:
            classified[info.get('type', 'unknown')] += 1

    print(f"  Total drawings: {total}")
    print(f"  Classified as None (skipped): {null_count}")
    for t, c in classified.most_common():
        print(f"  {t:20s}: {c}")
    print(f"  Valve detection passes: {valve_pass}")
    print(f"  Rectangle drawings (re/qu items): {len(rect_drawings)}")

    doc.close()
    print(f"\n{'=' * 80}")
    print("ANALYSIS COMPLETE")
    print(f"{'=' * 80}")


if __name__ == '__main__':
    main()
