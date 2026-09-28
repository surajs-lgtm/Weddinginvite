#!/usr/bin/env python3
"""
Generate the decorative artwork for the wedding site as SVG.

Everything here is drawn from scratch as vector art: the Maru-Gurjara mandala,
the Hawa Mahal jali lattice, toran garlands, and one line icon per ritual.

Run:  python3 make_art.py
Writes into docs/assets/ (decoration) and docs/assets/icons/ (ritual icons).
"""
import math
import os

try:
    from check_art import fit_viewbox
except ImportError:  # run from another directory
    import importlib.util
    _spec = importlib.util.spec_from_file_location(
        "check_art", os.path.join(os.path.dirname(os.path.abspath(__file__)), "check_art.py"))
    _ca = importlib.util.module_from_spec(_spec)
    _spec.loader.exec_module(_ca)
    fit_viewbox = _ca.fit_viewbox

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "docs", "assets")
ICONS = os.path.join(ASSETS, "icons")

GOLD = "#c9a227"
GOLD_LIGHT = "#e6c766"
GOLD_PALE = "#f0dca4"
MAROON = "#7b1e2b"
DEEP = "#3d0c11"
MARIGOLD = "#e8890c"
MARIGOLD_LT = "#f7b93b"
ALTA = "#a4243b"
LEAF = "#2e6b3e"
LEAF_LT = "#4a9a5c"
CLAY = "#b5651d"
CREAM = "#fdf6e9"


# ----------------------------------------------------------------- helpers
def polar(cx, cy, r, deg):
    a = math.radians(deg)
    return cx + r * math.cos(a), cy + r * math.sin(a)


def petal(cx, cy, ri, ro, angle, width, curve=0.55):
    """A pointed leaf/lotus petal from radius ri to ro, centred on `angle`."""
    tip_ctrl = polar(cx, cy, ro * curve, angle)
    x0, y0 = polar(cx, cy, ri, angle - width)
    x1, y1 = polar(cx, cy, ro, angle)
    x2, y2 = polar(cx, cy, ri, angle + width)
    return (f"M{x0:.2f} {y0:.2f}Q{tip_ctrl[0]:.2f} {tip_ctrl[1]:.2f} {x1:.2f} {y1:.2f}"
            f"Q{tip_ctrl[0]:.2f} {tip_ctrl[1]:.2f} {x2:.2f} {y2:.2f}Z")


def dots(cx, cy, count, radius, r=1.6):
    out = []
    for i in range(count):
        x, y = polar(cx, cy, radius, i * (360 / count))
        out.append(f'<circle cx="{x:.2f}" cy="{y:.2f}" r="{r}"/>')
    return "".join(out)


def svg(w, h, body, title=""):
    t = f'<title>{title}</title>' if title else ""
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}" '
            f'width="{w}" height="{h}" role="img">{t}{body}</svg>\n')


def write(path, content):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(content)
    print(f"  {os.path.relpath(path, HERE)}  ({len(content)} bytes)")


# ================================================================= mandala
def mandala():
    """A dense Maru-Gurjara rosette, the style carved into Marwar temples."""
    C = 300.0
    b = []

    # outer halo
    b.append(f'<circle cx="{C}" cy="{C}" r="286" fill="none" stroke="{GOLD}" stroke-width="2" opacity=".55"/>')
    b.append(f'<circle cx="{C}" cy="{C}" r="276" fill="none" stroke="{GOLD}" stroke-width="1" opacity=".4"/>')
    b.append("".join(f'<circle cx="{x:.2f}" cy="{y:.2f}" r="3" fill="{GOLD}" opacity=".5"/>'
                      for x, y in (polar(C, C, 286, i * 7.5) for i in range(48))))

    # outer ring of lotus petals
    b.append('<g fill="none" stroke="' + GOLD + '" stroke-width="1.8" opacity=".75">')
    for i in range(24):
        b.append(f'<path d="{petal(C, C, 208, 268, i * 15, 5.2)}"/>')
    b.append("</g>")

    # ring of pellets
    b.append(f'<g fill="{GOLD}" opacity=".4">{dots(C, C, 48, 200, 2.2)}</g>')

    # middle ring, alternating long and short petals
    b.append('<g fill="none" stroke="' + GOLD_LIGHT + '" stroke-width="1.6" opacity=".8">')
    for i in range(16):
        b.append(f'<path d="{petal(C, C, 128, 196, i * 22.5 + 11, 7.5, 0.5)}"/>')
    b.append("</g>")
    b.append('<g fill="' + GOLD + '" opacity=".22">')
    for i in range(16):
        b.append(f'<path d="{petal(C, C, 128, 190, i * 22.5 + 11, 3.6, 0.5)}"/>')
    b.append("</g>")

    b.append(f'<circle cx="{C}" cy="{C}" r="124" fill="none" stroke="{GOLD}" stroke-width="1.4" opacity=".6"/>')
    b.append(f'<g fill="{GOLD}" opacity=".5">{dots(C, C, 32, 112, 1.8)}</g>')

    # inner lotus
    b.append('<g fill="none" stroke="' + GOLD_LIGHT + '" stroke-width="1.5" opacity=".85">')
    for i in range(12):
        b.append(f'<path d="{petal(C, C, 58, 108, i * 30, 9.5, 0.48)}"/>')
    b.append("</g>")

    # interlaced octagons at the centre
    for rot, op in ((0, ".8"), (22.5, ".55")):
        pts = " ".join(f"{x:.2f},{y:.2f}" for x, y in
                       (polar(C, C, 54, i * 45 + rot) for i in range(8)))
        b.append(f'<polygon points="{pts}" fill="none" stroke="{GOLD}" stroke-width="1.6" opacity="{op}"/>')

    b.append(f'<circle cx="{C}" cy="{C}" r="30" fill="{MAROON}" opacity=".14"/>')
    b.append(f'<circle cx="{C}" cy="{C}" r="30" fill="none" stroke="{GOLD}" stroke-width="2" opacity=".85"/>')
    b.append(f'<circle cx="{C}" cy="{C}" r="17" fill="none" stroke="{GOLD_LIGHT}" stroke-width="1.4" opacity=".7"/>')
    b.append(f'<circle cx="{C}" cy="{C}" r="6" fill="{GOLD}" opacity=".8"/>')

    # radiating spokes out to the rim
    b.append('<g stroke="' + GOLD + '" stroke-width="1" opacity=".28">')
    for i in range(24):
        x0, y0 = polar(C, C, 214, i * 15 + 7.5)
        x1, y1 = polar(C, C, 272, i * 15 + 7.5)
        b.append(f'<line x1="{x0:.2f}" y1="{y0:.2f}" x2="{x1:.2f}" y2="{y1:.2f}"/>')
    b.append("</g>")

    return svg(600, 600, "".join(b), "Mandala")


# ==================================================================== jali
def jali():
    """A tileable Hawa Mahal lattice: octagrams linked by a diagonal grid."""
    T = 120
    b = [f'<g fill="none" stroke="{GOLD}" stroke-width="1.1" opacity=".5">']

    def octagram(cx, cy, r, rot=0):
        a = " ".join(f"{x:.2f},{y:.2f}" for x, y in
                     (polar(cx, cy, r, i * 45 + rot) for i in range(8)))
        b2 = " ".join(f"{x:.2f},{y:.2f}" for x, y in
                      (polar(cx, cy, r, i * 45 + rot + 22.5) for i in range(8)))
        return (f'<polygon points="{a}"/><polygon points="{b2}"/>'
                f'<circle cx="{cx}" cy="{cy}" r="{r * 0.52:.2f}"/>')

    # stars must overhang the edges so the tile repeats seamlessly
    for cx in (30, 90):
        for cy in (30, 90):
            b.append(octagram(cx, cy, 26))
    for cx in (0, 120):
        for cy in (0, 120):
            b.append(octagram(cx, cy, 26))
    for cx in (0, 120):
        for cy in (60,):
            b.append(octagram(cx, cy, 22))
    for cy in (0, 120):
        for cx in (60,):
            b.append(octagram(cx, cy, 22))

    # diamond grid tying the stars together
    b.append('<g opacity=".32">')
    for i in range(-1, 4):
        b.append(f'<line x1="{i * 60}" y1="0" x2="{i * 60 + 120}" y2="120"/>')
        b.append(f'<line x1="{i * 60}" y1="120" x2="{i * 60 + 120}" y2="0"/>')
    b.append("</g>")
    b.append("</g>")

    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {T} {T}" width="{T}" '
            f'height="{T}">{"".join(b)}</svg>\n')


# =================================================================== toran
def marigold(cx, cy, r, petal_color=MARIGOLD, core=MARIGOLD_LT):
    b = [f'<circle cx="{cx}" cy="{cy}" r="{r * 2.05:.2f}" fill="none" '
         f'stroke="{petal_color}" stroke-width="{r * 0.5:.2f}" opacity=".5"/>']
    for i in range(10):
        x, y = polar(cx, cy, r * 1.5, i * 36)
        b.append(f'<circle cx="{x:.2f}" cy="{y:.2f}" r="{r * 0.78:.2f}" fill="{petal_color}" opacity=".9"/>')
    b.append(f'<circle cx="{cx}" cy="{cy}" r="{r * 0.95:.2f}" fill="{core}"/>')
    b.append(f'<circle cx="{cx}" cy="{cy}" r="{r * 0.5:.2f}" fill="{MARIGOLD_LT}" opacity=".65"/>')
    return "".join(b)


def mango_leaves(cx, cy, scale=1.0, flip=False):
    b = []
    s = -1 if flip else 1
    for i, ang in enumerate((-62, -34, 0, 34, 62)):
        tipx = cx + s * math.sin(math.radians(ang)) * 34 * scale
        tipy = cy - math.cos(math.radians(ang)) * 34 * scale
        col = LEAF if i % 2 == 0 else LEAF_LT
        b.append(f'<path d="M{cx:.1f} {cy:.1f}Q{cx + s * 12 * scale:.1f} {cy - 20 * scale:.1f} '
                 f'{tipx:.1f} {tipy:.1f}" fill="none" stroke="{col}" stroke-width="2" stroke-linecap="round"/>')
        for k in (0.35, 0.62, 0.85):
            px = cx + (tipx - cx) * k
            py = cy + (tipy - cy) * k - 5 * scale
            b.append(f'<ellipse cx="{px:.1f}" cy="{py:.1f}" rx="{9 * scale:.1f}" ry="{4.2 * scale:.1f}" '
                     f'fill="{col}" transform="rotate({-s * ang * 0.55:.1f} {px:.1f} {py:.1f})"/>')
    return "".join(b)


def toran():
    """A hanging marigold garland with mango leaves, over a red valance."""
    W, H = 520, 150
    b = [f'<path d="M0 6 Q130 46 260 8 Q390 46 520 6" fill="none" '
         f'stroke="{ALTA}" stroke-width="9" stroke-linecap="round" opacity=".85"/>']
    b.append(f'<path d="M0 16 Q130 56 260 18 Q390 56 520 16" fill="none" '
             f'stroke="{GOLD}" stroke-width="2" opacity=".5"/>')

    # beaded string
    b.append(f'<path d="M0 8 Q130 48 260 10 Q390 48 520 8" fill="none" '
             f'stroke="{GOLD_LIGHT}" stroke-width="1" opacity=".55" stroke-dasharray="2 6"/>')

    for i in range(13):
        t = i / 12
        # sample both sagging curves
        if t <= 0.5:
            u = t * 2
            x = 0 + 260 * u
            y = 8 + (10 - 8 + (48 - 8) * 2 * u * (1 - u) * 2) * u
        else:
            u = (t - 0.5) * 2
            x = 260 + 260 * u
            y = 10 + (8 - 10 + (48 - 8) * 2 * u * (1 - u) * 2) * u
        b.append(marigold(x, y + 16, 7))

    for x, flip in ((72, False), (448, True)):
        b.append(f'<line x1="{x}" y1="26" x2="{x}" y2="52" stroke="{LEAF}" stroke-width="1.6" opacity=".7"/>')
        b.append(mango_leaves(x, 52, 0.85, flip))

    # central pendant
    b.append(f'<line x1="260" y1="12" x2="260" y2="96" stroke="{GOLD}" stroke-width="1.4" opacity=".65"/>')
    for i in range(5):
        y = 30 + i * 15
        b.append(marigold(260, y, 9 - i * 0.6, ALTA if i % 2 else MARIGOLD, MARIGOLD_LT))
    b.append(mango_leaves(260, 108, 1.0, False))

    return svg(W, H, "".join(b), "Toran garland")


def garland_vertical():
    W, H = 90, 420
    b = [f'<path d="M45 0 Q18 210 45 420" fill="none" stroke="{GOLD}" stroke-width="1" opacity=".4"/>']
    for i in range(9):
        y = 16 + i * 47
        x = 45 + (18 if i % 2 else -18) * 0.5
        b.append(f'<line x1="{x:.0f}" y1="{y - 16}" x2="{x:.0f}" y2="{y - 6}" stroke="{LEAF}" stroke-width="1.4" opacity=".55"/>')
        b.append(mango_leaves(x, y, 0.5, i % 2 == 0))
        b.append(marigold(x, y + 22, 8))
    return svg(W, H, "".join(b), "Marigold garland")


# ================================================================ kalash
def kalash():
    W, H = 170, 250
    cx = 85
    b = []
    # mango-leaf crown behind the coconut
    b.append(mango_leaves(cx, 54, 1.55, False))
    # coconut
    b.append(f'<circle cx="{cx}" cy="40" r="20" fill="{CLAY}" opacity=".9"/>')
    b.append(f'<circle cx="{cx}" cy="40" r="20" fill="none" stroke="{DEEP}" stroke-width="1.2" opacity=".4"/>')
    b.append(f'<path d="M{cx - 13} 46q7 -11 13 -5q6 -6 13 5" fill="none" stroke="{DEEP}" stroke-width="1.2" opacity=".45"/>')

    # pot
    b.append(f'<path d="M{cx - 26} 96h52c0 12 -3 18 -8 24c-9 12 -12 24 -12 40 '
             f'c0 30 10 48 30 56c-14 10 -46 10 -60 0c20 -8 30 -26 30 -56'
             f'c0 -16 -3 -28 -12 -40c-5 -6 -8 -12 -8 -24z" fill="{ALTA}" opacity=".92"/>')
    b.append(f'<path d="M{cx - 26} 96h52" stroke="{GOLD}" stroke-width="4" stroke-linecap="round"/>')
    b.append(f'<path d="M{cx - 30} 92h60" stroke="{GOLD}" stroke-width="3" stroke-linecap="round" opacity=".85"/>')
    # belly band with grain motif
    b.append(f'<rect x="{cx - 34}" y="182" width="68" height="26" rx="4" fill="{GOLD}" opacity=".9"/>')
    for i in range(9):
        x = cx - 30 + i * 7.5
        b.append(f'<ellipse cx="{x:.1f}" cy="195" rx="2.4" ry="4" fill="{DEEP}" opacity=".55"/>')
    # rice stream
    b.append('<g fill="' + GOLD_PALE + '">')
    for i in range(11):
        x, y = polar(cx + 2, 120, 26, i * 33)
        b.append(f'<ellipse cx="{x:.1f}" cy="{y:.1f}" rx="2.3" ry="3.6" transform="rotate({i * 27 % 70 - 35:.0f} {x:.1f} {y:.1f})" opacity=".75"/>')
    b.append("</g>")
    b.append(f'<path d="M{cx - 30} 230q30 14 60 0" fill="none" stroke="{GOLD}" stroke-width="2" opacity=".6"/>')
    return svg(W, H, "".join(b), "Kalash")


# ================================================================== diyas
def diya_row():
    W, H = 480, 90
    b = [f'<path d="M0 66h480" stroke="{GOLD}" stroke-width="1" opacity=".35"/>']

    def diya(x, scale=1.0, on=True):
        s = scale
        parts = []
        parts.append(f'<path d="M{x - 26 * s:.1f} {66:.1f} q{26 * s:.1f} {16 * s:.1f} {52 * s:.1f} 0 z" '
                     f'fill="{CLAY}" opacity=".85"/>')
        parts.append(f'<path d="M{x - 30 * s:.1f} {66:.1f} h{60 * s:.1f}" stroke="{GOLD}" '
                     f'stroke-width="{2.4 * s:.1f}" stroke-linecap="round" opacity=".9"/>')
        parts.append(f'<rect x="{x - 1.6 * s:.1f}" y="{66 - 30 * s:.1f}" width="{3.2 * s:.1f}" '
                     f'height="{28 * s:.1f}" rx="{1.5 * s:.1f}" fill="{GOLD_PALE}"/>')
        if on:
            fy = 66 - 30 * s
            parts.append(f'<path d="M{x:.1f} {fy:.1f} c{-8 * s:.1f} {-10 * s:.1f} {-6 * s:.1f} {-20 * s:.1f} '
                         f'{0:.1f} {-28 * s:.1f} c{6 * s:.1f} {8 * s:.1f} {8 * s:.1f} {18 * s:.1f} {0:.1f} {28 * s:.1f} z" '
                         f'fill="{MARIGOLD_LT}"/>')
            parts.append(f'<path d="M{x:.1f} {fy:.1f} c{-4 * s:.1f} {-5 * s:.1f} {-3 * s:.1f} {-11 * s:.1f} '
                         f'{0:.1f} {-16 * s:.1f} c{3 * s:.1f} {5 * s:.1f} {3 * s:.1f} {11 * s:.1f} {0:.1f} {16 * s:.1f} z" '
                         f'fill="{CREAM}" opacity=".9"/>')
        return "".join(parts)

    for x, sc, on in ((60, 0.72, False), (150, 0.92, True), (240, 1.08, True),
                      (330, 0.92, True), (420, 0.72, False)):
        b.append(diya(x, sc, on))
    return svg(W, H, "".join(b), "Row of diyas")


# ================================================================ mandap
def mandap():
    """A four-pillar wedding mandap with a scalloped canopy."""
    W, H = 420, 230
    b = []
    # canopy
    b.append(f'<path d="M40 66 Q210 6 380 66 Q210 44 40 66Z" fill="{ALTA}" opacity=".9"/>')
    b.append(f'<path d="M40 66 Q210 6 380 66" fill="none" stroke="{GOLD}" stroke-width="3"/>')
    for i in range(1, 8):
        x = 40 + i * 42.5
        y = 66 - 44 * math.sin(math.pi * i / 8) + 6
        b.append(f'<path d="M{x:.0f} 66 q{-10} 10 0 20" fill="none" stroke="{GOLD}" '
                 f'stroke-width="1.4" opacity=".55"/>')
    b.append(marigold(40, 62, 6))
    b.append(marigold(380, 62, 6))
    # kalash finials
    b.append(f'<path d="M30 66l10 -20l10 20z" fill="{GOLD}" opacity=".9"/>')
    b.append(f'<path d="M370 66l10 -20l10 20z" fill="{GOLD}" opacity=".9"/>')

    # pillars
    for x in (72, 348):
        b.append(f'<rect x="{x - 7}" y="80" width="14" height="128" rx="4" fill="{GOLD}" opacity=".85"/>')
        for i in range(5):
            yy = 96 + i * 26
            b.append(f'<circle cx="{x}" cy="{yy}" r="4.5" fill="{DEEP}" opacity=".4"/>')
        b.append(f'<rect x="{x - 16}" y="72" width="32" height="12" rx="3" fill="{GOLD}" opacity=".95"/>')
        b.append(f'<rect x="{x - 16}" y="204" width="32" height="14" rx="3" fill="{GOLD}" opacity=".95"/>')
    b.append(f'<rect x="56" y="216" width="308" height="9" rx="4" fill="{GOLD}" opacity=".7"/>')

    # back cloth
    b.append(f'<path d="M86 84h248v112H86z" fill="{MAROON}" opacity=".16"/>')
    b.append(marigold(210, 150, 13, ALTA, MARIGOLD_LT))
    b.append(mango_leaves(210, 186, 0.95, False))
    return svg(W, H, "".join(b), "Wedding mandap")


# ================================================================= divider
def divider():
    W, H = 300, 46
    cx, cy = 150, 23
    b = [f'<g stroke="{GOLD}" stroke-width="1.3" opacity=".65" fill="none">']
    b.append(f'<path d="M8 23h96M196 23h96"/>')
    b.append(f'<path d="M14 18v10M22 18v10M30 18v10" opacity=".55"/>')
    b.append(f'<path d="M262 18v10M270 18v10M278 18v10" opacity=".55"/>')
    b.append("</g>")
    # central lotus
    b.append(f'<g fill="none" stroke="{GOLD}" stroke-width="1.5" opacity=".8">')
    for i in range(8):
        b.append(f'<path d="{petal(cx, cy, 4, 16, i * 45, 13, 0.5)}"/>')
    b.append("</g>")
    b.append(f'<circle cx="{cx}" cy="{cy}" r="3.4" fill="{GOLD}" opacity=".85"/>')
    for x in (cx - 30, cx + 30):
        b.append(f'<circle cx="{x}" cy="{cy}" r="2.2" fill="{GOLD}" opacity=".5"/>')
    return svg(W, H, "".join(b), "Divider")


# ================================================================= peacock
def peacock():
    """Mor: the peacock is the Rajput emblem, and the bird of Krishna."""
    W, H = 300, 330
    b = []
    cx, cy = 150, 168
    # fan of train feathers
    for i in range(11):
        a = -180 + i * 16
        tipx, tipy = polar(cx, cy, 132, a)
        midx, midy = polar(cx, cy, 92, a)
        col = LEAF if i % 2 == 0 else "#1f5230"
        b.append(f'<path d="M{cx} {cy}Q{midx - 6 * math.sin(math.radians(a)):.1f} {midy:.1f} {tipx:.1f} {tipy:.1f}" '
                 f'fill="none" stroke="{col}" stroke-width="2.4" stroke-linecap="round" opacity=".85"/>')
        # eye spot
        b.append(f'<ellipse cx="{tipx:.1f}" cy="{tipy:.1f}" rx="13" ry="10" fill="{col}" opacity=".55"/>')
        b.append(f'<ellipse cx="{tipx:.1f}" cy="{tipy:.1f}" rx="9" ry="7" fill="{TEAL}" opacity=".85"/>')
        b.append(f'<ellipse cx="{tipx:.1f}" cy="{tipy:.1f}" rx="4.5" ry="3.4" fill="{GOLD_LIGHT}"/>')
    # body
    b.append(f'<path d="M{cx - 26} {cy + 46} c-10 -26 4 -50 26 -54 c22 -4 36 12 34 34 '
             f'c-2 24 -20 40 -40 40 c-8 0 -16 -8 -20 -20z" fill="{LEAF}"/>')
    b.append(f'<path d="M{cx - 20} {cy + 8} c14 6 26 2 34 -10" fill="none" stroke="{TEAL}" stroke-width="3" opacity=".7"/>')
    # neck and head
    b.append(f'<path d="M{cx + 8} {cy - 6} c14 -6 22 -18 24 -32" fill="none" stroke="{LEAF}" '
             f'stroke-width="9" stroke-linecap="round"/>')
    b.append(f'<circle cx="{cx + 34}" cy="{cy - 42}" r="10" fill="{LEAF}"/>')
    b.append(f'<circle cx="{cx + 38}" cy="{cy - 44}" r="2" fill="{CREAM}"/>')
    # crest
    for i, off in enumerate((-7, 0, 7)):
        b.append(f'<line x1="{cx + 32 + i * 2}" y1="{cy - 51}" x2="{cx + 28 + off * 1.3:.0f}" '
                 f'y2="{cy - 74}" stroke="{TEAL}" stroke-width="2" stroke-linecap="round"/>')
        b.append(f'<circle cx="{cx + 28 + off * 1.3:.0f}" cy="{cy - 76}" r="3" fill="{TEAL}"/>')
    # feet
    for dx in (-10, 10):
        b.append(f'<line x1="{cx + dx}" y1="{cy + 62}" x2="{cx + dx}" y2="{cy + 84}" stroke="{GOLD}" stroke-width="3"/>')
    b.append(f'<path d="M{cx - 34} {cy + 86}h68" stroke="{GOLD}" stroke-width="3" stroke-linecap="round" opacity=".7"/>')
    return svg(W, H, "".join(b), "Peacock")


TEAL = "#1f7a8c"

# ================================================================= corners
def corner():
    W = 120
    b = [f'<g fill="none" stroke="{GOLD}" stroke-width="1.4" opacity=".7">']
    b.append('<path d="M2 40V18a16 16 0 0 1 16-16h22"/>')
    b.append('<path d="M2 62V30a28 28 0 0 1 28-28h30" opacity=".5"/>')
    b.append('<path d="M14 30q18 0 18-18" opacity=".45"/>')
    b.append("</g>")
    b.append(f'<path d="{petal(30, 30, 3, 20, -45, 12, 0.5)}" fill="none" stroke="{GOLD_LIGHT}" stroke-width="1.2" opacity=".6"/>')
    b.append(f'<circle cx="30" cy="30" r="3" fill="{GOLD}" opacity=".7"/>')
    return svg(W, W, "".join(b), "Corner")


def hawa_arch():
    """A scalloped faceted arch, the Hawa Mahal silhouette."""
    W, H = 200, 240
    b = []
    outer = "M14 236V96C14 50 56 12 100 12s86 38 86 84v140"
    b.append(f'<path d="{outer}" fill="none" stroke="{GOLD}" stroke-width="3" opacity=".85"/>')
    b.append(f'<path d="M30 236V100c0 -38 34 -70 70 -70s70 32 70 70v136" fill="none" '
             f'stroke="{GOLD_LIGHT}" stroke-width="1.4" opacity=".5"/>')
    # facet scallops
    for i in range(9):
        x = 30 + i * 17.5
        h = 44 - abs(i - 4) * 8
        y = 96 - i * 3
        b.append(f'<path d="M{x:.0f} {y + h * 0.5:.0f}q{8.75:.1f} {-h * 0.55:.1f} {17.5:.1f} 0" '
                 f'fill="none" stroke="{GOLD}" stroke-width="1.1" opacity=".4"/>')
    b.append(marigold(100, 62, 9))
    b.append(f'<circle cx="100" cy="18" r="6" fill="none" stroke="{GOLD}" stroke-width="2" opacity=".8"/>')
    b.append(f'<path d="M100 8v-6" stroke="{GOLD}" stroke-width="2" stroke-linecap="round"/>')
    return svg(W, H, "".join(b), "Hawa Mahal arch")


# ============================================================ ritual icons
S = ('fill="none" stroke="currentColor" stroke-width="1.8" '
     'stroke-linecap="round" stroke-linejoin="round"')
SF = 'fill="currentColor"'


def icon(body):
    """A 64x64 line icon whose viewBox is auto-fitted, so every icon in the set
    reads at the same visual weight instead of each using a different slice."""
    vb = fit_viewbox(body, pad=3.0)
    view = f'viewBox="{vb[0]:.2f} {vb[1]:.2f} {vb[2]:.2f} {vb[3]:.2f}"' if vb else 'viewBox="0 0 64 64"'
    return (f'<svg {view} xmlns="http://www.w3.org/2000/svg" '
            f'aria-hidden="true" focusable="false" {S}>{body}</svg>\n')


def build_icons():
    ic = {}

    # Tilak - a kumkum mark on a forehead, with rice grains
    ic["tilak"] = icon(
        f'<path d="M32 12c-11 0-19 9-19 20v14h38V32c0-11-8-20-19-20z" opacity=".45"/>'
        f'<path d="M32 20c-6 4-8 9-8 14s3 8 8 8 8-3 8-8-2-10-8-14z" {SF} opacity=".85"/>'
        f'<ellipse cx="15" cy="52" rx="2.4" ry="3.6" transform="rotate(-25 15 52)"/>'
        f'<ellipse cx="24" cy="56" rx="2.4" ry="3.6" transform="rotate(12 24 56)"/>'
        f'<ellipse cx="43" cy="56" rx="2.4" ry="3.6" transform="rotate(-12 43 56)"/>'
        f'<ellipse cx="52" cy="51" rx="2.4" ry="3.6" transform="rotate(25 52 51)"/>')

    # Matkor - a woven sieve with a handle
    ic["matkor"] = icon(
        f'<ellipse cx="32" cy="28" rx="22" ry="9"/>'
        f'<ellipse cx="32" cy="26" rx="22" ry="9" opacity=".5"/>'
        f'<path d="M10 26v4c0 5 10 9 22 9s22-4 22-9v-4" opacity=".7"/>'
        f'<path d="M32 35v16"/>'
        f'<path d="M28 51h8" />'
        f'<path d="M18 22l24 8M22 18l20 9M14 28l28 4" opacity=".4"/>')

    # Madwa - the temporary bamboo canopy raised for the wedding, pitched roof
    # on four poles with mango and banana leaves hung along the eaves
    ic["madwa"] = icon(
        f'<path d="M5 31 14 16h36l9 15"/>'
        f'<path d="M14 16h36" opacity=".45"/>'
        f'<path d="M9.5 23.5 16 18.5M54.5 23.5 48 18.5" opacity=".3"/>'
        f'<path d="M5 31h54"/>'
        f'<path d="M14 31v23M50 31v23"/>'
        f'<path d="M11 54h6M47 54h6" opacity=".5"/>'
        f'<path d="M14 43h36" opacity=".35"/>'
        f'<ellipse cx="18" cy="36" rx="2" ry="3.6" transform="rotate(-22 18 36)" opacity=".5"/>'
        f'<ellipse cx="25" cy="35" rx="2" ry="3.4" opacity=".45"/>'
        f'<ellipse cx="32" cy="36.5" rx="2" ry="3.8" transform="rotate(8 32 36.5)" opacity=".5"/>'
        f'<ellipse cx="39" cy="35" rx="2" ry="3.4" opacity=".45"/>'
        f'<ellipse cx="46" cy="36" rx="2" ry="3.6" transform="rotate(22 46 36)" opacity=".5"/>'
        f'<path d="M8 58h48" opacity=".4"/>')

    # Haldi - a bowl of turmeric paste
    ic["haldi"] = icon(
        f'<path d="M12 34h40c0 11-9 18-20 18s-20-7-20-18z" {SF} opacity=".18"/>'
        f'<path d="M12 34h40" />'
        f'<path d="M18 34c2-7 8-10 14-10s12 3 14 10" {SF} opacity=".55"/>'
        f'<path d="M46 26c6-2 9-6 8-10-4 2-8 4-10 8" />'
        f'<path d="M10 54h44" opacity=".5"/>')

    # Kalra / Devpuji - a small five-wick lamp
    ic["kalra"] = icon(
        f'<path d="M14 42h36c0 8-8 12-18 12s-18-4-18-12z" {SF} opacity=".2"/>'
        f'<path d="M14 42h36"/>'
        f'<path d="M18 46c8 4 20 4 28 0" opacity=".45"/>'
        f'<path d="M22 42c0-5 2-7 2-7s2 2 2 7c0 2-1 3-2 3s-2-1-2-3z" {SF}/>'
        f'<path d="M42 42c0-5 2-7 2-7s2 2 2 7c0 2-1 3-2 3s-2-1-2-3z" {SF}/>'
        f'<path d="M32 42c0-6 2-9 2-9s2 3 2 9c0 2-1 3-2 3s-2-1-2-3z" {SF} opacity=".85"/>'
        f'<path d="M10 58h44" opacity=".45"/>')

    # Dhid Hari - ghee poured over the sacred fire
    ic["dhidhari"] = icon(
        f'<path d="M16 40c0-8 6-13 16-13s16 5 16 13c0 9-7 14-16 14s-16-5-16-14z" {SF} opacity=".2"/>'
        f'<path d="M18 40h28"/>'
        f'<path d="M32 27c-4 5-6 8-6 11a6 6 0 0 0 12 0c0-3-2-6-6-11z" {SF} opacity=".9"/>'
        f'<path d="M46 12c6 4 8 9 6 14" />'
        f'<path d="M52 8c-2 3-3 5-3 7" opacity=".5"/>'
        f'<path d="M22 50h20" opacity=".4"/>')

    # Janau - the sacred fire in its kund
    ic["janau"] = icon(
        f'<path d="M12 50h40v-6H12z" {SF} opacity=".2"/>'
        f'<path d="M12 44h40"/>'
        f'<path d="M16 38h32v-4H16z" opacity=".5"/>'
        f'<path d="M32 10c-7 9-12 14-12 21a12 12 0 0 0 24 0c0-7-5-12-12-21z" {SF} opacity=".9"/>'
        f'<path d="M32 30c-3 5-5 8-5 10a5 5 0 0 0 10 0c0-2-2-5-5-10z" fill="#fdf6e9" stroke="none"/>'
        f'<path d="M8 56h48" opacity=".5"/>')

    # Bhunga Lawa - an earthen pot with rice
    ic["bhungalawa"] = icon(
        f'<path d="M18 28h28c0 4-1 6-3 8-4 5-5 10-5 16 0 6 2 10 6 12H20c4-2 6-6 6-12 0-6-1-11-5-16-2-2-3-4-3-8z" {SF} opacity=".2"/>'
        f'<path d="M18 28h28"/>'
        f'<path d="M24 24c2-4 5-6 8-6s6 2 8 6" />'
        f'<path d="M22 20c4 2 16 2 20 0" opacity=".4"/>'
        f'<ellipse cx="32" cy="14" rx="9" ry="4" opacity=".35"/>'
        f'<path d="M20 46h24" opacity=".45"/>')

    # Parat - a gold tray with a mound of rice
    ic["parat"] = icon(
        f'<ellipse cx="32" cy="42" rx="24" ry="7" {SF} opacity=".25"/>'
        f'<ellipse cx="32" cy="40" rx="24" ry="7"/>'
        f'<path d="M8 40v3c0 4 11 7 24 7s24-3 24-7v-3"/>'
        f'<path d="M18 40c3-9 8-13 14-13s11 4 14 13" {SF} opacity=".55"/>'
        f'<path d="M26 34c0-3 2-5 6-5s6 2 6 5" opacity=".5"/>'
        f'<path d="M20 52h24" opacity=".4"/>')

    # Darwagar - the decorated threshold
    ic["darwagar"] = icon(
        f'<path d="M12 56V28c0-9 9-16 20-16s20 7 20 16v28" />'
        f'<path d="M20 56V30c0-6 5-10 12-10s12 4 12 10v26" opacity=".5"/>'
        f'<path d="M14 26h36" />'
        f'<circle cx="32" cy="30" r="4" {SF} opacity=".7"/>'
        f'<circle cx="20" cy="30" r="3" {SF} opacity=".55"/>'
        f'<circle cx="44" cy="30" r="3" {SF} opacity=".55"/>'
        f'<path d="M20 34v6M32 36v8M44 34v6" opacity=".6"/>'
        f'<path d="M26 44l6 5 6-5" opacity=".45"/>')

    # Vidai - the farewell, a conch and a parting of hands
    ic["vidai"] = icon(
        f'<path d="M12 44c0-11 9-20 20-20s20 9 20 20c0 6-5 10-11 10H23c-6 0-11-4-11-10z" {SF} opacity=".18"/>'
        f'<path d="M12 44c0-11 9-20 20-20s20 9 20 20"/>'
        f'<path d="M18 40c0-6 5-11 11-11" opacity=".55"/>'
        f'<path d="M32 34c-4-3-5-7-2-10 4 3 5 7 2 10z" {SF} opacity=".85"/>'
        f'<path d="M30 24c-3-1-6-4-6-8 4 0 7 3 6 8z" {SF} opacity=".6"/>'
        f'<path d="M14 52h36" opacity=".5"/>'
        f'<ellipse cx="42" cy="48" rx="2.2" ry="3.2" transform="rotate(20 42 48)" opacity=".6"/>'
        f'<ellipse cx="48" cy="43" rx="2.2" ry="3.2" transform="rotate(-15 48 43)" opacity=".6"/>')
    return ic


def favicon():
    """A small gold mandala on maroon, for the browser tab."""
    b = [f'<rect width="64" height="64" rx="12" fill="{MAROON}"/>']
    b.append(f'<g fill="none" stroke="{GOLD_LIGHT}" stroke-width="1.6" opacity=".95">')
    for i in range(8):
        b.append(f'<path d="{petal(32, 32, 3, 13, i * 45, 13, 0.5)}"/>')
    b.append("</g>")
    b.append(f'<circle cx="32" cy="32" r="22" fill="none" stroke="{GOLD}" stroke-width="1.2" opacity=".6"/>')
    b.append(f'<circle cx="32" cy="32" r="3" fill="{GOLD_LIGHT}"/>')
    return svg(64, 64, "".join(b), "Favicon")


# =================================================================== main
def main():
    print("Drawing artwork...")
    os.makedirs(ASSETS, exist_ok=True)

    write(os.path.join(ASSETS, "mandala.svg"), mandala())
    write(os.path.join(ASSETS, "jali.svg"), jali())
    write(os.path.join(ASSETS, "toran.svg"), toran())
    write(os.path.join(ASSETS, "garland.svg"), garland_vertical())
    write(os.path.join(ASSETS, "kalash.svg"), kalash())
    write(os.path.join(ASSETS, "diyas.svg"), diya_row())
    write(os.path.join(ASSETS, "mandap.svg"), mandap())
    write(os.path.join(ASSETS, "divider.svg"), divider())
    write(os.path.join(ASSETS, "peacock.svg"), peacock())
    write(os.path.join(ASSETS, "corner.svg"), corner())
    write(os.path.join(ASSETS, "hawa-arch.svg"), hawa_arch())
    write(os.path.join(ASSETS, "favicon.svg"), favicon())

    os.makedirs(ICONS, exist_ok=True)
    ic = build_icons()
    for name, markup in sorted(ic.items()):
        write(os.path.join(ICONS, f"{name}.svg"), markup)
    print(f"\n{12 + len(ic)} SVG files written.")


if __name__ == "__main__":
    main()
