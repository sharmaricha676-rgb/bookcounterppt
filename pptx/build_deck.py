#!/usr/bin/env python3
"""Builds BookCounter-interactive.pptx.

Interactivity is all native PowerPoint, so it works in a normal slide show:
  * buttons and hotspots that jump to other slides (a:hlinkClick ppaction://hlinksldjump)
  * hidden slides that only open from those buttons (part close-ups, quiz answers, demo steps)
  * Morph zoom between the device and each part, and Morph on the embedded 3D model
  * a real 3D model (glTF) PowerPoint 365 can rotate, with a picture fallback elsewhere
  * animated GIFs: a turntable loop, the shell opening, the IR beam, and one page turn per demo step

Look: early-2000s "Frutiger Aero" / Windows XP Luna / candy plastic. Sky-gradient slides (the renders are
captured on the same gradient), frosted glass panels, glossy Aqua and chrome pill buttons, and every GIF
or framed render inside an XP-style window, since those carry their own copy of the sky.

usage: build_deck.py <frames-and-stills dir> <gif dir> <glb file> <out.pptx>
"""
import json
import os
import sys

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

import pptx_extras as X  # transitions + 3D model XML (kept separate so it can be reviewed on its own)

ASSETS, GIFS, GLB, OUT = sys.argv[1:5]
WORK = os.path.join(os.path.dirname(os.path.abspath(OUT)), 'build_media')
os.makedirs(WORK, exist_ok=True)

# ---------------------------------------------------------------- look (Y2K)
C = dict(
    navy='0B2559', body='1D3A6B', muted='4F6B95',                       # text
    aqua='00A8FF', aquaInk='0068C2',                                    # aqua accent; the Ink shade is for text on glass
    pink='FF2E88', pinkInk='D0106A', lime='7CCF00', limeInk='3A7400',
    tangerine='FF9A1F', tangInk='B35A00', lilac='B18CFF',
    white='FFFFFF', ice='EEF7FF', rule='B9CDE3', shadow='1450A0',
    xp='0A58D6', xpHi='3A8CFF', xpEdge='0838A8', close='E0451F',
    oled='05080C', oledText='86D6FF', oledDim='5E6A7C',
)
SKY = [(0, '4FB0FF'), (38, '9FD8FF'), (72, 'D9F1FF'), (100, 'F4FBFF')]  # the renders are captured on exactly this
CHROME = [(0, 'FFFFFF'), (100, '9FB3C8')]
HEAD, BODY, MONO = 'Trebuchet MS', 'Verdana', 'Lucida Console'
EYE, PIX, SYM = 'Tahoma', 'Courier New', 'Segoe UI Symbol'  # eyebrows / pixel counters / sparkle glyphs
W_IN, H_IN = 13.333, 7.5
ML = 0.75  # left margin

prs = Presentation()
prs.slide_width = Inches(W_IN)
prs.slide_height = Inches(H_IN)
BLANK = prs.slide_layouts[6]
SLIDES = {}
LINKS = []  # (shape, target key) wired once every slide exists
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'


def rgb(h):
    return RGBColor.from_string(h)


def set_alpha(fill_parent_el, opacity):
    """opacity 0..1 on the first a:srgbClr under a fill element"""
    clr = fill_parent_el.find('.//' + qn('a:srgbClr'))
    a = etree.SubElement(clr, qn('a:alpha'))
    a.set('val', str(int(opacity * 100000)))


# ---------------------------------------------------------------- DrawingML helpers (gradients, gloss, shadows)
def grad_el(stops, ang=90, radial=None):
    """a:gradFill. stops: (pos 0-100, hex[, opacity]). ang: degrees clockwise, 90 = first stop at the top.
    radial: (l, t, r, b) focus rectangle in % for a circular gradient instead of a linear one."""
    gs = []
    for st in stops:
        alpha = '' if len(st) < 3 else '<a:alpha val="%d"/>' % round(st[2] * 100000)
        gs.append('<a:gs pos="%d"><a:srgbClr val="%s">%s</a:srgbClr></a:gs>' % (round(st[0] * 1000), st[1], alpha))
    if radial:
        shade = '<a:path path="circle"><a:fillToRect l="%d" t="%d" r="%d" b="%d"/></a:path>' % tuple(v * 1000 for v in radial)
    else:
        shade = '<a:lin ang="%d" scaled="0"/>' % (ang * 60000)
    return etree.fromstring('<a:gradFill xmlns:a="%s" rotWithShape="1"><a:gsLst>%s</a:gsLst>%s</a:gradFill>'
                            % (A_NS, ''.join(gs), shade))


def _swap_fill(parent, new):
    old = parent.find(qn('a:solidFill'))
    old.addprevious(new)
    parent.remove(old)


def set_grad(shp, stops, ang=90, radial=None):
    shp.fill.solid()
    _swap_fill(shp._element.spPr, grad_el(stops, ang, radial))


def set_line(shp, color=None, width=1.0, stops=None, opacity=None):
    shp.line.width = Pt(width)
    shp.line.fill.solid()
    ln = shp._element.spPr.find(qn('a:ln'))
    if stops:
        _swap_fill(ln, grad_el(stops, 90))
    else:
        shp.line.fill.fore_color.rgb = rgb(color)
        if opacity is not None:
            set_alpha(ln.find(qn('a:solidFill')), opacity)


def set_shadow(shp, blur=0.16, dist=0.05, alpha=0.25, color=C['shadow'], highlight=False):
    """soft blue drop shadow; highlight adds the 1 px white inner line along the top edge"""
    shp.shadow.inherit = False
    eff = shp._element.spPr.find(qn('a:effectLst'))
    if highlight:
        etree.SubElement(eff, qn('a:innerShdw'), blurRad='0', dist='12700', dir='5400000').append(
            etree.fromstring('<a:srgbClr xmlns:a="%s" val="FFFFFF"><a:alpha val="85000"/></a:srgbClr>' % A_NS))
    if blur is not None:
        etree.SubElement(eff, qn('a:outerShdw'), blurRad=str(int(blur * 914400)), dist=str(int(dist * 914400)),
                         dir='5400000', algn='t', rotWithShape='0').append(
            etree.fromstring('<a:srgbClr xmlns:a="%s" val="%s"><a:alpha val="%d"/></a:srgbClr>' % (A_NS, color, alpha * 100000)))


def text_shadow(shp, color='083A80', alpha=0.6):
    """dark glow under white text (the web version's text-shadow)"""
    for r in shp._element.iter(qn('a:rPr')):
        sf = r.find(qn('a:solidFill'))
        eff = etree.fromstring(
            '<a:effectLst xmlns:a="%s"><a:outerShdw blurRad="19050" dist="12700" dir="5400000" algn="t" rotWithShape="0">'
            '<a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:outerShdw></a:effectLst>' % (A_NS, color, alpha * 100000))
        (sf.addnext(eff) if sf is not None else r.insert(0, eff))


def no_theme_effect(shp):
    """effectRef idx 0: no theme effect style (its drop shadow) behind our own effects"""
    ref = shp._element.find('.//' + qn('a:effectRef'))
    if ref is not None:
        ref.set('idx', '0')
    return shp


def shape(s, kind, x, y, w, h, name=None, rad=None):
    shp = no_theme_effect(s.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h)))
    if rad is not None:
        shp.adjustments[0] = min(0.5, rad / max(0.01, min(w, h)))
    shp.shadow.inherit = False
    if name:
        shp.name = name
    return shp


def no_line(shp):
    shp.line.fill.background()
    return shp


def clear_fill(shp):
    """100% transparent fill: invisible but still clickable in a slide show (No Fill shapes are not)"""
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb('FFFFFF')
    set_alpha(shp._element.spPr.find(qn('a:solidFill')), 0)
    return shp


# ---------------------------------------------------------------- media prep
def jpg(src, name, quality=90):
    out = os.path.join(WORK, name)
    Image.open(src).convert('RGB').save(out, quality=quality, optimize=True)
    return out


def crop_square(src, cx, cy, size, name, out_px=420):
    im = Image.open(src).convert('RGB')
    w, h = im.size
    x0, y0 = int(cx * w - size / 2), int(cy * h - size / 2)
    x0, y0 = max(0, min(w - size, x0)), max(0, min(h - size, y0))
    out = os.path.join(WORK, name)
    im.crop((x0, y0, x0 + size, y0 + size)).resize((out_px, out_px), Image.LANCZOS).save(out, quality=88)
    return out


def scrim():
    """white frosted-glass gradient behind the text column of full-bleed slides"""
    out = os.path.join(WORK, 'scrim_glass.png')
    w, h = 1600, 90
    im = Image.new('RGBA', (w, h))
    px = im.load()
    for x in range(w):
        t = x / w
        a = 228 if t < 0.35 else int(228 * max(0.0, 1 - (t - 0.35) / 0.27) ** 1.4)
        for y in range(h):
            px[x, y] = (247, 252, 255, a)
    im.save(out)
    return out


IMG = dict(
    hub=jpg(f'{ASSETS}/hub.png', 'hub.jpg', 92),
    problem=jpg(f'{ASSETS}/problem.png', 'problem.jpg'),
    oled=jpg(f'{ASSETS}/oled.png', 'oled.jpg'),
    wiring=jpg(f'{ASSETS}/wiring.png', 'wiring.jpg'),
    dims=jpg(f'{ASSETS}/dims.png', 'dims.jpg'),
    demo0=jpg(f'{ASSETS}/fwd1/000.png', 'demo0.jpg'),
    scrim=scrim(),
    th_problem=crop_square(f'{ASSETS}/problem.png', 0.68, 0.55, 640, 'th_problem.jpg'),
    th_parts=crop_square(f'{ASSETS}/hub.png', 0.71, 0.52, 1000, 'th_parts.jpg'),
    th_tour=crop_square(f'{ASSETS}/spin/009.png', 0.5, 0.5, 430, 'th_tour.jpg'),
    th_demo=crop_square(f'{ASSETS}/fwd1/012.png', 0.5, 0.52, 470, 'th_demo.jpg'),
    th_quiz=crop_square(f'{ASSETS}/sensor/012.png', 0.5, 0.5, 430, 'th_quiz.jpg'),
    th_turner=crop_square(f'{ASSETS}/turner/012.png', 0.62, 0.55, 430, 'th_turner.jpg'),
    th_build=crop_square(f'{ASSETS}/dims.png', 0.68, 0.5, 700, 'th_build.jpg'),
)
SPOTS = json.load(open(f'{ASSETS}/hub_spots.json'))
SPOTS['sensorL'] = [0.4825, 0.578]  # centre of each board rather than the LED tips
SPOTS['sensorR'] = [0.725, 0.511]
SPOTS['oled'] = [0.5625, 0.393]    # left end of the screen, so its dot clears the button's (they are 0.3 in apart)
SPOTS['pipe'] = [0.815, 0.507]      # the part of the hollow beam you can see, outside the shell
SPOTS.setdefault('base', [0.455, 0.53])  # the cardboard base sheet; older hub_spots.json files have no 'base' key


# ---------------------------------------------------------------- building blocks
def new_slide(key, bg=None, hidden=False, notes=''):
    """every slide sits on the sky gradient (top to bottom), the same one the renders are captured on"""
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid()
    _swap_fill(s._element.cSld.bg.bgPr, grad_el(SKY, 90))
    if hidden:
        s._element.set('show', '0')
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    SLIDES[key] = s
    return s


def P(t=None, runs=None, size=16, color=None, font=BODY, bold=False, italic=False, align='left',
      after=0, line=None, spacing=None):
    return dict(t=t, runs=runs, size=size, color=color or C['body'], font=font, bold=bold, italic=italic,
                align=align, after=after, line=line, spacing=spacing)


ALIGN = dict(left=PP_ALIGN.LEFT, center=PP_ALIGN.CENTER, right=PP_ALIGN.RIGHT)
ANCHOR = dict(top=MSO_ANCHOR.TOP, middle=MSO_ANCHOR.MIDDLE, bottom=MSO_ANCHOR.BOTTOM)


def fill_tf(tf, paras, anchor='top', margins=(0, 0, 0, 0)):
    tf.word_wrap = True
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [Inches(m) for m in margins]
    tf.vertical_anchor = ANCHOR[anchor]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = ALIGN[p['align']]
        para.space_after = Pt(p['after'])
        if p['line']:
            para.line_spacing = p['line']
        for r in (p['runs'] or [dict(t=p['t'])]):
            run = para.add_run()
            run.text = r['t']
            f = run.font
            f.name = r.get('font', p['font'])
            f.size = Pt(r.get('size', p['size']))
            f.bold = r.get('bold', p['bold'])
            f.italic = r.get('italic', p['italic'])
            f.color.rgb = rgb(r.get('color', p['color']))
            if r.get('spacing', p['spacing']):
                f._rPr.set('spc', str(r.get('spacing', p['spacing'])))


def text(s, x, y, w, h, paras, anchor='top', name=None):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    fill_tf(tb.text_frame, paras, anchor)
    if name:
        tb.name = name
    return tb


def rect(s, x, y, w, h, fill=None, line=None, radius=None, opacity=None, name=None, line_w=1.0):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    r = no_theme_effect(s.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h)))
    if radius:
        r.adjustments[0] = radius
    if fill:
        r.fill.solid()
        r.fill.fore_color.rgb = rgb(fill)
        if opacity is not None:
            set_alpha(r._element.spPr.find(qn('a:solidFill')), opacity)
    else:
        r.fill.background()
    if line:
        r.line.color.rgb = rgb(line)
        r.line.width = Pt(line_w)
    else:
        r.line.fill.background()
    r.shadow.inherit = False
    if name:
        r.name = name
    return r


def glass(s, x, y, w, h, opacity=0.62, rad=0.17, name='Glass panel', gloss=0.5, border=True, shadow=True,
          target=None, tint='FFFFFF'):
    """frosted glass: translucent white, chrome edge, soft blue shadow, white gloss over the top half"""
    p = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, name=name, rad=rad)
    p.fill.solid()
    p.fill.fore_color.rgb = rgb(tint)
    set_alpha(p._element.spPr.find(qn('a:solidFill')), opacity)
    if border:
        set_line(p, stops=CHROME, width=2)
    else:
        no_line(p)
    if shadow:
        set_shadow(p, blur=0.2, dist=0.06, alpha=0.22, highlight=True)
    parts = [p]
    if gloss:
        g = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + 0.05, y + 0.05, w - 0.1, h * 0.46, name=name + ' gloss',
                  rad=max(0.02, rad - 0.04))
        set_grad(g, [(0, 'FFFFFF', gloss), (100, 'FFFFFF', 0)], 90)
        no_line(g)
        parts.append(g)
    if target:
        for shp in parts:
            LINKS.append((shp, target))
    return p


def card(s, x, y, w, h, paras, anchor='top', margins=(0.2, 0.18, 0.2, 0.15), name='Card', target=None,
         opacity=0.82, rad=0.14, gloss=0.55, border=True):
    """a glass card with its text on a clear top layer, so the gloss never washes the text out"""
    glass(s, x, y, w, h, opacity=opacity, rad=rad, name=name + ' glass', gloss=gloss, border=border, target=target)
    top = clear_fill(no_line(shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, name=name, rad=rad)))
    fill_tf(top.text_frame, paras, anchor, margins)
    if target:
        LINKS.append((top, target))
    return top


BTN = dict(  # glossy Aqua pill, chrome pill, clear glass pill, flat disabled
    primary=dict(stops=[(0, 'BFE9FF'), (36, '33A1FF'), (100, '0B74E5')], line='0A5FC4', color='FFFFFF', gloss=0.8, tshadow=True),
    secondary=dict(stops=[(0, 'FFFFFF'), (50, 'D5DDE8'), (100, 'AEBCCD')], line='8597AE', color=C['navy'], gloss=0.9, tshadow=False),
    ghost=dict(stops=None, glass=True, line='9FB3C8', color=C['navy'], gloss=0.6, tshadow=False),
    disabled=dict(stops=None, solid='C9D6E6', line='B6C5D8', color='8AA0BD', gloss=0, tshadow=False),
)
BTN_ALIAS = dict(dark='primary', darkline='secondary', menu='secondary')


def button(s, x, y, w, h, label, target=None, style='secondary', size=14, name=None):
    st = BTN[BTN_ALIAS.get(style, style)]
    name = name or f'Button: {label.strip(" ►◄")}'
    base = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, name=name + ' base', rad=h / 2)
    if st['stops']:
        set_grad(base, st['stops'], 90)
    else:
        base.fill.solid()
        base.fill.fore_color.rgb = rgb(st.get('solid', 'FFFFFF'))
        if st.get('glass'):
            set_alpha(base._element.spPr.find(qn('a:solidFill')), 0.55)
    set_line(base, st['line'], 1.0)
    if style != 'disabled':
        set_shadow(base, blur=0.07, dist=0.03, alpha=0.3)
    layers = [base]
    if st['gloss']:
        gh = h * 0.44
        g = no_line(shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + h * 0.2, y + h * 0.06, w - h * 0.4, gh,
                          name=name + ' gloss', rad=gh / 2))
        set_grad(g, [(0, 'FFFFFF', st['gloss']), (100, 'FFFFFF', 0.12)], 90)
        layers.append(g)
    top = clear_fill(no_line(shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, name=name, rad=h / 2)))
    fill_tf(top.text_frame, [P(label, size=size, bold=True, font=HEAD, color=st['color'], align='center')],
            anchor='middle', margins=(0.08, 0.02, 0.08, 0.02))
    if st['tshadow']:
        text_shadow(top)
    layers.append(top)
    if target and style != 'disabled':
        for shp in layers:
            LINKS.append((shp, target))
    return top


def picture(s, path, x, y, w=None, h=None, name=None, alt='', target=None):
    p = s.shapes.add_picture(path, Inches(x), Inches(y), Inches(w) if w else None, Inches(h) if h else None)
    if name:
        p.name = name
    p._element.nvPicPr.cNvPr.set('descr', alt)
    if target:
        LINKS.append((p, target))
    return p


def line(s, x1, y1, x2, y2, color=C['aquaInk'], width=1.25):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = rgb(color)
    c.line.width = Pt(width)
    return c


def sparkle(s, x, y, size=22, color='FFFFFF', char='✦'):
    d = size / 72 * 1.5
    return text(s, x - d / 2, y - d / 2, d, d, [P(char, size=size, font=SYM, color=color, align='center')],
                anchor='middle', name='Sparkle')


def orb(s, cx, cy, d, color=C['aqua'], solid=False, name='Bubble', label=None, size=16):
    """solid: a glossy candy orb (badges, icons). otherwise a see-through soap bubble for decoration."""
    o = shape(s, MSO_SHAPE.OVAL, cx - d / 2, cy - d / 2, d, d, name=name)
    if solid:
        set_grad(o, [(0, 'FFFFFF'), (22, color), (100, color)], radial=(30, 22, 70, 78))
        set_line(o, 'FFFFFF', 1.25)
        set_shadow(o, blur=0.06, dist=0.025, alpha=0.3)
    else:
        set_grad(o, [(0, 'FFFFFF', 0.85), (24, 'FFFFFF', 0.25), (78, color, 0.16), (100, color, 0.5)], radial=(32, 26, 68, 74))
        set_line(o, 'FFFFFF', 1.0, opacity=0.7)
    if label is not None:
        fill_tf(o.text_frame, [P(label, size=size, bold=True, font=HEAD, color='FFFFFF', align='center')], anchor='middle')
        text_shadow(o, alpha=0.5)
    return o


def eyebrow(s, t, x=ML, y=0.72, color=C['aquaInk'], w=7, star=C['pink']):
    return text(s, x, y, w, 0.3, [P(runs=[dict(t='✦  ', font=SYM, color=star, spacing=0), dict(t=t.upper())],
                                    size=12, bold=True, font=EYE, color=color, spacing=200)])


def title(s, t, x=ML, y=1.05, w=5.4, h=1.4, size=38, color=C['navy']):
    return text(s, x, y, w, h, [P(t, size=size, bold=True, font=HEAD, color=color, line=0.92)])


def menu_pill(s, style='menu', x=11.83, y=0.42, w=0.9, h=0.36):
    return button(s, x, y, w, h, 'Menu', 'menu', style=style, size=11, name='Menu button')


def full_bleed(s, path, alt, name=None, with_scrim=True, dx=0.0, crop_bottom=0.0):
    p = picture(s, path, dx, 0, W_IN, H_IN * (1 - crop_bottom), name=name, alt=alt)
    if crop_bottom:
        p.crop_bottom = crop_bottom
    if dx:  # keep the picture inside the slide: trim what would hang off the right edge
        p.crop_right = dx / W_IN
        p.width = Inches(W_IN - dx)
    if with_scrim:
        picture(s, IMG['scrim'], 0, 0, W_IN, H_IN, name='Glass scrim', alt='')


TB, BW = 0.34, 0.06  # XP window: title bar height, border width


def xp_window(s, x, y, w, h, caption, name='Window', body=None):
    """Windows XP Luna window whose BODY is exactly (x, y, w, h); the media goes there afterwards.
    body: None (media covers it), 'sky' (the slide gradient, for the 3D model) or a hex fill.
    Returns the x where a title-bar button (left of the caption buttons) can end."""
    fx, fy, fw, fh = x - BW, y - TB, w + 2 * BW, h + TB + BW
    fr = shape(s, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, fx, fy, fw, fh, name=name + ' frame')
    fr.adjustments[0] = 0.11 / min(fw, fh)
    fr.adjustments[1] = 0.0
    fr.fill.solid()
    fr.fill.fore_color.rgb = rgb(C['xp'])
    set_line(fr, C['xpEdge'], 0.75)
    set_shadow(fr, blur=0.22, dist=0.08, alpha=0.32)
    bar = no_line(shape(s, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, fx, fy, fw, TB, name=name + ' title bar'))
    bar.adjustments[0] = 0.11 / TB
    bar.adjustments[1] = 0.0
    set_grad(bar, [(0, C['xp']), (32, C['xpHi']), (100, C['xp'])], 90)
    line(s, fx + 0.09, fy + 0.018, fx + fw - 0.09, fy + 0.018, color='9CCBFF', width=1.0)  # the lighter top line
    if body:
        b = no_line(shape(s, MSO_SHAPE.RECTANGLE, x, y, w, h, name=name + ' body'))
        if body == 'sky':
            set_grad(b, SKY, 90)
        else:
            b.fill.solid()
            b.fill.fore_color.rgb = rgb(body)
    orb(s, fx + 0.2, fy + TB / 2, 0.19, color=C['aqua'], solid=True, name=name + ' icon')
    bs, gap = 0.24, 0.035
    by = fy + (TB - bs) / 2
    xs = [fx + fw - 0.07 - bs * (3 - i) - gap * (2 - i) for i in range(3)]  # minimise, maximise, close
    for kind, bx in zip(('min', 'max', 'close'), xs):
        cb = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, bs, bs, name=f'{name} {kind} button', rad=0.045)
        if kind == 'close':
            set_grad(cb, [(0, 'F7A486'), (45, C['close']), (100, 'B32E0E')], 90)
        else:
            set_grad(cb, [(0, '8DBBFF'), (45, '3D85F2'), (100, '1B55CF')], 90)
        set_line(cb, 'FFFFFF', 1.0)
        if kind == 'close':
            g = no_line(shape(s, MSO_SHAPE.MATH_MULTIPLY, bx + 0.025, by + 0.025, bs - 0.05, bs - 0.05, name=f'{name} close x'))
            g.fill.solid(); g.fill.fore_color.rgb = rgb('FFFFFF')
        elif kind == 'max':
            sq = shape(s, MSO_SHAPE.RECTANGLE, bx + 0.06, by + 0.06, bs - 0.12, bs - 0.12, name=f'{name} max glyph')
            sq.fill.background(); set_line(sq, 'FFFFFF', 1.0)
            t = no_line(shape(s, MSO_SHAPE.RECTANGLE, bx + 0.06, by + 0.06, bs - 0.12, 0.03, name=f'{name} max glyph top'))
            t.fill.solid(); t.fill.fore_color.rgb = rgb('FFFFFF')
        else:
            t = no_line(shape(s, MSO_SHAPE.RECTANGLE, bx + 0.06, by + bs - 0.09, 0.08, 0.032, name=f'{name} min glyph'))
            t.fill.solid(); t.fill.fore_color.rgb = rgb('FFFFFF')
    cap = text(s, fx + 0.36, fy, xs[0] - fx - 0.5, TB,
               [P(caption, size=11.5, bold=True, font=HEAD, color='FFFFFF')], anchor='middle', name=name + ' caption')
    text_shadow(cap, color='06307A', alpha=0.7)
    return xs[0] - 0.1


def window_media(s, path, x, y, w, caption, alt, name, aspect=9 / 16):
    """a GIF or render in an XP window; its own sky stays inside the window instead of clashing with the slide's"""
    h = w * aspect
    xp_window(s, x, y, w, h, caption, name=name)
    return picture(s, path, x, y, w, h, name=name, alt=alt)


# right-hand media window used on the text + media slides
RX, RW = 6.3, 6.62
RH = RW * 9 / 16


def left_panel(s, y=0.45, h=6.6, w=5.5, x=0.42, **kw):
    return glass(s, x, y, w, h, name='Text panel', **kw)


# ================================================================ slides
# 1 ---------------------------------------------------------------- title
s = new_slide('title', notes=(
    "Hello everyone. This is Book Counter, a bookmark we designed and 3D-printed that counts the pages you turn, "
    "and with its automatic page turner it can turn the page for you at the press of a button. "
    "This presentation is interactive: the buttons work during the slide show. Click Start, or press the right arrow."))
orb(s, 12.55, 6.55, 1.25, color=C['aqua'])
orb(s, 6.05, 6.75, 0.55, color=C['lilac'])
orb(s, 12.75, 0.7, 0.6, color=C['pink'])
window_media(s, f'{GIFS}/spin.gif', RX, 1.85, RW, 'spin.gif', 'The Book Counter turning on a turntable', name='Spinning model')
left_panel(s, y=0.95, h=6.1)
eyebrow(s, '3D-printed smart bookmark', y=1.3)
text(s, ML, 1.65, 5.0, 1.75, [P('BOOK', size=60, bold=True, font=HEAD, color=C['navy'], line=0.88),
                              P('COUNTER', size=60, bold=True, font=HEAD, color=C['aquaInk'], line=0.88)])
sparkle(s, 3.55, 1.95, 26, C['pink'])
sparkle(s, 4.05, 1.68, 14, C['aqua'], '✧')
text(s, ML, 3.5, 4.85, 0.95, [P('A bookmark that counts every page you turn, and turns the page for you at the press of a button.',
                               size=16, color=C['body'], line=1.1)])
for i, (num, lab, col) in enumerate([('2', 'IR sensors', C['pinkInk']), ('0.96"', 'OLED screen', C['aquaInk']),
                                     ('1', 'page turner', C['limeInk'])]):
    card(s, ML + i * 1.6, 4.62, 1.48, 0.86, [P(num, size=19, bold=True, font=PIX, color=col),
                                             P(lab.upper(), size=8.5, bold=True, font=EYE, color=C['muted'], spacing=60)],
         anchor='middle', margins=(0.14, 0.04, 0.06, 0.04), name=f'Chip: {lab}', opacity=0.75, rad=0.12, gloss=0.6)
text(s, ML, 5.7, 4.8, 0.3, [P('[Presenter names · Class]', size=11, bold=True, font=EYE, color=C['muted'])])
button(s, ML, 6.2, 1.6, 0.5, 'Start  ►', 'problem', 'primary', size=14)
button(s, ML + 1.75, 6.2, 1.25, 0.5, 'Menu', 'menu', 'secondary', size=14)

# 2 ---------------------------------------------------------------- menu
s = new_slide('menu', notes=(
    "This is the menu. Every tile is a button: click one to jump straight to that part of the talk. "
    "If you just press the right arrow, the presentation goes through everything in order."))
orb(s, 12.6, 0.75, 0.9, color=C['aqua'])
glass(s, 0.42, 0.42, 12.5, 6.66, opacity=0.5, name='Menu panel', gloss=0.4)
eyebrow(s, 'Explore')
title(s, 'Where do you want to go?', w=11)
TILES = [
    ('START HERE', 'The problem', 'Why paper books need a counter', 'problem', 'th_problem'),
    ('CLICK THE PARTS', 'The device', 'Explore it piece by piece', 'hub', 'th_parts'),
    ('REAL 3D', '3D tour', 'A 3D model you can turn', 'tour1', 'th_tour'),
    ('TRY IT', 'Live demo', 'Turn pages, watch it count', 'demo0', 'th_demo'),
    ('PRESS A BUTTON', 'Page turner', 'It turns the page for you', 'turner', 'th_turner'),
    ('HOW IT IS MADE', 'Build', 'Wiring and the printed shell', 'wiring', 'th_build'),
]
for i, (lab, name, desc, target, th) in enumerate(TILES):
    col, row = i % 3, i // 3
    x, y, w, h = ML + col * 4.05, 2.15 + row * 2.2, 3.8, 2.0
    card(s, x, y, w, h, [P(lab, size=9.5, bold=True, font=EYE, color=C['aquaInk'], spacing=120, after=6),
                         P(name, size=19, bold=True, font=HEAD, color=C['navy'], after=4, line=0.95),
                         P(desc, size=11, color=C['muted'], line=1.05)],
         margins=(0.22, 0.22, 1.85, 0.18), name=f'Tile: {name}', target=target, opacity=0.8)
    th_pic = picture(s, IMG[th], x + w - 1.68, y + 0.27, 1.45, 1.45, name=f'Thumb: {name}', alt=name, target=target)
    th_pic.auto_shape_type = MSO_SHAPE.ROUNDED_RECTANGLE
    th_pic.line.color.rgb = rgb('FFFFFF')
    th_pic.line.width = Pt(2.25)
    set_shadow(th_pic, blur=0.08, dist=0.03, alpha=0.3)
text(s, ML, 6.58, 9, 0.3, [P('Each tile is a button. Press → to go through everything in order.', size=11, color=C['muted'])])

# 3 ---------------------------------------------------------------- problem
s = new_slide('problem', notes=(
    "E-readers like a Kindle track exactly how much you read. Paper books don't. If you want a reading goal, "
    "like twenty pages a day, you have to remember where you started and do the maths yourself, and most people forget. "
    "Reading apps help, but you still have to type your pages in every time. And turning pages needs a free hand, "
    "which is not always easy."))
full_bleed(s, IMG['problem'], 'An open book on a desk with the Book Counter above its top edge')
menu_pill(s)
eyebrow(s, 'Why we built it')
title(s, 'Paper books keep no score')
text(s, ML, 2.55, 4.4, 0.8, [P("An e-reader knows how far you got. A paper book doesn't.", size=17, color=C['body'], line=1.1)])
BUL = lambda t: P(runs=[dict(t='✦  ', font=SYM, color=C['pinkInk']), dict(t=t)], size=15, after=10, line=1.1)
text(s, ML, 3.55, 4.4, 2.6, [BUL('Reading goals need numbers: pages a day, pages a week.'),
                             BUL('Page numbers written down by hand get forgotten.'),
                             BUL('Reading apps make you type in every session.'),
                             BUL('Turning pages needs a free hand.')])

# 4 ---------------------------------------------------------------- idea
s = new_slide('idea', notes=(
    "So our idea was simple: make the bookmark do the counting. You put it on your book, read normally, "
    "and it counts every page you turn by itself. Then we went one step further and made it turn the page too."))
orb(s, 11.35, 1.55, 2.4, color=C['aqua'])
orb(s, 12.35, 5.85, 1.3, color=C['pink'])
orb(s, 0.65, 6.85, 0.8, color=C['lilac'])
glass(s, 0.42, 1.15, 11.3, 5.6, opacity=0.6, name='Idea panel')
sparkle(s, 10.95, 1.62, 30, 'FFFFFF')
sparkle(s, 11.45, 2.25, 16, C['pink'], '✧')
menu_pill(s, 'darkline')
eyebrow(s, 'The idea', y=1.6)
text(s, ML, 2.05, 10.4, 2.4, [P('What if the bookmark kept count for you?', size=54, bold=True, font=HEAD, color=C['navy'], line=0.95)])
text(s, ML, 4.6, 9.8, 1.0, [P('Book Counter sits at the top of the book and counts each page as it turns. '
                              'Press its button and it even turns the page for you.', size=18, color=C['body'], line=1.15)])
button(s, ML, 5.85, 2.75, 0.55, 'Meet the device  ►', 'hub', 'dark', size=14)

# 5 ---------------------------------------------------------------- hub with hotspots
PARTS = [  # key, name, spec, description, (label dx, dy, width) from the dot, speaker note
    ('oled', 'OLED display', '0.96" OLED · 128 × 64 px · SSD1306 · I²C',
     'Shows the running page count. It needs only two data wires, SDA and SCL, to the microcontroller.',
     (-1.72, -1.03, 1.6), 'This is the screen. It shows the count, and it only needs two data wires to talk to the ESP32.'),
    ('button', 'Turn button', '6 mm push button',
     'Press it and the automatic page turner turns one page for you. The sensors then count that page like any other.',
     (0.55, -1.33, 1.55), 'This is the turn button. Press it and the page turner turns one page for you.'),
    ('shell', '3D-printed shell', 'PLA · 105 × 60 × 35 mm · two pieces',
     'A base tray and a top plate with cut-outs for the screen, the button and the sensor cables. '
     'BOOK COUNTER is embossed on the front.',
     (-0.83, 1.53, 1.95), 'We designed this shell and printed it in two pieces, with the name embossed on the front.'),
    ('sensorL', 'IR sensor, left', 'IR obstacle module (FC-51) · LM393',
     'Watches the left side of the book. When a page goes forward, it sees the page second, as the page lands.',
     (-1.23, 1.62, 1.85), 'The left sensor sees the page second when it goes forward, as the page lands on the left.'),
    ('sensorR', 'IR sensor, right', 'IR obstacle module (FC-51) · LM393',
     'Watches the right side of the book. When a page goes forward, it sees the page first, as it lifts.',
     (-0.22, 2.12, 1.85), 'The right sensor sees the page first when it goes forward, as the page lifts off the right side.'),
    ('pipe', 'Hollow beam', '3D-printed hollow structure · 180 × 16 × 13 mm',
     'A hollow 3D-printed beam joins the shell to the page turner and holds it over the page. The servo cable runs inside it.',
     (-0.27, -1.05, 1.45), 'This hollow beam is 3D-printed too. It holds the page turner out over the page, and the servo cable runs inside it.'),
    ('servo', 'Page turner', 'Servo motor (SG90 size) · rubber wheel',
     'When the turn button is pressed, the motor driver powers the servo, the servo spins the rubber wheel, '
     'and the wheel pushes the top page over.',
     (-0.74, 0.69, 1.6), 'This is the automatic page turner: a servo motor that spins a rubber wheel resting on the page.'),
    ('base', 'Cardboard base', 'Corrugated cardboard sheet',
     'A sheet of cardboard under the bookmark gives everything one flat base to sit on.',
     (-2.1, 0.45, 1.7), 'This is the cardboard base: a sheet of corrugated cardboard under the whole bookmark, '
                        'so everything has one flat base to sit on.'),
]
s = new_slide('hub', notes=(
    "Here's the device, sitting on its cardboard base. Everything you can see has one job, including the page turner "
    "on its hollow beam, and the sheet of corrugated cardboard underneath that gives it all one flat base. "
    "Click any label, or any blue dot, to zoom in on that part. Each close-up has a Back button."))
picture(s, IMG['hub'], 0, 0, W_IN, H_IN, name='!!device', alt='The Book Counter on its cardboard base: white printed box with '
        'OLED window, turn button, two blue IR sensor boards, and a hollow beam holding the servo page turner')
HUB_CARD = (0.42, 0.45, 4.95, 3.62)
glass(s, *HUB_CARD, opacity=0.66, name='Text panel')
menu_pill(s)
eyebrow(s, 'The device')
title(s, 'Two eyes, a screen and a hand', w=4.5, size=34)
text(s, ML, 2.65, 4.45, 1.2, [P('Click any part of the model to look closer.', size=16, color=C['body'], line=1.1, after=8),
                              P('Every label and blue dot is a button.', size=13, color=C['muted'])])
card_bottom = HUB_CARD[1] + HUB_CARD[3]
boxes = []
for key, name, spec, desc, (ox, oy, pw), _ in PARTS:
    u, v = SPOTS[key]
    dx, dy = u * W_IN, v * H_IN
    py = min(H_IN - 0.62, max(0.95, dy + oy))
    left_lim = ML if py >= card_bottom + 0.15 else HUB_CARD[0] + HUB_CARD[2] + 0.15  # clear of the text panel
    px = min(W_IN - 0.45 - pw, max(left_lim, dx + ox))
    # leader from the dot to the nearest edge of its label
    lx = px + pw if px + pw < dx else (px if px > dx else dx)
    ly = py + 0.4 if py + 0.4 < dy else (py if py > dy else py + 0.2)
    line(s, dx, dy, lx, ly, width=1.5)
    halo = shape(s, MSO_SHAPE.OVAL, dx - 0.2, dy - 0.2, 0.4, 0.4, name=f'Hotspot halo: {name}')
    halo.fill.solid(); halo.fill.fore_color.rgb = rgb(C['aqua']); set_alpha(halo._element.spPr.find(qn('a:solidFill')), 0.3)
    set_line(halo, 'FFFFFF', 0.75, opacity=0.8)
    dot = shape(s, MSO_SHAPE.OVAL, dx - 0.1, dy - 0.1, 0.2, 0.2, name=f'Hotspot: {name}')
    set_grad(dot, [(0, 'FFFFFF'), (30, '7FD3FF'), (100, '0082E0')], radial=(30, 25, 70, 75))
    set_line(dot, 'FFFFFF', 1.5)
    set_shadow(dot, blur=0.05, dist=0.02, alpha=0.4)
    pill = rect(s, px, py, pw, 0.4, fill='FFFFFF', line=C['aqua'], radius=0.5, opacity=0.92, name=f'Label: {name}', line_w=1.5)
    set_shadow(pill, blur=0.08, dist=0.03, alpha=0.28)
    fill_tf(pill.text_frame, [P(name, size=12, bold=True, font=HEAD, color=C['navy'], align='center')], anchor='middle',
            margins=(0.06, 0.02, 0.06, 0.02))
    for shp in (halo, dot, pill):
        LINKS.append((shp, f'det_{key}'))
    boxes.append((name, (px, py, px + pw, py + 0.4), (dx - 0.2, dy - 0.2, dx + 0.2, dy + 0.2)))


def _hit(a, b):
    return a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]


for i, (n1, lab1, dot1) in enumerate(boxes):  # layout check: labels inside the slide, clear of each other and every dot
    if lab1[0] < 0 or lab1[1] < 0 or lab1[2] > W_IN or lab1[3] > H_IN:
        print(f'WARNING hub: label "{n1}" leaves the slide', file=sys.stderr)
    for j, (n2, lab2, dot2) in enumerate(boxes):
        if i < j and _hit(lab1, lab2):
            print(f'WARNING hub: labels "{n1}" and "{n2}" overlap', file=sys.stderr)
        if _hit(lab1, dot2):
            print(f'WARNING hub: label "{n1}" covers the dot of "{n2}"', file=sys.stderr)

# 6-13 ------------------------------------------------------------- part close-ups (hidden, Morph zoom)
Z = 2.4
for i, (key, name, spec, desc, _, say) in enumerate(PARTS):
    s = new_slide(f'det_{key}', hidden=True, notes=say + ' Click Back to return to the whole device, or Next part.')
    u, v = SPOTS[key]
    iw, ih = W_IN * Z, H_IN * Z
    left = min(0.0, max(W_IN - iw, 8.9 - u * iw))
    top = min(0.0, max(H_IN - ih, 3.9 - v * ih))
    picture(s, IMG['hub'], left, top, iw, ih, name='!!device', alt=f'Close-up of the {name}')
    glass(s, 0.35, 0.35, 5.2, 6.8, opacity=0.9, name='Detail panel')
    menu_pill(s)
    eyebrow(s, f'Part {i + 1} of {len(PARTS)}')
    text(s, ML, 1.05, 4.55, 1.5, [P(name, size=34, bold=True, font=HEAD, color=C['navy'], line=0.95)])
    text(s, ML, 2.6, 4.55, 0.7, [P(spec, size=12, bold=True, font=MONO, color=C['aquaInk'], line=1.15)])
    text(s, ML, 3.35, 4.45, 2.2, [P(desc, size=15, color=C['body'], line=1.15)])
    button(s, ML, 5.75, 2.55, 0.52, '◄  Back to the device', 'hub', 'primary', size=13)
    nxt = PARTS[(i + 1) % len(PARTS)][0]
    button(s, ML + 2.7, 5.75, 1.75, 0.52, 'Next part  ►', f'det_{nxt}', 'secondary', size=13)

# 14-17 ------------------------------------------------------------ 3D tour (real 3D model, Morph between views)
TOUR = [
    ('tour1', 'A real 3D model', 'This is the actual 3D file of the device. Press → and PowerPoint turns it for you.',
     'In edit mode, click the model and drag the round arrow in its middle to spin it any way you like.', 'front'),
    ('tour2', 'The sensor arms', 'An IR board hangs off each front corner. Both look down towards the pages.',
     'The jumper wires go into the shell through two ports in the top plate.', 'side'),
    ('tour3', 'Screen on top', 'The OLED window and the turn button sit on the top plate, where you can reach them while reading.',
     'The screen shows the count in big pixel digits.', 'top'),
    ('tour4', 'All the way round', 'From here you can see the hollow beam that carries the page turner out over the page.',
     'Press → to continue, or click Menu to jump somewhere else.', 'back'),
]
for i, (key, ttl, line1, line2, view) in enumerate(TOUR):
    s = new_slide(key, notes=(
        f"{line1} {line2} This model is the same 3D file we used in the live 3D version. "
        "Between these slides PowerPoint's Morph transition turns it smoothly."))
    # the window goes in first so the model sits on top of its sky-filled body
    bar_x = xp_window(s, 4.9, 0.55, 8.1, 6.4, 'book-counter.glb', name='3D window', body='sky')
    X.add_model3d(prs, s, GLB, view, WORK, ASSETS, x=4.9, y=0.55, w=8.1, h=6.4, name='!!model')
    menu_pill(s, x=bar_x - 0.82, y=0.55 - TB + 0.045, w=0.82, h=TB - 0.09)  # in the window's title bar
    left_panel(s, w=4.25)
    eyebrow(s, f'3D tour · {i + 1} of {len(TOUR)}')
    title(s, ttl, w=3.8)
    text(s, ML, 2.55, 3.75, 1.6, [P(line1, size=16, color=C['body'], line=1.12)])
    text(s, ML, 4.25, 3.75, 1.2, [P(line2, size=12, color=C['muted'], line=1.12)])
    if i < len(TOUR) - 1:
        button(s, ML, 5.9, 1.9, 0.5, 'Turn it  ►', TOUR[i + 1][0], 'primary', size=13)
    else:
        button(s, ML, 5.9, 1.9, 0.5, 'Inside  ►', 'inside', 'primary', size=13)
    if i > 0:
        button(s, ML + 2.05, 5.9, 1.4, 0.5, '◄  Back', TOUR[i - 1][0], 'secondary', size=13)

# 18 --------------------------------------------------------------- inside
s = new_slide('inside', notes=(
    "If we lift the top plate off, you can see the inside. The microcontroller is an ESP32 DevKit, the black board: it reads the sensors "
    "and keeps the count. The OLED screen is mounted under the window in the top plate. The red board is the motor driver "
    "for the page turner. Behind the shell sits the battery pack: two 18650 lithium cells that power everything. The USB cable is for uploading the code."))
window_media(s, f'{GIFS}/explode.gif', RX, 1.85, RW, 'explode.gif',
             'The shell opening up to show the parts inside, then closing again', name='Exploded view')
menu_pill(s)
left_panel(s)
eyebrow(s, 'Inside the shell')
title(s, "What's inside")
text(s, ML, 2.0, 4.8, 0.8, [P('Lift off the top plate and four parts do the real work.', size=16, color=C['body'], line=1.1)])
for i, (nm, ds) in enumerate([('Microcontroller', 'ESP32 DevKit. Reads, decides, counts.'),
                              ('OLED display', '0.96", under the window.'),
                              ('Motor driver', 'Gives the servo the power to move.'),
                              ('Battery pack', 'Two 18650 cells power everything.')]):
    col, row = i % 2, i // 2
    card(s, ML + col * 2.48, 3.0 + row * 1.32, 2.36, 1.18, [P(nm, size=14, bold=True, font=HEAD, color=C['navy'], after=4),
                                                           P(ds, size=11, color=C['muted'], line=1.05)],
         margins=(0.16, 0.14, 0.12, 0.1), name=f'Card: {nm}', opacity=0.85, rad=0.12, gloss=0.6)
button(s, ML, 5.85, 2.3, 0.5, 'Click the parts  ►', 'hub', 'secondary', size=13)

# 19 --------------------------------------------------------------- sensor
s = new_slide('sensor', notes=(
    "Each sensor is an infrared module. The clear LED sends out infrared light, which we can't see; the pink cone just "
    "shows where it goes. When a page moves into the beam, it reflects the light back to the dark photodiode. "
    "The LM393 chip turns that into a clean ON or OFF signal for the ESP32, and a small green LED lights up. "
    "The blue screw is a trimmer: we turn it to set how close a page has to be before it counts."))
window_media(s, f'{GIFS}/sensor.gif', RX, 1.85, RW, 'ir_sensor.gif',
             'A page lifting into the right sensor beam, which turns brighter', name='Sensor view')
menu_pill(s)
left_panel(s)
eyebrow(s, 'Sensing')
title(s, 'Seeing a page')
for i, (b, t) in enumerate([('Send.', 'The clear LED shines infrared light.'),
                            ('Reflect.', 'A page in the beam bounces it back.'),
                            ('Receive.', 'The dark photodiode picks it up.'),
                            ('Decide.', 'The LM393 chip outputs a clean ON or OFF.')]):
    y = 2.05 + i * 0.78
    orb(s, ML + 0.23, y + 0.23, 0.46, color=C['pink'] if i == 0 else '1E90FF', solid=True, name=f'Step {i + 1}',
        label=str(i + 1), size=15)
    text(s, ML + 0.65, y + 0.02, 4.4, 0.7, [P(runs=[dict(t=b + ' ', bold=True, color=C['navy'], font=HEAD), dict(t=t)],
                                              size=15, line=1.05)])
text(s, ML, 5.2, 4.7, 0.7, [P('The blue trimmer sets how close a page must come. Real IR is invisible; pink shows the beam.',
                              size=11, color=C['muted'], line=1.1)])
button(s, ML, 6.05, 2.3, 0.5, 'Page turner  ►', 'turner', 'primary', size=13)

# 20 --------------------------------------------------------------- automatic page turner
s = new_slide('turner', notes=(
    "This is our automatic page turner. When you press the turn button, the ESP32 sends a signal to the motor driver. "
    "The motor driver powers the servo motor, and the servo spins a rubber wheel that rests on the page. The wheel grips "
    "the top page and pushes it over. Because the page passes both IR sensors, it gets counted just like a page you turn "
    "by hand. The servo hangs from a hollow 3D-printed beam, and its cable runs inside the beam."))
window_media(s, f'{GIFS}/turner.gif', 6.75, 1.6, 6.1, 'page_turner.exe',
             'The servo spins its rubber wheel and the page turns over', name='Page turner view', aspect=540 / 760)
menu_pill(s)
left_panel(s, y=0.45, h=6.68, w=5.55)
eyebrow(s, 'Automatic page turner')
title(s, 'Press a button, the page turns', w=5.2)
for i, (b, t) in enumerate([('Press.', 'You press the turn button.'),
                            ('Signal.', 'The ESP32 tells the motor driver.'),
                            ('Power.', 'The motor driver powers the servo.'),
                            ('Spin.', 'The servo spins the rubber wheel.'),
                            ('Turn.', 'The wheel pushes the top page over.'),
                            ('Count.', 'Both IR sensors see it pass: +1.')]):
    y = 2.4 + i * 0.57
    orb(s, ML + 0.21, y + 0.21, 0.42, color=C['lime'] if i == 5 else '1E90FF', solid=True, name=f'Step {i + 1}',
        label=str(i + 1), size=14)
    text(s, ML + 0.6, y + 0.05, 4.5, 0.5, [P(runs=[dict(t=b + ' ', bold=True, color=C['navy'], font=HEAD), dict(t=t)],
                                             size=14, line=1.0)])
text(s, ML, 5.85, 4.8, 0.5, [P('The servo hangs from the hollow beam, a 3D-printed hollow structure, 180 × 16 × 13 mm. '
                              'Its cable runs inside.', size=11, color=C['muted'], line=1.1)])
button(s, ML, 6.45, 2.3, 0.48, 'Quick quiz  ►', 'quiz', 'primary', size=13)
button(s, ML + 2.45, 6.45, 2.3, 0.48, 'Try it in the demo  ►', 'demo0', 'secondary', size=12)

# 21 --------------------------------------------------------------- quiz
s = new_slide('quiz', notes=(
    "Quick question for the audience: when you turn a page forward, which sensor sees it first? "
    "Let someone answer, then click the button they chose."))
glass(s, 0.42, 0.42, 12.5, 6.66, opacity=0.5, name='Quiz panel', gloss=0.4)
menu_pill(s)
eyebrow(s, 'Quick check')
text(s, ML, 1.05, 11.3, 1.6, [P('You turn a page forward. Which sensor sees it first?', size=36, bold=True, font=HEAD,
                                color=C['navy'], line=0.95)])
for i, (lab, sub, target) in enumerate([('Left sensor', 'The board on the left corner', 'quiz_wrong'),
                                        ('Right sensor', 'The board on the right corner', 'quiz_right')]):
    card(s, ML + i * 6.0, 3.15, 5.75, 2.2, [P(lab, size=30, bold=True, font=HEAD, color=C['navy'], align='center', after=6),
                                            P(sub, size=13, color=C['muted'], align='center')],
         anchor='middle', name=f'Answer: {lab}', target=target, opacity=0.85, rad=0.18, gloss=0.65)
    LINKS.append((orb(s, ML + i * 6.0 + 0.55, 3.7, 0.5, color='1E90FF', solid=True, name=f'Answer badge {i + 1}',
                      label='LR'[i], size=14), target))
text(s, ML, 5.75, 9, 0.4, [P('Click an answer.', size=13, color=C['muted'])])

s = new_slide('quiz_wrong', hidden=True, notes=(
    "Not quite. Ask them to picture the page: it starts on the right-hand side of the open book. Then click Try again."))
glass(s, 0.42, 0.42, 12.5, 5.0, opacity=0.6, name='Answer panel')
orb(s, 11.3, 2.05, 1.5, color=C['tangerine'], solid=True, name='Not quite badge', label='?', size=54)
menu_pill(s)
eyebrow(s, 'Not quite', color=C['tangInk'], star=C['tangerine'])
text(s, ML, 1.05, 9.8, 1.2, [P('Think about where the page starts', size=40, bold=True, font=HEAD, color=C['navy'])])
text(s, ML, 2.45, 9.4, 1.4, [P('You read the right-hand page, then lift it from the right side. Which sensor does it pass first?',
                               size=18, color=C['body'], line=1.15)])
button(s, ML, 4.2, 2.0, 0.56, '◄  Try again', 'quiz', 'primary', size=14)

s = new_slide('quiz_right', hidden=True, notes=(
    "Correct. The page starts on the right, so as you lift it, it passes the right sensor first and then the left one as "
    "it lands. Right then left means plus one. That order is how the bookmark knows which way you turned."))
orb(s, 12.4, 6.4, 1.4, color=C['aqua'])
glass(s, 0.42, 0.95, 12.5, 5.6, opacity=0.6, name='Answer panel')
orb(s, 11.1, 2.45, 1.7, color=C['lime'], solid=True, name='Correct badge', label='✓', size=60)
sparkle(s, 12.15, 1.55, 26, C['pink'])
sparkle(s, 10.05, 1.5, 16, 'FFFFFF', '✧')
menu_pill(s, 'darkline')
eyebrow(s, 'Correct', color=C['limeInk'], y=1.4, star=C['lime'])
text(s, ML, 1.85, 9.6, 1.2, [P('Right, then left', size=54, bold=True, font=HEAD, color=C['navy'])])
text(s, ML, 3.2, 9.4, 1.6, [P('A page you turn forward lifts off the right side, so the right sensor sees it first. '
                              'Then it lands on the left. Right then left means one more page read.', size=18, color=C['body'], line=1.15)])
button(s, ML, 5.3, 3.4, 0.56, 'Try it in the live demo  ►', 'demo0', 'dark', size=14)
button(s, ML + 3.6, 5.3, 2.5, 0.56, '◄  Back to the question', 'quiz', 'darkline', size=13)


# 24-34 ------------------------------------------------------------ live demo: one slide per count
def demo_slide(key, count, media, alt, status, hidden, last_dir):
    notes = {
        None: "This is the live demo. Click Press turn button: each click is one press of the real button, so the page turner "
              "turns one page and the count goes up. Flip back by hand takes it down again. Start over rewinds this demo to zero.",
        1: f"The page turner's wheel pushed the page over. It lifted off the right side first, then landed on the left. "
           f"Right then left: plus one. The count is now {count}.",
        -1: f"This time we flipped the page back by hand: left sensor first, then right. That means minus one. The count is now {count}.",
    }[last_dir]
    s = new_slide(key, hidden=hidden, notes=notes)
    window_media(s, media, 6.12, 1.75, 6.8, 'live_demo.exe', alt, name='Demo view')
    menu_pill(s)
    left_panel(s, w=5.4)
    eyebrow(s, 'Live demo')
    title(s, 'Turn a page, watch it count', w=4.6)
    bezel = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, ML, 2.45, 2.95, 1.35, name='Counter bezel', rad=0.16)
    set_grad(bezel, [(0, 'FFFFFF'), (45, 'D5DDE8'), (100, '9FB3C8')], 90)
    set_line(bezel, '8597AE', 0.75)
    set_shadow(bezel, blur=0.1, dist=0.04, alpha=0.3)
    box = rect(s, ML + 0.1, 2.55, 2.75, 1.15, fill=C['oled'], line='2A3647', radius=0.08, name='Counter')
    fill_tf(box.text_frame, [P(f'{count:03d}', size=46, bold=True, font=PIX, color=C['oledText'], align='center'),
                             P('PAGES TURNED', size=9, bold=True, font=EYE, color=C['oledDim'], align='center', spacing=150)],
            anchor='middle')
    # pixel progress bar: one block per page, out of five
    text(s, 3.9, 2.5, 1.6, 0.3, [P(f'PAGE {count}/5', size=10, bold=True, font=PIX, color=C['muted'])])
    for j in range(5):
        lit = j < count
        rect(s, 3.92 + j * 0.27, 2.85, 0.21, 0.21, fill=C['lime'] if lit else 'DCE7F3', line='4F8F00' if lit else 'A9BDD4',
             name=f'Progress {j + 1}', line_w=0.75)
    text(s, 3.9, 3.2, 1.6, 0.5, [P('BATT · IR OK', size=9, bold=True, font=PIX, color=C['muted'])])
    text(s, ML, 3.98, 4.6, 0.6, [P(status, size=13, color=C['body'])])
    nxt = f'fwd{count + 1}' if count < 5 else None
    prv = f'back{count - 1}' if count > 0 else None
    button(s, ML, 4.7, 2.4, 0.55, 'Press turn button  ►', nxt, 'primary' if nxt else 'disabled', size=12)
    button(s, ML + 2.52, 4.7, 2.2, 0.55, '◄  Flip back by hand', prv, 'secondary' if prv else 'disabled', size=11.5)
    button(s, ML, 5.42, 1.35, 0.46, 'Start over', 'demo0' if count else None, 'ghost' if count else 'disabled', size=12)
    button(s, ML + 1.5, 5.42, 1.7, 0.46, 'Continue  ►', 'logic', 'ghost', size=12)
    hint = 'Each click is one press of the turn button.' if count < 5 else 'Five pages. Click Start over to begin again.'
    text(s, ML, 6.15, 4.6, 0.4, [P(hint, size=11, color=C['muted'])])


demo_slide('demo0', 0, IMG['demo0'], 'The open book with the Book Counter and its page turner above it, count at zero',
           'Ready. Press the turn button.', False, None)
for k in range(1, 6):
    demo_slide(f'fwd{k}', k, f'{GIFS}/fwd{k}.gif', f'The page turner turning a page forward; the count goes from {k - 1} to {k}',
               'Page turner: right sensor, then left:  +1', True, 1)
for k in range(4, -1, -1):
    demo_slide(f'back{k}', k, f'{GIFS}/back{k}.gif', f'A page flipped back by hand; the count goes from {k + 1} to {k}',
               'By hand: left sensor, then right:  −1', True, -1)

# 35 --------------------------------------------------------------- logic
s = new_slide('logic', notes=(
    "This is the heart of our code, simplified. Every time a sensor sees a page, it calls hit with R or L. "
    "The first sensor to fire just arms and waits. When the other sensor fires, the page has finished turning: "
    "if the right one was first, we add a page, otherwise we take one away. The first line is the lock-out: "
    "if we counted less than 400 milliseconds ago, we ignore the hit."))
full_bleed(s, IMG['oled'], 'Top view of the device with the OLED reading 027 pages', dx=0.35, crop_bottom=0.14)
menu_pill(s)
eyebrow(s, 'Firmware')
title(s, 'The counting logic', w=6.0)
xp_window(s, ML, 2.2, 5.45, 3.05, 'counter.ino - Notepad', name='Code window')
code = rect(s, ML, 2.2, 5.45, 3.05, fill='FFFFFF', name='Code')
K, D, M = C['aquaInk'], C['navy'], '5C7A99'
CODE = [
    [('void', K), (' hit(', D), ('char', K), (' side) {', D)],
    [('  if', K), (' (millis() - lastCount < LOCKOUT_MS) ', D), ('return', K), (';', D)],
    [('  if', K), (" (armed && armed != side) {", D)],
    [("    pages += (armed == 'R') ? 1 : -1;", D)],
    [('    lastCount = millis();', D)],
    [('    armed = 0;', D)],
    [('    showCount(pages);', D)],
    [('  } ', D), ('else', K), (' {', D)],
    [('    armed = side;  ', D), ('// wait for the other sensor', M)],
    [('  }', D)],
    [('}', D)],
]
fill_tf(code.text_frame, [P(runs=[dict(t=t, color=c) for t, c in ln], size=11.5, font=MONO, line=1.12) for ln in CODE],
        margins=(0.2, 0.18, 0.1, 0.12))
text(s, ML, 5.55, 4.9, 0.6, [P('Simplified. The first sensor arms; the second one counts. The turn button only starts the servo: '
                              'the page it turns is counted by this same code.', size=11, color=C['muted'], line=1.1)])

# 36 --------------------------------------------------------------- wiring
s = new_slide('wiring', notes=(
    "Here's how it's wired. Each sensor has three pins: power, ground and an output. The outputs go to GPIO 32 "
    "and 33 on the ESP32. The OLED screen uses I²C, which needs only two data wires, on GPIO 21 and 22. "
    "The turn button is on GPIO 4. The ESP32 sends the servo signal on GPIO 25 to the motor driver, and the servo cable runs through "
    "the hollow beam to the driver. Two 18650 lithium cells power the motor driver, and the driver's 5 volt output powers the ESP32. The USB cable is only for uploading the code."))
full_bleed(s, IMG['wiring'], 'Top view of the opened shell with the ESP32 DevKit, OLED, turn button, motor driver, battery pack and USB cable labelled', dx=0.95)
menu_pill(s)
eyebrow(s, 'Circuit')
title(s, 'Wiring')
ROWS = [('Part', 'Pin', 'Goes to'), ('IR sensor, left', 'OUT', 'GPIO 32'), ('IR sensor, right', 'OUT', 'GPIO 33'),
        ('Both IR sensors', 'VCC · GND', 'VIN · GND'), ('OLED display', 'SDA · SCL', 'GPIO 21 · 22'),
        ('OLED display', 'VCC · GND', '3V3 · GND'), ('Turn button', 'leg 1 · leg 2', 'GPIO 4 · GND'),
        ('Motor driver', 'signal in', 'GPIO 25'), ('Battery pack', '+ · −', 'motor driver in'),
        ('Motor driver', '5V out · GND', 'ESP32 VIN · GND'), ('Servo (in the beam)', 'cable', 'motor driver'),
        ('Computer', 'USB cable', 'uploads the code')]
glass(s, ML - 0.1, 1.85, 5.5, 0.33 * len(ROWS) + 0.2, opacity=0.7, rad=0.14, name='Table frame', gloss=0)
gf = s.shapes.add_table(len(ROWS), 3, Inches(ML), Inches(1.95), Inches(5.3), Inches(0.33 * len(ROWS)))
gf.name = 'Wiring table'
tbl = gf.table
tbl.first_row = True
tbl.horz_banding = False
for ci, wdt in enumerate((2.2, 1.4, 1.7)):
    tbl.columns[ci].width = Inches(wdt)
for ri, row in enumerate(ROWS):
    tbl.rows[ri].height = Inches(0.33)
    for ci, val in enumerate(row):
        cell = tbl.cell(ri, ci)
        tcPr = cell._tc.get_or_add_tcPr()
        for edge in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
            ln = etree.SubElement(tcPr, qn(edge))
            if edge == 'a:lnB' and ri:
                ln.set('w', '9525')
                sf = etree.SubElement(ln, qn('a:solidFill'))
                etree.SubElement(sf, qn('a:srgbClr')).set('val', 'C8DCF0')
            else:
                ln.set('w', '0')
                etree.SubElement(ln, qn('a:noFill'))
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb('FFFFFF' if ri % 2 else 'EAF5FF')
        if ri == 0:
            _swap_fill(tcPr, grad_el([(0, C['xpHi']), (100, C['xp'])], 90))
        cell.margin_left = cell.margin_right = Inches(0.1)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        head = ri == 0
        color = 'FFFFFF' if head else (C['aquaInk'] if ci == 2 else C['body'])
        fill_tf(cell.text_frame, [P(val.upper() if head else val, size=9.5 if head else 11.5, bold=head or ci == 2,
                                    font=EYE if head else (MONO if ci == 2 else BODY), color=color,
                                    spacing=120 if head else None)], anchor='middle', margins=(0.1, 0.02, 0.06, 0.02))
text(s, ML, 6.42, 5.2, 0.6, [P('Two 18650 lithium cells power everything. The USB cable from the computer uploads the code.', size=11, color=C['muted'], line=1.1)])

# 37 --------------------------------------------------------------- build
s = new_slide('build', notes=(
    "We designed the shell around the parts and 3D-printed it in white PLA. It's about 105 by 60 by 35 millimetres, "
    "in two pieces, a base and a top plate, so the electronics drop in from the top. The top plate has cut-outs for the "
    "screen, the button and the sensor cables, and the name is embossed right into the front wall. "
    "We also printed a hollow beam, 180 millimetres long, that holds the page turner and carries its cable. "
    "And the whole bookmark sits on a sheet of corrugated cardboard, which gives it one flat base."))
full_bleed(s, IMG['dims'], 'The shell with dimension lines: 105 mm wide, 60 mm deep, 35 mm tall', dx=0.75)
menu_pill(s)
eyebrow(s, 'Build')
title(s, 'Designed and 3D-printed')
FACTS = [('Size', '105 × 60 × 35 mm', True), ('Material', 'PLA, white', False),
         ('Pieces', 'Base tray, top plate and hollow beam', False),
         ('Hollow beam', '180 × 16 × 13 mm, hollow', True),
         ('Openings', 'Screen window, button hole, two cable ports', False),
         ('Lettering', 'BOOK COUNTER, embossed on the front', False),
         ('Base', 'Corrugated cardboard sheet under the whole bookmark', False)]
FY, FS = 2.4, 0.6
for i, (k, v, mono) in enumerate(FACTS):
    y = FY + i * FS
    line(s, ML, y, ML + 4.6, y, color=C['rule'], width=0.75)
    text(s, ML, y + 0.14, 1.45, 0.4, [P(k.upper(), size=9.5, bold=True, font=EYE, color=C['muted'], spacing=100)])
    text(s, ML + 1.5, y + 0.1, 3.1, 0.5, [P(v, size=12 if not mono else 13, bold=mono, font=MONO if mono else BODY,
                                             color=C['aquaInk'] if mono else C['navy'], line=1.0)])
line(s, ML, FY + len(FACTS) * FS, ML + 4.6, FY + len(FACTS) * FS, color=C['rule'], width=0.75)

# 38 --------------------------------------------------------------- challenges
s = new_slide('challenges', notes=(
    "A page counter sounds simple, but a few things have to be right. One page could trigger a sensor more than once, "
    "so we added a short lock-out. Flipping back shouldn't add pages, which is why we use two sensors instead of one. "
    "Infrared also reflects off other things, so we tuned each sensor's range with its trimmer. And the page turner "
    "must move exactly one page, so the servo only moves briefly and the soft wheel grips just the top page."))
glass(s, 0.42, 0.42, 12.5, 6.66, opacity=0.5, name='Challenges panel', gloss=0.4)
menu_pill(s)
eyebrow(s, 'Design challenges')
title(s, 'What a page counter has to get right', w=11.0, h=0.8, size=34)
for i, (q, a) in enumerate([('One page could count twice.', 'A 400 ms lock-out after every count.'),
                            ('Flipping back must not add pages.', 'Two sensors, so the order of the hits gives the direction.'),
                            ('Room light and the page below reflect IR too.', "Trimmers tuned for a short range, aimed at the page's path."),
                            ('The page turner must move exactly one page.', 'A short servo movement and a soft rubber wheel that grips only the top page.')]):
    col, row = i % 2, i // 2
    card(s, ML + col * 6.0, 2.15 + row * 2.3, 5.75, 2.05,
         [P(q, size=18, bold=True, font=HEAD, color=C['navy'], after=10, line=1.0),
          P(runs=[dict(t='FIX  ', bold=True, font=EYE, color=C['limeInk'], size=11, spacing=120), dict(t=a)], size=14, line=1.12)],
         margins=(0.32, 0.28, 0.3, 0.2), name=f'Challenge {i + 1}', opacity=0.82, rad=0.16)

# 39 --------------------------------------------------------------- next
s = new_slide('next', notes=(
    "There's a lot we'd like to add next. Bluetooth, so each reading session goes to a phone app with graphs. "
    "Reading speed in pages per minute. A battery meter on the screen, so you know when to charge. Letting the page turner "
    "flip pages back as well, and an adjustable clip so it fits any book."))
glass(s, 0.42, 0.42, 12.5, 6.66, opacity=0.55, name='Next panel', gloss=0.4)
menu_pill(s)
eyebrow(s, 'Next version')
title(s, 'Where it goes next', w=10)
for i, (k, v) in enumerate([('Bluetooth sync', 'Send each session to a phone app with daily and weekly graphs.'),
                            ('Reading speed', 'Show pages per minute right on the screen.'),
                            ('Battery meter', 'Show how much charge is left, right on the OLED.'),
                            ('Turn back too', 'Spin the wheel the other way to flip a page back.'),
                            ('Fits any book', 'An adjustable clip for thin paperbacks and thick textbooks.')]):
    y = 2.15 + i * 0.92
    line(s, ML, y, W_IN - ML, y, color=C['rule'], width=0.75)
    orb(s, ML + 0.17, y + 0.46, 0.22, color=[C['aqua'], C['pink'], C['lime'], C['tangerine'], C['lilac']][i], solid=True,
        name=f'Bullet {i + 1}')
    text(s, ML + 0.5, y + 0.2, 3.2, 0.6, [P(k, size=21, bold=True, font=HEAD, color=C['aquaInk'])])
    text(s, ML + 3.75, y + 0.26, 8.0, 0.6, [P(v, size=15, color=C['body'])])
line(s, ML, 2.15 + 5 * 0.92, W_IN - ML, 2.15 + 5 * 0.92, color=C['rule'], width=0.75)

# 40 --------------------------------------------------------------- thanks
s = new_slide('thanks', notes=(
    "Thank you for listening. Book Counter counts the pages so you don't have to. We're happy to take questions, "
    "and the menu button lets us jump back to any part you'd like to see again."))
orb(s, 12.55, 6.55, 1.25, color=C['aqua'])
orb(s, 6.05, 6.75, 0.55, color=C['pink'])
orb(s, 12.75, 0.7, 0.6, color=C['lilac'])
window_media(s, f'{GIFS}/spin.gif', RX, 1.85, RW, 'spin.gif', 'The Book Counter turning on a turntable', name='Spinning model')
left_panel(s, y=1.45, h=4.75)
eyebrow(s, 'Thank you', y=1.85)
text(s, ML, 2.25, 5.0, 1.2, [P('Questions?', size=58, bold=True, font=HEAD, color=C['navy'])])
sparkle(s, 5.5, 2.35, 24, C['pink'])
text(s, ML, 3.55, 4.7, 0.9, [P("Book Counter counts the pages so you don't have to.", size=16, color=C['body'], line=1.1)])
text(s, ML, 4.5, 4.8, 0.3, [P('[Presenter names · Class]', size=11, bold=True, font=EYE, color=C['muted'])])
button(s, ML, 5.25, 1.9, 0.5, 'Menu', 'menu', 'primary', size=14)
button(s, ML + 2.05, 5.25, 1.9, 0.5, 'Start again', 'title', 'secondary', size=14)

# ================================================================ wiring the buttons
for shp, key in LINKS:
    shp.click_action.target_slide = SLIDES[key]

# ================================================================ transitions
for key in SLIDES:
    if key == 'hub' or key.startswith('det_'):
        X.set_transition(SLIDES[key], 'morph', 1100)
    elif key.startswith('tour'):
        X.set_transition(SLIDES[key], 'morph', 1600)
    elif key.startswith('fwd') or key.startswith('back'):
        X.set_transition(SLIDES[key], 'fade', 200)
    elif key in ('idea', 'quiz_right'):
        X.set_transition(SLIDES[key], 'curl', 1200)
    else:
        X.set_transition(SLIDES[key], 'fade', 500)

prs.core_properties.title = 'Book Counter'
prs.core_properties.subject = 'Interactive presentation of a 3D-printed page-counting bookmark'
prs.save(OUT)
X.finalize(OUT)  # package-level additions (3D model parts, content types)
hidden = sum(1 for s in prs.slides if s._element.get('show') == '0')
print(f'wrote {OUT}: {len(prs.slides)} slides ({hidden} hidden, reached by buttons), {len(LINKS)} buttons')
