#!/usr/bin/env python3
"""Builds BookCounter-interactive.pptx.

Interactivity is all native PowerPoint, so it works in a normal slide show:
  * buttons and hotspots that jump to other slides (a:hlinkClick ppaction://hlinksldjump)
  * hidden slides that only open from those buttons (part close-ups, quiz answers, demo steps)
  * Morph zoom between the device and each part, and Morph on the embedded 3D model
  * a real 3D model (glTF) PowerPoint 365 can rotate, with a picture fallback elsewhere
  * animated GIFs: a turntable loop, the shell opening, the IR beam, and one page turn per demo step

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

# ---------------------------------------------------------------- look
C = dict(bg='0C1118', bg2='131B26', panel='0F1620', line='2A3647', text='ECE6D9', body='C9CFD8',
         muted='9BA5B5', dim='5E6A7C', accent='86D6FF', ir='FF6479', ink='0C1118', oled='05080C')
HEAD, BODY, MONO = 'Arial', 'Cambria', 'Courier New'
W_IN, H_IN = 13.333, 7.5
ML = 0.75  # left margin

prs = Presentation()
prs.slide_width = Inches(W_IN)
prs.slide_height = Inches(H_IN)
BLANK = prs.slide_layouts[6]
SLIDES = {}
LINKS = []  # (shape, target key) wired once every slide exists


def rgb(h):
    return RGBColor.from_string(h)


def set_alpha(fill_parent_el, opacity):
    """opacity 0..1 on the first a:srgbClr under a fill element"""
    clr = fill_parent_el.find('.//' + qn('a:srgbClr'))
    a = etree.SubElement(clr, qn('a:alpha'))
    a.set('val', str(int(opacity * 100000)))


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
    out = os.path.join(WORK, 'scrim.png')
    w, h = 1600, 90
    im = Image.new('RGBA', (w, h))
    px = im.load()
    for x in range(w):
        t = x / w
        a = 242 if t < 0.36 else max(0, int(242 * (1 - (t - 0.36) / 0.26)))
        for y in range(h):
            px[x, y] = (12, 17, 24, a)
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
    th_build=crop_square(f'{ASSETS}/dims.png', 0.68, 0.5, 700, 'th_build.jpg'),
)
SPOTS = json.load(open(f'{ASSETS}/hub_spots.json'))
SPOTS['sensorL'] = [0.545, 0.68]   # centre of the board rather than the LED tips
SPOTS['sensorR'] = [0.893, 0.578]


# ---------------------------------------------------------------- building blocks
def new_slide(key, bg=C['bg'], hidden=False, notes=''):
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = rgb(bg)
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
            if p['spacing']:
                f._rPr.set('spc', str(p['spacing']))


def text(s, x, y, w, h, paras, anchor='top', name=None):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    fill_tf(tb.text_frame, paras, anchor)
    if name:
        tb.name = name
    return tb


def rect(s, x, y, w, h, fill=None, line=None, radius=None, opacity=None, name=None, line_w=1.0):
    shape = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
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


BTN = dict(
    primary=dict(fill=C['accent'], line=None, color=C['ink']),
    secondary=dict(fill=C['bg2'], line='3B4A60', color=C['text']),
    ghost=dict(fill=None, line=C['line'], color=C['muted']),
    disabled=dict(fill=C['bg2'], line=C['line'], color=C['dim']),
    dark=dict(fill=C['ink'], line=None, color=C['accent']),
    darkline=dict(fill=None, line=C['ink'], color=C['ink']),
    menu=dict(fill=C['ink'], line=C['line'], color=C['muted']),
)


def button(s, x, y, w, h, label, target=None, style='secondary', size=14, name=None):
    st = BTN[style]
    b = rect(s, x, y, w, h, fill=st['fill'], line=st['line'], radius=0.28 if h < 0.9 else 0.12, name=name)
    fill_tf(b.text_frame, [P(label, size=size, bold=True, font=HEAD, color=st['color'], align='center')],
            anchor='middle', margins=(0.08, 0.02, 0.08, 0.02))
    if target and style != 'disabled':
        LINKS.append((b, target))
    return b


def picture(s, path, x, y, w=None, h=None, name=None, alt='', target=None):
    p = s.shapes.add_picture(path, Inches(x), Inches(y), Inches(w) if w else None, Inches(h) if h else None)
    if name:
        p.name = name
    p._element.nvPicPr.cNvPr.set('descr', alt)
    if target:
        LINKS.append((p, target))
    return p


def line(s, x1, y1, x2, y2, color=C['accent'], width=1.25):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = rgb(color)
    c.line.width = Pt(width)
    return c


def eyebrow(s, t, x=ML, y=0.72, color=C['accent'], w=7):
    return text(s, x, y, w, 0.3, [P(t.upper(), size=12, bold=True, font=HEAD, color=color, spacing=200)])


def title(s, t, x=ML, y=1.05, w=5.4, h=1.4, size=38, color=C['text']):
    return text(s, x, y, w, h, [P(t, size=size, bold=True, font=HEAD, color=color, line=0.92)])


def menu_pill(s, style='menu'):
    return button(s, 11.83, 0.42, 0.9, 0.36, 'Menu', 'menu', style=style, size=11, name='Menu button')


def full_bleed(s, path, alt, name=None, with_scrim=True, dx=0.0, crop_bottom=0.0):
    p = picture(s, path, dx, 0, W_IN, H_IN * (1 - crop_bottom), name=name, alt=alt)
    if crop_bottom:
        p.crop_bottom = crop_bottom
    if dx:  # keep the picture inside the slide: trim what would hang off the right edge
        p.crop_right = dx / W_IN
        p.width = Inches(W_IN - dx)
    if with_scrim:
        picture(s, IMG['scrim'], 0, 0, W_IN, H_IN, name='Scrim', alt='')


def gif(s, path, x, y, w, alt, name=None):
    h = w * 9 / 16
    return picture(s, path, x, y, w, h, name=name, alt=alt)


# ================================================================ slides
# 1 ---------------------------------------------------------------- title
s = new_slide('title', notes=(
    "Hello everyone. This is Book Counter, a bookmark we designed and 3D-printed that counts the pages you turn. "
    "This presentation is interactive: the buttons work during the slide show. Click Start, or press the right arrow."))
gif(s, f'{GIFS}/spin.gif', 5.45, 1.2, 7.75, 'The Book Counter turning on a turntable', name='Spinning model')
eyebrow(s, '3D-printed smart bookmark', y=1.55)
text(s, ML, 1.9, 5.2, 1.9, [P('BOOK', size=60, bold=True, font=HEAD, color=C['text'], line=0.88),
                            P('COUNTER', size=60, bold=True, font=HEAD, color=C['text'], line=0.88)])
text(s, ML, 3.85, 4.7, 0.9, [P('A bookmark that counts every page you turn and shows the total on its own screen.',
                               size=18, color=C['body'], line=1.1)])
for i, (num, lab) in enumerate([('2', 'IR sensors'), ('0.96"', 'OLED screen'), ('1', 'printed shell')]):
    chip = rect(s, ML + i * 1.55, 4.95, 1.42, 0.82, fill=None, line=C['line'], radius=0.12)
    fill_tf(chip.text_frame, [P(num, size=18, bold=True, font=MONO, color=C['accent']),
                              P(lab, size=11, bold=True, font=HEAD, color=C['muted'])], anchor='middle',
            margins=(0.14, 0.04, 0.08, 0.04))
text(s, ML, 5.98, 4.8, 0.3, [P('[Presenter names · Class]', size=12, bold=True, font=HEAD, color=C['muted'])])
button(s, ML, 6.4, 1.55, 0.5, 'Start  ►', 'problem', 'primary', size=14)
button(s, ML + 1.7, 6.4, 1.2, 0.5, 'Menu', 'menu', 'secondary', size=14)

# 2 ---------------------------------------------------------------- menu
s = new_slide('menu', notes=(
    "This is the menu. Every tile is a button: click one to jump straight to that part of the talk. "
    "If you just press the right arrow, the presentation goes through everything in order."))
eyebrow(s, 'Explore')
title(s, 'Where do you want to go?', w=11)
TILES = [
    ('START HERE', 'The problem', 'Why paper books need a counter', 'problem', 'th_problem'),
    ('CLICK THE PARTS', 'The device', 'Explore it piece by piece', 'hub', 'th_parts'),
    ('REAL 3D', '3D tour', 'A 3D model you can turn', 'tour1', 'th_tour'),
    ('TRY IT', 'Live demo', 'Turn pages, watch it count', 'demo0', 'th_demo'),
    ('AUDIENCE', 'Quick quiz', 'One question about the sensors', 'quiz', 'th_quiz'),
    ('HOW IT IS MADE', 'Build', 'Wiring and the printed shell', 'wiring', 'th_build'),
]
for i, (lab, name, desc, target, th) in enumerate(TILES):
    col, row = i % 3, i // 3
    x, y, w, h = ML + col * 4.05, 2.3 + row * 2.3, 3.8, 2.05
    tile = rect(s, x, y, w, h, fill=C['bg2'], line=C['line'], radius=0.08, name=f'Tile: {name}')
    fill_tf(tile.text_frame, [P(lab, size=10, bold=True, font=HEAD, color=C['accent'], spacing=150, after=6),
                              P(name, size=20, bold=True, font=HEAD, color=C['text'], after=4),
                              P(desc, size=12, color=C['muted'], line=1.05)],
            anchor='top', margins=(0.24, 0.24, 1.9, 0.2))
    LINKS.append((tile, target))
    picture(s, IMG[th], x + w - 1.72, y + 0.3, 1.45, 1.45, name=f'Thumb: {name}', alt=name, target=target)
text(s, ML, 6.95, 9, 0.3, [P('Each tile is a button. Press → to go through everything in order.', size=12, color=C['muted'])])

# 3 ---------------------------------------------------------------- problem
s = new_slide('problem', notes=(
    "E-readers like a Kindle track exactly how much you read. Paper books don't. If you want a reading goal, "
    "like twenty pages a day, you have to remember where you started and do the maths yourself, and most people forget. "
    "Reading apps help, but you still have to type your pages in every time."))
full_bleed(s, IMG['problem'], 'An open book on a desk with the Book Counter above its top edge')
menu_pill(s)
eyebrow(s, 'Why we built it')
title(s, 'Paper books keep no score')
text(s, ML, 2.55, 4.3, 0.8, [P("An e-reader knows how far you got. A paper book doesn't.", size=18, color=C['body'], line=1.1)])
text(s, ML, 3.55, 4.3, 2.4, [
    P(runs=[dict(t='•  ', color=C['accent']), dict(t='Reading goals need numbers: pages a day, pages a week.')], size=16, after=10, line=1.1),
    P(runs=[dict(t='•  ', color=C['accent']), dict(t='Page numbers written down by hand get forgotten.')], size=16, after=10, line=1.1),
    P(runs=[dict(t='•  ', color=C['accent']), dict(t='Reading apps make you type in every session.')], size=16, line=1.1)])

# 4 ---------------------------------------------------------------- idea
s = new_slide('idea', bg=C['accent'], notes=(
    "So our idea was simple: make the bookmark do the counting. You put it on your book, read normally, "
    "and it counts every page you turn by itself."))
menu_pill(s, 'darkline')
eyebrow(s, 'The idea', color=C['ink'], y=1.6)
text(s, ML, 2.05, 11.6, 2.4, [P('What if the bookmark kept count for you?', size=56, bold=True, font=HEAD, color=C['ink'], line=0.95)])
text(s, ML, 4.6, 9.5, 1.0, [P('Book Counter sits at the top of the book and counts each page as you turn it. '
                              'No buttons to press, no app, nothing to write down.', size=20, color='1B2A3A', line=1.15)])
button(s, ML, 5.95, 2.6, 0.55, 'Meet the device  ►', 'hub', 'dark', size=14)

# 5 ---------------------------------------------------------------- hub with hotspots
PARTS = [
    ('oled', 'OLED display', '0.96" OLED · 128 × 64 px · SSD1306 · I²C',
     'Shows the running page count. It needs only two data wires, SDA and SCL, to the microcontroller.',
     (7.0, 2.1, 1.6), 'This is the screen. It shows the count, and it only needs two data wires to talk to the Arduino.'),
    ('button', 'Reset button', '6 mm push button',
     'Press it to set the count back to zero at the start of a reading session.',
     (9.75, 1.85, 1.6), 'The reset button sets the count back to zero when you start a new reading session.'),
    ('shell', '3D-printed shell', 'PLA · 105 × 60 × 35 mm · two pieces',
     'A base tray and a top plate with cut-outs for the screen, the button and the two sensor cables. '
     'BOOK COUNTER is embossed on the front.',
     (8.65, 5.9, 1.95), 'We designed this shell and printed it in two pieces, with the name embossed on the front.'),
    ('sensorL', 'IR sensor, left', 'IR obstacle module (FC-51) · LM393',
     'Watches the left side of the book. When you read forward, it sees the page second, as the page lands.',
     (5.4, 6.3, 1.75), 'The left sensor sees the page second when you turn forward, as the page lands on the left.'),
    ('sensorR', 'IR sensor, right', 'IR obstacle module (FC-51) · LM393',
     'Watches the right side of the book. When you read forward, it sees the page first, as you lift it.',
     (10.95, 6.2, 1.75), 'The right sensor sees the page first when you turn forward, as you lift it off the right side.'),
]
s = new_slide('hub', notes=(
    "Here's the device. Everything you can see from outside has one job. "
    "Click any label, or any blue dot, to zoom in on that part. Each close-up has a Back button."))
picture(s, IMG['hub'], 0, 0, W_IN, H_IN, name='!!device', alt='The Book Counter: white printed box, OLED window, reset button and two blue IR sensor boards')
menu_pill(s)
eyebrow(s, 'The device')
title(s, 'Two eyes, one screen')
text(s, ML, 2.55, 4.6, 1.2, [P('Click any part of the model to look closer.', size=18, color=C['body'], line=1.1, after=8),
                             P('Every label and blue dot is a button.', size=14, color=C['muted'])])
for key, name, spec, desc, (px, py, pw), _ in PARTS:
    u, v = SPOTS[key]
    dx, dy = u * W_IN, v * H_IN
    # leader from the dot to the nearest edge of its label
    lx = px + pw if px + pw < dx else (px if px > dx else dx)
    ly = py + 0.4 if py + 0.4 < dy else (py if py > dy else py + 0.2)
    line(s, dx, dy, lx, ly)
    halo = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(dx - 0.2), Inches(dy - 0.2), Inches(0.4), Inches(0.4))
    halo.fill.solid(); halo.fill.fore_color.rgb = rgb(C['accent']); set_alpha(halo._element.spPr.find(qn('a:solidFill')), 0.28)
    halo.line.fill.background(); halo.shadow.inherit = False; halo.name = f'Hotspot halo: {name}'
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(dx - 0.1), Inches(dy - 0.1), Inches(0.2), Inches(0.2))
    dot.fill.solid(); dot.fill.fore_color.rgb = rgb(C['accent'])
    dot.line.color.rgb = rgb(C['ink']); dot.line.width = Pt(1.5); dot.shadow.inherit = False; dot.name = f'Hotspot: {name}'
    pill = rect(s, px, py, pw, 0.4, fill=C['ink'], line=C['accent'], radius=0.3, opacity=0.92, name=f'Label: {name}')
    fill_tf(pill.text_frame, [P(name, size=12, bold=True, font=HEAD, color=C['text'], align='center')], anchor='middle',
            margins=(0.06, 0.02, 0.06, 0.02))
    for shp in (halo, dot, pill):
        LINKS.append((shp, f'det_{key}'))

# 6-10 ------------------------------------------------------------- part close-ups (hidden, Morph zoom)
Z = 2.4
for i, (key, name, spec, desc, _, say) in enumerate(PARTS):
    s = new_slide(f'det_{key}', hidden=True, notes=say + ' Click Back to return to the whole device, or Next part.')
    u, v = SPOTS[key]
    iw, ih = W_IN * Z, H_IN * Z
    left = min(0.0, max(W_IN - iw, 8.9 - u * iw))
    top = min(0.0, max(H_IN - ih, 3.9 - v * ih))
    picture(s, IMG['hub'], left, top, iw, ih, name='!!device', alt=f'Close-up of the {name}')
    rect(s, 0, 0, 5.7, H_IN, fill=C['bg'], opacity=0.9, name='Detail panel')
    menu_pill(s)
    eyebrow(s, f'Part {i + 1} of {len(PARTS)}')
    text(s, ML, 1.05, 4.6, 1.5, [P(name, size=36, bold=True, font=HEAD, color=C['text'], line=0.95)])
    text(s, ML, 2.6, 4.6, 0.7, [P(spec, size=13, bold=True, font=MONO, color=C['accent'], line=1.1)])
    text(s, ML, 3.35, 4.5, 2.0, [P(desc, size=17, color=C['body'], line=1.15)])
    button(s, ML, 5.75, 2.55, 0.52, '◄  Back to the device', 'hub', 'primary', size=13)
    nxt = PARTS[(i + 1) % len(PARTS)][0]
    button(s, ML + 2.7, 5.75, 1.75, 0.52, 'Next part  ►', f'det_{nxt}', 'secondary', size=13)

# 11-14 ------------------------------------------------------------ 3D tour (real 3D model, Morph between views)
TOUR = [
    ('tour1', 'A real 3D model', 'This is the actual 3D file of the device. Press → and PowerPoint turns it for you.',
     'In edit mode, click the model and drag the round arrow in its middle to spin it any way you like.', 'front'),
    ('tour2', 'The sensor arms', 'An IR board hangs off each front corner. Both look down towards the pages.',
     'The jumper wires go into the shell through two ports in the top plate.', 'side'),
    ('tour3', 'Screen on top', 'The OLED window and the reset button sit on the top plate, where you can see them while reading.',
     'The screen shows the count in big pixel digits.', 'top'),
    ('tour4', 'All the way round', 'From behind you can see the power switch and the two cable ports.',
     'Press → to continue, or click Menu to jump somewhere else.', 'back'),
]
for i, (key, ttl, line1, line2, view) in enumerate(TOUR):
    s = new_slide(key, notes=(
        f"{line1} {line2} This model is the same 3D file we used in the live 3D version. "
        "Between these slides PowerPoint's Morph transition turns it smoothly."))
    X.add_model3d(prs, s, GLB, view, WORK, ASSETS, x=4.9, y=0.55, w=8.1, h=6.4, name='!!model')
    menu_pill(s)
    eyebrow(s, f'3D tour · {i + 1} of {len(TOUR)}')
    title(s, ttl, w=4.4)
    text(s, ML, 2.55, 4.0, 1.6, [P(line1, size=18, color=C['body'], line=1.12)])
    text(s, ML, 4.25, 4.0, 1.2, [P(line2, size=13, color=C['muted'], line=1.12)])
    if i < len(TOUR) - 1:
        button(s, ML, 5.9, 1.9, 0.5, 'Turn it  ►', TOUR[i + 1][0], 'primary', size=13)
    else:
        button(s, ML, 5.9, 1.9, 0.5, 'Inside  ►', 'inside', 'primary', size=13)
    if i > 0:
        button(s, ML + 2.05, 5.9, 1.4, 0.5, '◄  Back', TOUR[i - 1][0], 'secondary', size=13)

# 15 --------------------------------------------------------------- inside
s = new_slide('inside', notes=(
    "If we lift the top plate off, you can see the inside. The microcontroller is an Arduino Nano: it reads the sensors "
    "and keeps the count. The OLED screen is mounted under the window in the top plate. A battery powers everything, "
    "with a switch to turn it off. The sensor wires come in through two holes in the top."))
gif(s, f'{GIFS}/explode.gif', 5.55, 1.3, 7.4, 'The shell opening up to show the parts inside, then closing again', name='Exploded view')
menu_pill(s)
eyebrow(s, 'Inside the shell')
title(s, "What's inside")
text(s, ML, 2.0, 4.8, 0.8, [P('Lift off the top plate and four parts do the real work.', size=17, color=C['body'], line=1.1)])
for i, (nm, ds) in enumerate([('Microcontroller', 'Arduino Nano. Reads, decides, counts.'),
                              ('OLED display', '0.96", under the window.'),
                              ('Battery', '9 V, with a power switch.'),
                              ('Jumper wires', 'Bring both sensor signals in.')]):
    col, row = i % 2, i // 2
    card = rect(s, ML + col * 2.48, 3.0 + row * 1.32, 2.36, 1.18, fill=C['bg2'], line=C['line'], radius=0.1)
    fill_tf(card.text_frame, [P(nm, size=15, bold=True, font=HEAD, color=C['text'], after=4),
                              P(ds, size=12, color=C['muted'], line=1.05)], margins=(0.18, 0.16, 0.14, 0.1))
button(s, ML, 5.85, 2.3, 0.5, 'Click the parts  ►', 'hub', 'secondary', size=13)

# 16 --------------------------------------------------------------- sensor
s = new_slide('sensor', notes=(
    "Each sensor is an infrared module. The clear LED sends out infrared light, which we can't see; the red cone just "
    "shows where it goes. When a page moves into the beam, it reflects the light back to the dark photodiode. "
    "The LM393 chip turns that into a clean ON or OFF signal for the Arduino, and a small green LED lights up. "
    "The blue screw is a trimmer: we turn it to set how close a page has to be before it counts."))
gif(s, f'{GIFS}/sensor.gif', 5.55, 1.3, 7.4, 'A page lifting into the right sensor beam, which turns brighter', name='Sensor view')
menu_pill(s)
eyebrow(s, 'Sensing')
title(s, 'Seeing a page')
for i, (b, t) in enumerate([('Send.', 'The clear LED shines infrared light.'),
                            ('Reflect.', 'A page in the beam bounces it back.'),
                            ('Receive.', 'The dark photodiode picks it up.'),
                            ('Decide.', 'The LM393 chip outputs a clean ON or OFF.')]):
    y = 2.05 + i * 0.78
    badge = rect(s, ML, y, 0.46, 0.46, fill=C['accent'], radius=0.18)
    fill_tf(badge.text_frame, [P(str(i + 1), size=16, bold=True, font=MONO, color=C['ink'], align='center')], anchor='middle')
    text(s, ML + 0.65, y + 0.02, 4.0, 0.7, [P(runs=[dict(t=b + ' ', bold=True, color=C['text']), dict(t=t)], size=16, line=1.05)])
text(s, ML, 5.2, 4.6, 0.7, [P('The blue trimmer sets how close a page must come. Real IR is invisible; red shows the beam.', size=12, color=C['muted'], line=1.1)])
button(s, ML, 6.05, 2.3, 0.5, 'Quick quiz  ►', 'quiz', 'primary', size=13)

# 17 --------------------------------------------------------------- quiz
s = new_slide('quiz', bg=C['bg2'], notes=(
    "Quick question for the audience: when you turn a page forward, which sensor sees it first? "
    "Let someone answer, then click the button they chose."))
menu_pill(s)
eyebrow(s, 'Quick check')
text(s, ML, 1.05, 11.6, 1.6, [P('You turn a page forward. Which sensor sees it first?', size=38, bold=True, font=HEAD, color=C['text'], line=0.95)])
for i, (lab, sub, target) in enumerate([('Left sensor', 'The board on the left corner', 'quiz_wrong'),
                                        ('Right sensor', 'The board on the right corner', 'quiz_right')]):
    t = rect(s, ML + i * 6.0, 3.15, 5.75, 2.2, fill=C['bg'], line='3B4A60', radius=0.08, name=f'Answer: {lab}', line_w=1.5)
    fill_tf(t.text_frame, [P(lab, size=30, bold=True, font=HEAD, color=C['text'], align='center', after=6),
                           P(sub, size=14, color=C['muted'], align='center')], anchor='middle')
    LINKS.append((t, target))
text(s, ML, 5.75, 9, 0.4, [P('Click an answer.', size=14, color=C['muted'])])

s = new_slide('quiz_wrong', bg=C['bg2'], hidden=True, notes=(
    "Not quite. Ask them to picture the page: it starts on the right-hand side of the open book. Then click Try again."))
menu_pill(s)
eyebrow(s, 'Not quite')
text(s, ML, 1.05, 11.5, 1.2, [P('Think about where the page starts', size=40, bold=True, font=HEAD, color=C['text'])])
text(s, ML, 2.45, 9.5, 1.4, [P('You read the right-hand page, then lift it from the right side. Which sensor does it pass first?',
                               size=20, color=C['body'], line=1.15)])
button(s, ML, 4.2, 2.0, 0.56, '◄  Try again', 'quiz', 'primary', size=14)

s = new_slide('quiz_right', bg=C['accent'], hidden=True, notes=(
    "Correct. The page starts on the right, so as you lift it, it passes the right sensor first and then the left one as "
    "it lands. Right then left means plus one. That order is how the bookmark knows which way you turned."))
menu_pill(s, 'darkline')
eyebrow(s, 'Correct', color=C['ink'], y=1.4)
text(s, ML, 1.85, 11.5, 1.2, [P('Right, then left', size=56, bold=True, font=HEAD, color=C['ink'])])
text(s, ML, 3.2, 9.8, 1.6, [P('A page you turn forward lifts off the right side, so the right sensor sees it first. '
                              'Then it lands on the left. Right then left means one more page read.', size=20, color='1B2A3A', line=1.15)])
button(s, ML, 5.3, 3.4, 0.56, 'Try it in the live demo  ►', 'demo0', 'dark', size=14)
button(s, ML + 3.6, 5.3, 2.4, 0.56, '◄  Back to the question', 'quiz', 'darkline', size=13)

# 20-30 ------------------------------------------------------------ live demo: one slide per count
def demo_slide(key, count, media, alt, status, hidden, last_dir):
    notes = {
        None: "This is the live demo. Click Turn page: each click plays one page turn and the count goes up. "
              "Flip back takes it down again. Reset goes back to zero. Continue moves on.",
        1: f"The page lifted off the right side first, then landed on the left. Right then left: plus one. The count is now {count}.",
        -1: f"This time the page went the other way: left sensor first, then right. That means minus one. The count is now {count}.",
    }[last_dir]
    s = new_slide(key, hidden=hidden, notes=notes)
    picture(s, media, 5.45, 1.3, 7.6, 7.6 * 9 / 16, name='Demo view', alt=alt)
    menu_pill(s)
    eyebrow(s, 'Live demo')
    title(s, 'Turn a page, watch it count', w=4.5)
    box = rect(s, ML, 2.55, 2.9, 1.25, fill=C['oled'], line=C['accent'], radius=0.1, name='Counter')
    fill_tf(box.text_frame, [P(f'{count:03d}', size=48, bold=True, font=MONO, color=C['accent'], align='center'),
                             P('PAGES TURNED', size=10, bold=True, font=HEAD, color=C['muted'], align='center', spacing=150)],
            anchor='middle')
    text(s, ML, 4.0, 4.6, 0.5, [P(status, size=15, color=C['body'])])
    nxt = f'fwd{count + 1}' if count < 5 else None
    prv = f'back{count - 1}' if count > 0 else None
    button(s, ML, 4.7, 2.15, 0.55, 'Turn page  ►', nxt, 'primary' if nxt else 'disabled', size=14)
    button(s, ML + 2.3, 4.7, 2.0, 0.55, '◄  Flip back', prv, 'secondary' if prv else 'disabled', size=14)
    button(s, ML, 5.42, 1.3, 0.46, 'Reset', 'demo0' if count else None, 'ghost' if count else 'disabled', size=12)
    button(s, ML + 1.45, 5.42, 1.7, 0.46, 'Continue  ►', 'logic', 'ghost', size=12)
    hint = 'Each click plays one page turn.' if count < 5 else 'Five pages. Press Reset to start again.'
    text(s, ML, 6.15, 4.6, 0.4, [P(hint, size=12, color=C['muted'])])


demo_slide('demo0', 0, IMG['demo0'], 'The open book with the Book Counter above it, count at zero',
           'Ready. Click Turn page.', False, None)
for k in range(1, 6):
    demo_slide(f'fwd{k}', k, f'{GIFS}/fwd{k}.gif', f'A page turning forward; the count goes from {k - 1} to {k}',
               'Right sensor, then left sensor:  +1', True, 1)
for k in range(4, -1, -1):
    demo_slide(f'back{k}', k, f'{GIFS}/back{k}.gif', f'A page turning back; the count goes from {k + 1} to {k}',
               'Left sensor, then right sensor:  −1', True, -1)

# 31 --------------------------------------------------------------- logic
s = new_slide('logic', notes=(
    "This is the heart of our code, simplified. Every time a sensor sees a page, it calls hit with R or L. "
    "The first sensor to fire just arms and waits. When the other sensor fires, the page has finished turning: "
    "if the right one was first, we add a page, otherwise we take one away. The first line is the lock-out: "
    "if we counted less than 400 milliseconds ago, we ignore the hit."))
full_bleed(s, IMG['oled'], 'Top view of the device with the OLED reading 027 pages', dx=0.35, crop_bottom=0.14)
menu_pill(s)
eyebrow(s, 'Firmware')
title(s, 'The counting logic', w=6.0)
code = rect(s, ML, 2.0, 5.45, 3.75, fill=C['bg2'], line=C['line'], radius=0.05, name='Code')
K, D, M = C['accent'], C['text'], '7F8B9D'
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
fill_tf(code.text_frame, [P(runs=[dict(t=t, color=c) for t, c in ln], size=12, font=MONO, line=1.18) for ln in CODE],
        margins=(0.25, 0.25, 0.15, 0.2))
text(s, ML, 5.95, 5.4, 0.4, [P('Simplified. The first sensor arms; the second one counts.', size=12, color=C['muted'])])

# 32 --------------------------------------------------------------- wiring
s = new_slide('wiring', notes=(
    "Here's how it's wired. Each sensor has three pins: power, ground and an output. The outputs go to digital pins D2 "
    "and D3 on the Arduino. The OLED screen uses I²C, which needs only two data wires, on A4 and A5. "
    "The reset button is on D4, and the battery goes through the power switch into VIN."))
full_bleed(s, IMG['wiring'], 'Top view of the opened shell with the Arduino Nano, battery and OLED labelled', dx=0.95)
menu_pill(s)
eyebrow(s, 'Circuit')
title(s, 'Wiring')
ROWS = [('Part', 'Pin', 'Arduino'), ('IR sensor, left', 'OUT', 'D2'), ('IR sensor, right', 'OUT', 'D3'),
        ('Both IR sensors', 'VCC · GND', '5V · GND'), ('OLED display', 'SDA · SCL', 'A4 · A5'),
        ('OLED display', 'VCC · GND', '5V · GND'), ('Reset button', 'leg 1 · leg 2', 'D4 · GND'),
        ('Battery (via switch)', '+ · −', 'VIN · GND')]
gf = s.shapes.add_table(len(ROWS), 3, Inches(ML), Inches(1.95), Inches(5.3), Inches(0.42 * len(ROWS)))
gf.name = 'Wiring table'
tbl = gf.table
tbl.first_row = True
tbl.horz_banding = False
for ci, wdt in enumerate((2.3, 1.5, 1.5)):
    tbl.columns[ci].width = Inches(wdt)
for ri, row in enumerate(ROWS):
    tbl.rows[ri].height = Inches(0.42)
    for ci, val in enumerate(row):
        cell = tbl.cell(ri, ci)
        tcPr = cell._tc.get_or_add_tcPr()
        for edge in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
            ln = etree.SubElement(tcPr, qn(edge))
            if edge == 'a:lnB':
                ln.set('w', '9525')
                sf = etree.SubElement(ln, qn('a:solidFill'))
                etree.SubElement(sf, qn('a:srgbClr')).set('val', C['line'])
            else:
                ln.set('w', '0')
                etree.SubElement(ln, qn('a:noFill'))
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(C['bg2'] if ri == 0 else C['bg'])
        cell.margin_left = cell.margin_right = Inches(0.1)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        head = ri == 0
        color = C['muted'] if head else (C['accent'] if ci == 2 else C['body'])
        fill_tf(cell.text_frame, [P(val.upper() if head else val, size=10 if head else 13, bold=head or ci == 2,
                                    font=HEAD if head else (MONO if ci == 2 else BODY), color=color,
                                    spacing=120 if head else None)], anchor='middle', margins=(0.1, 0.02, 0.06, 0.02))
text(s, ML, 5.55, 5.2, 0.6, [P('Each sensor needs one digital pin. The screen shares the two-wire I²C bus.', size=12, color=C['muted'], line=1.1)])

# 33 --------------------------------------------------------------- build
s = new_slide('build', notes=(
    "We designed the shell around the parts and 3D-printed it in white PLA. It's about 105 by 60 by 35 millimetres, "
    "in two pieces, a base and a top plate, so the electronics drop in from the top. The top plate has cut-outs for the "
    "screen, the button and the sensor cables, and the name is embossed right into the front wall."))
full_bleed(s, IMG['dims'], 'The shell with dimension lines: 105 mm wide, 60 mm deep, 35 mm tall', dx=0.75)
menu_pill(s)
eyebrow(s, 'Build')
title(s, 'Designed and 3D-printed')
FACTS = [('Size', '105 × 60 × 35 mm', True), ('Material', 'PLA, white', False), ('Pieces', 'Base tray and top plate', False),
         ('Openings', 'Screen window, button hole, two cable ports', False), ('Lettering', 'BOOK COUNTER, embossed on the front', False)]
for i, (k, v, mono) in enumerate(FACTS):
    y = 2.6 + i * 0.72
    line(s, ML, y, ML + 4.5, y, color=C['line'], width=0.75)
    text(s, ML, y + 0.14, 1.3, 0.4, [P(k.upper(), size=10, bold=True, font=HEAD, color=C['muted'], spacing=120)])
    text(s, ML + 1.4, y + 0.1, 3.1, 0.6, [P(v, size=14 if not mono else 15, bold=mono, font=MONO if mono else BODY,
                                             color=C['accent'] if mono else C['text'], line=1.05)])
line(s, ML, 2.6 + 5 * 0.72, ML + 4.5, 2.6 + 5 * 0.72, color=C['line'], width=0.75)

# 34 --------------------------------------------------------------- challenges
s = new_slide('challenges', bg=C['bg2'], notes=(
    "A page counter sounds simple, but a few things have to be right. One page could trigger a sensor more than once, "
    "so we added a short lock-out. Flipping back shouldn't add pages, which is why we use two sensors instead of one. "
    "Infrared also reflects off other things, so we tuned each sensor's range with its trimmer. And because pages curl "
    "as they turn, the sensors watch the top edge where every page moves the same way."))
menu_pill(s)
eyebrow(s, 'Design challenges')
title(s, 'What a page counter has to get right', w=11.5, h=0.8)
for i, (q, a) in enumerate([('One page could count twice.', 'A 400 ms lock-out after every count.'),
                            ('Flipping back must not add pages.', 'Two sensors, so the order of the hits gives the direction.'),
                            ('Room light and the page below reflect IR too.', "Trimmers tuned for a short range, aimed at the page's path."),
                            ('Pages are thin and curl as they turn.', 'The sensors watch the top edge, where every page passes the same way.')]):
    col, row = i % 2, i // 2
    card = rect(s, ML + col * 6.0, 2.15 + row * 2.3, 5.75, 2.05, fill=C['bg'], line=C['line'], radius=0.07)
    fill_tf(card.text_frame, [P(q, size=19, bold=True, font=HEAD, color=C['text'], after=10, line=1.0),
                              P(runs=[dict(t='Fix  ', bold=True, font=HEAD, color=C['accent'], size=13), dict(t=a)], size=16, line=1.12)],
            margins=(0.35, 0.3, 0.3, 0.2))

# 35 --------------------------------------------------------------- next
s = new_slide('next', notes=(
    "There's a lot we'd like to add next. Bluetooth, so each reading session goes to a phone app with graphs. "
    "Reading speed in pages per minute. A rechargeable battery with USB-C. A custom circuit board to make it smaller, "
    "and an adjustable clip so it fits any book."))
menu_pill(s)
eyebrow(s, 'Next version')
title(s, 'Where it goes next', w=10)
for i, (k, v) in enumerate([('Bluetooth sync', 'Send each session to a phone app with daily and weekly graphs.'),
                            ('Reading speed', 'Show pages per minute right on the screen.'),
                            ('Rechargeable', 'A lithium cell with USB-C charging instead of a 9 V battery.'),
                            ('Smaller', 'A custom circuit board so the sensor arms fold flat.'),
                            ('Fits any book', 'An adjustable clip for thin paperbacks and thick textbooks.')]):
    y = 2.15 + i * 0.92
    line(s, ML, y, W_IN - ML, y, color=C['line'], width=0.75)
    text(s, ML, y + 0.2, 3.4, 0.6, [P(k, size=22, bold=True, font=HEAD, color=C['accent'])])
    text(s, ML + 3.6, y + 0.24, 8.2, 0.6, [P(v, size=17, color=C['body'])])
line(s, ML, 2.15 + 5 * 0.92, W_IN - ML, 2.15 + 5 * 0.92, color=C['line'], width=0.75)

# 36 --------------------------------------------------------------- thanks
s = new_slide('thanks', notes=(
    "Thank you for listening. Book Counter counts the pages so you don't have to. We're happy to take questions, "
    "and the menu button lets us jump back to any part you'd like to see again."))
gif(s, f'{GIFS}/spin.gif', 5.45, 1.2, 7.75, 'The Book Counter turning on a turntable', name='Spinning model')
eyebrow(s, 'Thank you', y=2.0)
text(s, ML, 2.4, 5.0, 1.2, [P('Questions?', size=60, bold=True, font=HEAD, color=C['text'])])
text(s, ML, 3.65, 4.6, 0.9, [P("Book Counter counts the pages so you don't have to.", size=18, color=C['body'], line=1.1)])
text(s, ML, 4.6, 4.8, 0.3, [P('[Presenter names · Class]', size=12, bold=True, font=HEAD, color=C['muted'])])
button(s, ML, 5.3, 1.9, 0.5, 'Menu', 'menu', 'primary', size=14)
button(s, ML + 2.05, 5.3, 1.9, 0.5, 'Start again', 'title', 'secondary', size=14)

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
