#!/usr/bin/env python3
"""Navigation audit for the interactive deck: every button resolves, every hidden slide is reachable and has a way out."""
import sys
import zipfile
from collections import deque

from pptx import Presentation
from pptx.enum.action import PP_ACTION

path = sys.argv[1]
prs = Presentation(path)
slides = list(prs.slides)
idx = {s.slide_id: i for i, s in enumerate(slides)}
hidden = {i for i, s in enumerate(slides) if s._element.get('show') == '0'}
edges = {i: set() for i in range(len(slides))}
problems = []

for i, s in enumerate(slides):
    for shp in s.shapes:
        try:
            ca = shp.click_action
        except Exception:
            continue
        if ca.action == PP_ACTION.HYPERLINK:
            problems.append(f'slide {i + 1}: "{shp.name}" links to an external address {ca.hyperlink.address}')
        if ca.action == PP_ACTION.NAMED_SLIDE:
            tgt = ca.target_slide
            if tgt is None or tgt.slide_id not in idx:
                problems.append(f'slide {i + 1}: "{shp.name}" points at a missing slide')
            else:
                edges[i].add(idx[tgt.slide_id])

# sequential advance: from a visible slide, "next" goes to the next visible slide
visible = [i for i in range(len(slides)) if i not in hidden]
for a, b in zip(visible, visible[1:]):
    edges[a].add(b)

seen, q = {0}, deque([0])
while q:
    n = q.popleft()
    for m in edges[n]:
        if m not in seen:
            seen.add(m)
            q.append(m)
for i in hidden:
    if i not in seen:
        problems.append(f'hidden slide {i + 1} can never be reached')
    if not any(m != i for m in edges[i]):
        problems.append(f'hidden slide {i + 1} has no button leading out')
for i in range(len(slides)):
    if i not in seen:
        problems.append(f'slide {i + 1} unreachable')

with zipfile.ZipFile(path) as z:
    names = z.namelist()
    gifs = [n for n in names if n.endswith('.gif')]
    glbs = [n for n in names if n.endswith('.glb')]
    size = sum(z.getinfo(n).file_size for n in names)

print(f'{len(slides)} slides, {len(hidden)} hidden, {sum(len(e) for e in edges.values())} navigation edges')
print(f'media: {len(gifs)} gif, {len(glbs)} glb, package {size / 1e6:.1f} MB uncompressed')
for i, s in enumerate(slides):
    outs = sorted(m + 1 for m in edges[i])
    print(f'  {i + 1:2d}{" (hidden)" if i in hidden else "         "} -> {outs}')
print('PROBLEMS:' if problems else 'OK: every button resolves; every hidden slide is reachable and has a way out')
for p in problems:
    print('  -', p)
sys.exit(1 if problems else 0)
