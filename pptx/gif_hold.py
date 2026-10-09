#!/usr/bin/env python3
"""Make a GIF stop on its last frame: set that frame's delay to the maximum (655.35 s).
Works whatever the viewer does with the loop count, so a page-turn clip plays once and holds the new count."""
import sys


def frame_delay_offsets(b):
    assert b[:6] in (b'GIF87a', b'GIF89a'), 'not a GIF'
    pos = 13
    if b[10] & 0x80:
        pos += 3 * (2 ** ((b[10] & 0x07) + 1))
    gce = []
    while pos < len(b):
        tag = b[pos]
        if tag == 0x3B:          # trailer
            break
        if tag == 0x21:          # extension
            label = b[pos + 1]
            if label == 0xF9:    # graphic control: [21 F9 04 packed delayLo delayHi trans 00]
                gce.append(pos + 4)
            pos += 2
            while b[pos]:
                pos += b[pos] + 1
            pos += 1
        elif tag == 0x2C:        # image descriptor
            packed = b[pos + 9]
            pos += 10
            if packed & 0x80:
                pos += 3 * (2 ** ((packed & 0x07) + 1))
            pos += 1             # LZW minimum code size
            while b[pos]:
                pos += b[pos] + 1
            pos += 1
        else:
            raise ValueError(f'unexpected block 0x{tag:02x} at {pos}')
    return gce


for path in sys.argv[1:]:
    data = bytearray(open(path, 'rb').read())
    offs = frame_delay_offsets(data)
    data[offs[-1]:offs[-1] + 2] = (65535).to_bytes(2, 'little')
    open(path, 'wb').write(data)
    print(f'{path}: {len(offs)} frames, last frame held')
