#!/usr/bin/env python3
"""Draws the Wurd app icons into icons/: the header logo (a white W on an indigo rounded square).

Standard library only, so it runs anywhere Python 3 does. Shapes are drawn from signed distances,
which gives anti-aliased edges with one sample per pixel. Run it from anywhere:

    python3 tools/make-icons.py
"""
import math
import os
import struct
import zlib

INDIGO = (0x4F, 0x46, 0xE5)
WHITE = (0xFF, 0xFF, 0xFF)

# All geometry is in units of the icon's width (0..1)
CORNER = 0.25  # matches .brand__logo: 8px radius on a 32px square
W_POINTS = [(0.24, 0.31), (0.37, 0.69), (0.5, 0.4), (0.63, 0.69), (0.76, 0.31)]
W_STROKE = 0.09

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'icons')


def rounded_square_distance(x, y, radius):
    # Signed distance to a 0..1 square with rounded corners; negative inside
    qx = abs(x - 0.5) - (0.5 - radius)
    qy = abs(y - 0.5) - (0.5 - radius)
    outside = math.hypot(max(qx, 0.0), max(qy, 0.0))
    return outside + min(max(qx, qy), 0.0) - radius


def segment_distance(x, y, a, b):
    ax, ay = a
    bx, by = b
    dx, dy = bx - ax, by - ay
    t = max(0.0, min(1.0, ((x - ax) * dx + (y - ay) * dy) / (dx * dx + dy * dy)))
    return math.hypot(x - ax - t * dx, y - ay - t * dy)


def w_distance(x, y):
    # Distance to the W stroke's edge (round caps and joins); negative inside
    d = min(segment_distance(x, y, a, b) for a, b in zip(W_POINTS, W_POINTS[1:]))
    return d - W_STROKE / 2


def coverage(distance, size):
    # Distance in icon units -> how much of a pixel the shape covers
    return max(0.0, min(1.0, 0.5 - distance * size))


def render(size, rounded):
    rows = []
    for py in range(size):
        row = bytearray([0])  # PNG filter: none
        y = (py + 0.5) / size
        for px in range(size):
            x = (px + 0.5) / size
            alpha = coverage(rounded_square_distance(x, y, CORNER), size) if rounded else 1.0
            ink = coverage(w_distance(x, y), size)
            rgb = [round(bg + (fg - bg) * ink) for bg, fg in zip(INDIGO, WHITE)]
            row += bytes(rgb + [round(alpha * 255)])
        rows.append(bytes(row))
    return b''.join(rows)


def write_png(path, size, pixels):
    def chunk(kind, data):
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))

    header = struct.pack('>IIBBBBB', size, size, 8, 6, 0, 0, 0)  # 8-bit RGBA
    with open(path, 'wb') as f:
        f.write(b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', header) + chunk(b'IDAT', zlib.compress(pixels, 9)) + chunk(b'IEND', b''))


def write_svg(path):
    s = 512
    points = ' '.join(f'{x * s:g},{y * s:g}' for x, y in W_POINTS)
    with open(path, 'w') as f:
        f.write(
            f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {s} {s}">'
            f'<rect width="{s}" height="{s}" rx="{CORNER * s:g}" fill="#4f46e5"/>'
            f'<polyline points="{points}" fill="none" stroke="#fff" stroke-width="{W_STROKE * s:g}" '
            'stroke-linecap="round" stroke-linejoin="round"/></svg>\n'
        )


def main():
    os.makedirs(OUT, exist_ok=True)
    write_svg(os.path.join(OUT, 'icon.svg'))
    # The W sits inside the maskable safe zone (a centred circle of radius 0.4), so the same
    # drawing serves as a maskable icon once it fills the square edge to edge.
    # iOS rounds the apple-touch-icon itself, and needs it opaque.
    for name, size, rounded in [
        ('icon-192.png', 192, True),
        ('icon-512.png', 512, True),
        ('icon-maskable-512.png', 512, False),
        ('apple-touch-icon.png', 180, False),
    ]:
        write_png(os.path.join(OUT, name), size, render(size, rounded))
        print(f'icons/{name}')
    print('icons/icon.svg')


if __name__ == '__main__':
    main()
