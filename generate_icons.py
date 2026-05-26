#!/usr/bin/env python3
"""
generate_icons.py — creates simple placeholder icons for the add-in.
Run once: python3 generate_icons.py
You can replace the icons in assets/ with your own 16x16, 32x32, 80x80 PNG files.
"""
import struct, zlib, os

def make_png(size, colour=(26, 179, 148)):
    """Creates a solid-colour PNG at the given pixel size."""
    r, g, b = colour
    # PNG signature
    sig = b'\x89PNG\r\n\x1a\n'
    # IHDR chunk
    w = h = size
    ihdr_data = struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)
    ihdr = make_chunk(b'IHDR', ihdr_data)
    # IDAT chunk (image data)
    raw_row = b'\x00' + bytes([r, g, b]) * w   # filter byte + RGB pixels
    raw = raw_row * h
    idat = make_chunk(b'IDAT', zlib.compress(raw))
    # IEND chunk
    iend = make_chunk(b'IEND', b'')
    return sig + ihdr + idat + iend

def make_chunk(chunk_type, data):
    c = chunk_type + data
    return struct.pack('>I', len(data)) + c + struct.pack('>I', zlib.crc32(c) & 0xffffffff)

os.makedirs('assets', exist_ok=True)
for size in [16, 32, 80]:
    path = f'assets/icon-{size}.png'
    with open(path, 'wb') as f:
        f.write(make_png(size))
    print(f'Created {path}')
print('Icons created. Replace with your own if desired.')
