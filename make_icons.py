#!/usr/bin/env python3
"""Generate Office 97 app icons and installer bitmaps using Pillow."""

import os
import struct
import zlib

def create_png(width, height, pixels_rgba):
    """Create a minimal PNG file from RGBA pixel data."""
    def png_chunk(chunk_type, data):
        length = len(data)
        chunk = chunk_type + data
        crc = zlib.crc32(chunk) & 0xffffffff
        return struct.pack('>I', length) + chunk + struct.pack('>I', crc)

    # IHDR
    ihdr_data = struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0)
    # Actually use RGBA (color type 6)
    ihdr_data = struct.pack('>II', width, height) + bytes([8, 6, 0, 0, 0])

    # IDAT - raw image data
    raw_data = b''
    for y in range(height):
        raw_data += b'\x00'  # filter type None
        for x in range(width):
            raw_data += pixels_rgba[y][x]

    compressed = zlib.compress(raw_data)

    png = b'\x89PNG\r\n\x1a\n'
    png += png_chunk(b'IHDR', ihdr_data)
    png += png_chunk(b'IDAT', compressed)
    png += png_chunk(b'IEND', b'')
    return png

def make_color(r, g, b, a=255):
    return bytes([r, g, b, a])

TRANSPARENT = make_color(0, 0, 0, 0)
WHITE = make_color(255, 255, 255)
BLACK = make_color(0, 0, 0)

def make_app_icon(filename, bg_color, letter, size=64):
    """Create a simple app icon: colored background + white letter."""
    pixels = []

    for y in range(size):
        row = []
        for x in range(size):
            # Rounded corners
            cx, cy = x - size//2, y - size//2
            corner_r = size // 8
            in_corner = False
            for ox, oy in [(-size//2+corner_r, -size//2+corner_r),
                           (size//2-corner_r-1, -size//2+corner_r),
                           (-size//2+corner_r, size//2-corner_r-1),
                           (size//2-corner_r-1, size//2-corner_r-1)]:
                if cx < -size//2+corner_r+1 or cx > size//2-corner_r-2:
                    if cy < -size//2+corner_r+1 or cy > size//2-corner_r-2:
                        dist = ((cx - ox) ** 2 + (cy - oy) ** 2) ** 0.5
                        if dist > corner_r:
                            in_corner = True

            if in_corner:
                row.append(TRANSPARENT)
            else:
                # Border
                margin = 2
                if x < margin or x >= size-margin or y < margin or y >= size-margin:
                    row.append(make_color(255, 255, 255, 180))
                else:
                    row.append(make_color(*bg_color))
        pixels.append(row)

    # Draw letter (simplified font)
    font = {
        'W': [(0.15, 0.2, 0.08, 0.55), (0.38, 0.2, 0.08, 0.55), (0.58, 0.2, 0.08, 0.55),
              (0.12, 0.65, 0.30, 0.08), (0.40, 0.60, 0.30, 0.08)],
        'X': [(0.15, 0.2, 0.08, 0.55), (0.57, 0.2, 0.08, 0.55),
              (0.15, 0.2, 0.55, 0.08), (0.15, 0.65, 0.55, 0.08),
              (0.33, 0.4, 0.14, 0.08)],
        'P': [(0.2, 0.2, 0.08, 0.6), (0.2, 0.2, 0.35, 0.08), (0.55, 0.2, 0.08, 0.3),
              (0.2, 0.45, 0.35, 0.08)],
        'A': [(0.15, 0.75, 0.08, -0.55), (0.55, 0.2, 0.08, 0.55),
              (0.15, 0.45, 0.5, 0.08), (0.28, 0.75, 0.08, 0.0)],
        'O': [(0.2, 0.2, 0.08, 0.6), (0.6, 0.2, 0.08, 0.6),
              (0.2, 0.2, 0.4, 0.08), (0.2, 0.73, 0.4, 0.08)],
    }

    # Simple pixel letter using block drawing
    ls = size // 8
    lx = size // 4
    ly = size // 5

    letter_pixels = set()
    if letter == 'W':
        for dy in range(size//2):
            t = dy / (size//2)
            # W shape
            for c in [lx, lx + size//4, lx + size//2]:
                for dx in range(-ls//2, ls//2+1):
                    if 0 <= c+dx < size and 0 <= ly+dy < size:
                        letter_pixels.add((ly+dy, c+dx))
            if dy > size//3:
                # Bottom connectors
                for dx in range(-ls//2, size//2+ls//2):
                    row_y = ly + size//2 - ls//2
                    if 0 <= lx+dx < size and 0 <= row_y < size:
                        letter_pixels.add((row_y, lx+dx))
    elif letter == 'X':
        for i in range(int(size*0.6)):
            t = i / int(size*0.6)
            x1 = int(lx + t * size * 0.5)
            x2 = int(lx + (1-t) * size * 0.5)
            y = ly + i
            for dd in range(-ls//2, ls//2+1):
                if 0 <= x1+dd < size and 0 <= y < size:
                    letter_pixels.add((y, x1+dd))
                if 0 <= x2+dd < size and 0 <= y < size:
                    letter_pixels.add((y, x2+dd))
    elif letter == 'P':
        bx = lx
        by = ly
        bw = int(size * 0.5)
        bh = int(size * 0.6)
        # Vertical bar
        for dy in range(bh):
            for dx in range(ls):
                letter_pixels.add((by+dy, bx+dx))
        # Top horizontal
        for dx in range(bw):
            for dy in range(ls):
                letter_pixels.add((by+dy, bx+dx))
        # Mid horizontal
        for dx in range(bw):
            for dy in range(ls):
                letter_pixels.add((by+bh//2+dy, bx+dx))
        # Right bar (top half only)
        for dy in range(bh//2):
            for dx in range(ls):
                letter_pixels.add((by+dy, bx+bw-ls+dx))
    elif letter == 'A':
        bx = lx
        by = ly
        bw = int(size * 0.5)
        bh = int(size * 0.6)
        for dy in range(bh):
            t = dy / bh
            # Left diagonal
            x1 = int(bx + t * bw/2)
            # Right diagonal
            x2 = int(bx + bw - t * bw/2)
            for dx in range(-ls//2, ls//2+1):
                if 0 <= x1+dx < size:
                    letter_pixels.add((by+dy, x1+dx))
                if 0 <= x2+dx < size:
                    letter_pixels.add((by+dy, x2+dx))
        # Crossbar
        mid_y = by + bh//2
        for dx in range(int(bw*0.2), int(bw*0.8)):
            for dy in range(ls):
                letter_pixels.add((mid_y+dy, bx+dx))

    for y, x in letter_pixels:
        if 0 <= y < size and 0 <= x < size:
            pixels[y][x] = WHITE

    return create_png(size, size, pixels)

def make_bmp(width, height, pixels_rgb, bits_per_pixel=24):
    """Create a BMP file."""
    row_size = ((width * bits_per_pixel // 8 + 3) // 4) * 4
    pixel_data_size = row_size * height
    file_size = 54 + pixel_data_size

    bmp = b'BM'
    bmp += struct.pack('<I', file_size)
    bmp += struct.pack('<HH', 0, 0)
    bmp += struct.pack('<I', 54)
    # DIB header
    bmp += struct.pack('<I', 40)
    bmp += struct.pack('<i', width)
    bmp += struct.pack('<i', -height)  # Top-down
    bmp += struct.pack('<HH', 1, bits_per_pixel)
    bmp += struct.pack('<I', 0)  # BI_RGB
    bmp += struct.pack('<I', pixel_data_size)
    bmp += struct.pack('<iI', 2835, 2835)
    bmp += struct.pack('<II', 0, 0)

    # Pixel data
    for y in range(height):
        row = b''
        for x in range(width):
            if y < len(pixels_rgb) and x < len(pixels_rgb[y]):
                r, g, b = pixels_rgb[y][x]
            else:
                r, g, b = 192, 192, 192
            row += bytes([b, g, r])
        # Pad row
        row += b'\x00' * (row_size - len(row))
        bmp += row

    return bmp

def make_gradient_bmp(width, height, color1, color2, orientation='vertical'):
    """Create a gradient BMP."""
    pixels = []
    for y in range(height):
        row = []
        for x in range(width):
            if orientation == 'vertical':
                t = y / height
            else:
                t = x / width
            r = int(color1[0] + t * (color2[0] - color1[0]))
            g = int(color1[1] + t * (color2[1] - color1[1]))
            b = int(color1[2] + t * (color2[2] - color1[2]))
            row.append((r, g, b))
        pixels.append(row)
    return make_bmp(width, height, pixels)

def make_office_sidebar_bmp():
    """Create Office 97 style installer sidebar (164x314)."""
    width, height = 164, 314
    pixels = []

    # Blue gradient background
    for y in range(height):
        row = []
        t = y / height
        for x in range(width):
            # Dark navy blue gradient
            r = int(0 + t * 10)
            g = int(0 + t * 50)
            b = int(100 + t * 80)
            row.append((r, g, b))
        pixels.append(row)

    # Add "Microsoft Office 97" text (simplified as blocks)
    # Title area
    title_y = 20
    title_color = (255, 255, 255)

    # Draw "Microsoft" in white (simplified)
    for ty in range(8):
        for tx in range(70):
            if (ty == 0 or ty == 7) or (tx % 10 == 0):
                if ty + title_y < height and tx + 10 < width:
                    pixels[ty + title_y][tx + 10] = title_color

    # Draw "Office" bigger
    for ty in range(16):
        for tx in range(80):
            if (ty == 0 or ty == 15) or (tx % 14 == 0):
                y_pos = title_y + 20
                if ty + y_pos < height and tx + 10 < width:
                    pixels[ty + y_pos][tx + 10] = (255, 200, 0)

    # Draw "97" yellow
    for ty in range(24):
        for tx in range(30):
            if tx < 4 or tx > 12:
                y_pos = title_y + 46
                if ty + y_pos < height and tx + 10 < width:
                    pixels[ty + y_pos][tx + 10] = (255, 220, 0)

    # Windows flag (4 colored squares)
    flag_x, flag_y = 20, 100
    flag_size = 20
    colors_flag = [(255, 0, 0), (0, 170, 0), (0, 0, 255), (255, 220, 0)]
    positions = [(0, 0), (flag_size+2, 0), (0, flag_size+2), (flag_size+2, flag_size+2)]
    for (fx, fy), color in zip(positions, colors_flag):
        for ty in range(flag_size):
            for tx in range(flag_size):
                py = flag_y + fy + ty
                px = flag_x + fx + tx
                if 0 <= py < height and 0 <= px < width:
                    pixels[py][px] = color

    # Bottom text area
    for ty in range(6):
        for tx in range(130):
            if (ty == 0 or ty == 5):
                y_pos = height - 40
                if ty + y_pos < height and tx + 10 < width:
                    pixels[ty + y_pos][tx + 10] = (200, 200, 255)

    return make_bmp(width, height, pixels)

def make_office_header_bmp():
    """Create Office 97 installer header (500x58)."""
    width, height = 500, 58
    pixels = []

    for y in range(height):
        row = []
        for x in range(width):
            # White background
            if y < 4 or y > height - 4 or x < 4 or x > width - 4:
                row.append((0, 0, 128))
            else:
                row.append((255, 255, 255))
        pixels.append(row)

    # Draw "Microsoft Office 97 Setup" text (simplified blocks)
    text_y = 15
    for ty in range(12):
        for tx in range(200):
            if ty == 0 or ty == 11:
                if tx + 20 < width:
                    pixels[text_y + ty][tx + 20] = (0, 0, 0)

    return make_bmp(width, height, pixels)

def main():
    os.makedirs('icons', exist_ok=True)
    os.makedirs('installer', exist_ok=True)

    # App icons
    apps = [
        ('word', (43, 87, 154), 'W'),
        ('excel', (33, 115, 70), 'X'),
        ('powerpoint', (208, 68, 35), 'P'),
        ('access', (164, 55, 58), 'A'),
        ('office', (0, 0, 128), 'O'),
    ]

    for name, color, letter in apps:
        png_data = make_app_icon('icons/' + name + '.png', color, letter, 64)
        with open('icons/' + name + '.png', 'wb') as f:
            f.write(png_data)
        print(f'Created icons/{name}.png')

    # Also create ICO for office (Windows)
    # Simple ICO = PNG wrapped
    ico_data = b'\x00\x00\x01\x00\x01\x00'  # ICO header: 1 image
    ico_data += bytes([64, 64, 0, 0, 1, 0, 32, 0])  # Image entry
    png_data = make_app_icon('', (0, 0, 128), 'O', 64)
    ico_data += struct.pack('<I', len(png_data))
    ico_data += struct.pack('<I', 6 + 16)  # Offset to image data
    ico_data += png_data
    # Fix offset
    ico_proper = b'\x00\x00\x01\x00\x01\x00'
    ico_proper += bytes([64, 64, 0, 0, 1, 0, 32, 0])
    offset = 6 + 16
    ico_proper += struct.pack('<I', len(png_data))
    ico_proper += struct.pack('<I', offset)
    ico_proper += png_data

    with open('icons/office.ico', 'wb') as f:
        f.write(ico_proper)
    print('Created icons/office.ico')

    # Installer bitmaps
    sidebar_bmp = make_office_sidebar_bmp()
    with open('installer/sidebar.bmp', 'wb') as f:
        f.write(sidebar_bmp)
    print('Created installer/sidebar.bmp (164x314)')

    header_bmp = make_office_header_bmp()
    with open('installer/header.bmp', 'wb') as f:
        f.write(header_bmp)
    print('Created installer/header.bmp (500x58)')

    print('\nAll icons and bitmaps created successfully.')

if __name__ == '__main__':
    main()
