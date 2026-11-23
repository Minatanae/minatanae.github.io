# Builds icons and Timeline-Teams.zip. Run: python3 -I build.py (from this folder)
import zlib, struct, zipfile, os, math
here = os.path.dirname(os.path.abspath(__file__))

def png(path, w, h, px):
    raw = b''.join(b'\x00' + bytes(px[y*w*4:(y+1)*w*4]) for y in range(h))
    chunk = lambda t, d: struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    with open(path, 'wb') as f:
        f.write(b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 6, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw, 9)) + chunk(b'IEND', b''))

def glyph(u, v):
    # u,v in 0..1. horizontal line with three nodes, plus small stubs up/down on the middle node
    cy = 0.5
    nodes = [0.2, 0.5, 0.8]
    r = 0.085
    for nx in nodes:
        if math.hypot(u-nx, v-cy) <= r: return True
    if 0.2 <= u <= 0.8 and abs(v-cy) <= 0.025: return True
    if abs(u-0.5) <= 0.025 and 0.24 <= v <= 0.76: return True
    for vy in (0.24, 0.76):
        if math.hypot(u-0.5, v-vy) <= 0.05: return True
    return False

def render(size, bg, ss=4):
    px = []
    for y in range(size):
        for x in range(size):
            hit = sum(glyph((x+(i+.5)/ss)/size, (y+(j+.5)/ss)/size) for i in range(ss) for j in range(ss)) / (ss*ss)
            if bg:
                r, g, b = [int(c*(1-hit) + 255*hit) for c in bg]; px += [r, g, b, 255]
            else:
                px += [255, 255, 255, int(255*hit)]
    return px

png(os.path.join(here, 'color.png'), 192, 192, render(192, (0x3b, 0x5b, 0xdb)))
png(os.path.join(here, 'outline.png'), 32, 32, render(32, None))
with zipfile.ZipFile(os.path.join(here, 'Timeline-Teams.zip'), 'w', zipfile.ZIP_DEFLATED) as z:
    for n in ('manifest.json', 'color.png', 'outline.png'):
        z.write(os.path.join(here, n), n)
print('ok')
