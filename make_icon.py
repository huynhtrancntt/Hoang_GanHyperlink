# -*- coding: utf-8 -*-
"""
make_icon.py  - Build icon.ico (multi-size) from icon_src.jpg
Compatible with Pillow 12.x+ and Windows cp1252 CI environments.

Usage:
    venv\\Scripts\\python make_icon.py
    python make_icon.py

Requires: pip install Pillow
"""
import io
import struct
import sys
from pathlib import Path

# Force UTF-8 output so emoji/unicode don't crash on Windows CI (cp1252)
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ── Install Pillow if missing ────────────────────────────────────────────────
try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])
    from PIL import Image, ImageDraw, ImageFont

# ── Paths ────────────────────────────────────────────────────────────────────
HERE    = Path(__file__).parent
JPG_SRC = HERE / "icon_src.jpg"
ICO_DST = HERE / "icon.ico"

# ── Load or create source image ──────────────────────────────────────────────
if JPG_SRC.exists():
    print(f"[INFO] Source image: {JPG_SRC.name}")
    src = Image.open(JPG_SRC).convert("RGBA")
else:
    print(f"[WARN] {JPG_SRC.name} not found - generating placeholder icon.")
    src = Image.new("RGBA", (512, 512), color=(13, 27, 42, 255))
    draw = ImageDraw.Draw(src)
    try:
        font = ImageFont.truetype("arial.ttf", 200)
    except Exception:
        font = ImageFont.load_default()
    draw.text((80, 120), "GH", fill=(255, 190, 40, 255), font=font)

# ── Resize to 7 standard Windows icon sizes ──────────────────────────────────
SIZES = [16, 24, 32, 48, 64, 128, 256]

png_bufs = []
for s in SIZES:
    buf = io.BytesIO()
    src.resize((s, s), Image.LANCZOS).save(buf, format="PNG")
    png_bufs.append(buf.getvalue())

# ── Write ICO binary manually (Windows ICO format) ───────────────────────────
# Avoids Pillow 12.x multi-frame ICO bug
n = len(png_bufs)
ico_header = struct.pack("<HHH", 0, 1, n)  # reserved=0, type=1(ICO), count=n

offset = 6 + n * 16  # 6-byte header + 16 bytes per directory entry
dir_entries = b""
for s, data in zip(SIZES, png_bufs):
    w = s if s < 256 else 0   # 256 -> 0 per ICO spec
    h = s if s < 256 else 0
    dir_entries += struct.pack(
        "<BBBBHHII",
        w, h,        # width, height
        0, 0,        # color count, reserved
        1,           # planes
        32,          # bit count
        len(data),   # image data size
        offset,      # offset to image data
    )
    offset += len(data)

with open(ICO_DST, "wb") as f:
    f.write(ico_header + dir_entries)
    for data in png_bufs:
        f.write(data)

sz = ICO_DST.stat().st_size
print(f"[OK] icon.ico saved -> {ICO_DST.name}  ({sz:,} bytes | {sz // 1024} KB)")
print(f"     Sizes: {SIZES}")
