r"""
make_icon.py
============
Tạo file icon.ico (multi-size) từ ảnh nguồn icon_src.jpg.
Dùng cách viết ICO thủ công để tương thích Pillow 12.x+.

Chạy:
    venv\Scripts\python make_icon.py
    (hoặc: python make_icon.py)

Yêu cầu: pip install Pillow
"""
import io
import struct
import sys
from pathlib import Path

# ── Cài Pillow nếu chưa có ──────────────────────────────────────────────────
try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])
    from PIL import Image, ImageDraw, ImageFont

# ── Đường dẫn ───────────────────────────────────────────────────────────────
HERE    = Path(__file__).parent
JPG_SRC = HERE / "icon_src.jpg"
ICO_DST = HERE / "icon.ico"

# ── Load hoặc tạo ảnh nguồn ─────────────────────────────────────────────────
if JPG_SRC.exists():
    print(f"📂 Dùng ảnh nguồn: {JPG_SRC.name}")
    src = Image.open(JPG_SRC).convert("RGBA")
else:
    print(f"⚠  Không tìm thấy {JPG_SRC.name} – tạo icon placeholder.")
    src = Image.new("RGBA", (512, 512), color=(13, 27, 42, 255))
    draw = ImageDraw.Draw(src)
    try:
        font = ImageFont.truetype("arial.ttf", 200)
    except Exception:
        font = ImageFont.load_default()
    draw.text((80, 120), "GH", fill=(255, 190, 40, 255), font=font)

# ── Resize thành 7 kích cỡ chuẩn Windows ───────────────────────────────────
SIZES = [16, 24, 32, 48, 64, 128, 256]

png_bufs = []
for s in SIZES:
    buf = io.BytesIO()
    src.resize((s, s), Image.LANCZOS).save(buf, format="PNG")
    png_bufs.append(buf.getvalue())

# ── Viết file ICO thủ công (chuẩn Windows ICO binary format) ────────────────
# Tương thích Pillow 12.x (tránh bug multi-frame ICO)
n = len(png_bufs)
ico_header = struct.pack("<HHH", 0, 1, n)   # reserved=0, type=1(ICO), count=n

offset = 6 + n * 16   # header 6 bytes + mỗi entry 16 bytes
dir_entries = b""
for s, data in zip(SIZES, png_bufs):
    w = s if s < 256 else 0          # 256 → 0 theo chuẩn ICO
    h = s if s < 256 else 0
    dir_entries += struct.pack(
        "<BBBBHHII",
        w, h,          # width, height
        0, 0,          # color count, reserved
        1,             # planes
        32,            # bit count
        len(data),     # size of image data
        offset,        # offset to image data
    )
    offset += len(data)

with open(ICO_DST, "wb") as f:
    f.write(ico_header + dir_entries)
    for data in png_bufs:
        f.write(data)

sz = ICO_DST.stat().st_size
print(f"✅ Icon saved → {ICO_DST}  ({sz:,} bytes | {sz // 1024} KB)")
print(f"   Gồm {n} kích cỡ: {SIZES}")
