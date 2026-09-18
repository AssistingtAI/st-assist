# -*- coding: utf-8 -*-
"""替换提额页的客服二维码。

用法（把二维码图片拖到终端，或直接给路径）：
    python set_wx_qr.py "C:\\path\\to\\你的二维码.png"
    python set_wx_qr.py "你的二维码.jpg" --size 400

做的事：
  1. 校验图片可用（PIL 能打开、非空）
  2. 自动补白边变正方形（二维码留白不足会扫不出）
  3. 缩放到 --size（默认 400px，页面显示时再缩到 92px，保证高清屏不糊）
  4. 备份原文件为 wx-qr.png.bak（若已存在旧备份则保留最早那份）
  5. 覆盖写入 wx-qr.png

替换后需推送：git push origin main
"""
import argparse
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
TARGET = os.path.join(HERE, "wx-qr.png")


def main():
    ap = argparse.ArgumentParser(description="替换提额页客服二维码")
    ap.add_argument("src", help="二维码图片路径（png/jpg/jpeg/webp）")
    ap.add_argument("--size", type=int, default=400, help="输出边长像素，默认 400")
    ap.add_argument("--quiet", action="store_true", help="静默模式")
    args = ap.parse_args()

    src = args.src.strip().strip('"').strip("'")
    if not os.path.isabs(src):
        src = os.path.join(os.getcwd(), src)
    if not os.path.exists(src):
        print("[错误] 找不到文件: %s" % src)
        sys.exit(1)

    try:
        from PIL import Image
    except ImportError:
        print("[错误] 需要 Pillow：pip install Pillow")
        sys.exit(1)

    try:
        im = Image.open(src)
        im.load()
    except Exception as e:
        print("[错误] 图片无法解析: %s" % e)
        sys.exit(1)

    w, h = im.size
    if w < 120 or h < 120:
        print("[警告] 原图仅 %dx%d，偏小可能影响扫码识别（建议 >=300px）" % (w, h))

    # 转 RGB（去 alpha，避免白底变黑）
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        bg.alpha_composite(im)
        im = bg.convert("RGB")
    elif im.mode != "RGB":
        im = im.convert("RGB")

    # 补成正方形（白底居中），二维码四周需留白
    side = max(im.size)
    pad_px = int(side * 0.06)          # 约 6% 留白
    canvas_side = side + pad_px * 2
    canvas = Image.new("RGB", (canvas_side, canvas_side), (255, 255, 255))
    canvas.paste(im, ((canvas_side - w) // 2, (canvas_side - h) // 2))

    out = canvas.resize((args.size, args.size), Image.LANCZOS)

    # 备份（只保留最早的原始占位图）
    bak = TARGET + ".bak"
    if os.path.exists(TARGET) and not os.path.exists(bak):
        shutil.copy2(TARGET, bak)
        if not args.quiet:
            print("[备份] 原占位图 -> %s" % os.path.basename(bak))

    out.save(TARGET, "PNG", optimize=True)

    if not args.quiet:
        print("[完成] %dx%d -> %s (%dx%d, %d 字节)"
              % (w, h, os.path.basename(TARGET), args.size, args.size,
                 os.path.getsize(TARGET)))
        print("[提示] 页面显示尺寸 92px，自动缩放；现在可 git push origin main")


if __name__ == "__main__":
    main()
