import sys, glob
from PIL import Image, ImageDraw
fs = sorted(glob.glob(sys.argv[1] + "/s*.png")); W, H = 533, 300; cols = 4
rows = (len(fs) + cols - 1) // cols
sh = Image.new("RGB", (cols * W, rows * (H + 18)), "white"); d = ImageDraw.Draw(sh)
for i, f in enumerate(fs):
    im = Image.open(f).convert("RGB").resize((W, H)); x, y = (i % cols) * W, (i // cols) * (H + 18)
    sh.paste(im, (x, y + 18)); d.text((x + 4, y + 3), f.split("/")[-1], fill="black")
sh.save(sys.argv[2])
