import sys
from PIL import Image, ImageEnhance
src, out = sys.argv[1], sys.argv[2]
boxes = {"nhb": ((10, 322, 800, 505), 1.38), "ibm": ((28, 118, 860, 345), 1), "c6": ((35, 262, 820, 395), 1),
         "star": ((80, 350, 1120, 440), 1), "metr": ((60, 140, 1220, 455), 1), "amzn": ((225, 160, 1055, 440), 1)}
for n, (b, br) in boxes.items():
    im = Image.open(f"{src}/{n}.png").convert("RGB").crop(b)
    if br != 1: im = ImageEnhance.Brightness(im).enhance(br)
    im.save(f"{out}/src_{n}.png"); print(n, im.size, round(im.size[0]/im.size[1], 3))
