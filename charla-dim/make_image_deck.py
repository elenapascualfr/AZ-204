"""Genera una versión solo-imágenes (compatible con Keynote) con las notas del orador."""
import subprocess, sys, tempfile, glob, os
from pptx import Presentation
from pptx.util import Inches

src, pdf, out = sys.argv[1:4]
tmp = tempfile.mkdtemp()
subprocess.run(["pdftoppm", "-r", "110", "-jpeg", "-jpegopt", "quality=88", pdf, f"{tmp}/s"], check=True)
imgs = sorted(glob.glob(f"{tmp}/s-*.jpg"))
orig = Presentation(src)
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
blank = prs.slide_layouts[6]
for img, o in zip(imgs, orig.slides):
    sl = prs.slides.add_slide(blank)
    sl.shapes.add_picture(img, 0, 0, prs.slide_width, prs.slide_height)
    if o.has_notes_slide:
        sl.notes_slide.notes_text_frame.text = o.notes_slide.notes_text_frame.text
prs.save(out)
print(len(imgs), "diapositivas ->", out)
