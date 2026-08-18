#!/usr/bin/env python3
"""
make_cover.py -- substitutes the real Grade/Unit text into
KNO_Reading_TEMPLATE_8_7.pptx's own slide 1 (the RedThread wordmark cover),
then renders JUST that slide to a PNG. This keeps the cover graphic
pixel-exact to the source template rather than hand-recreated as vector
shapes, since it's a complex gradient logo + wave illustration.

Usage:
    python3 make_cover.py "Grade 4, Knowledge Unit 2" /path/to/output.png
"""
import sys
import shutil
import zipfile
import os
import subprocess

TEMPLATE = os.path.join(os.path.dirname(__file__), "KNO_Reading_TEMPLATE_8_7.pptx")

def make_cover(grade_unit_text, out_png):
    work = "/tmp/cover_work_" + str(abs(hash(grade_unit_text)))
    if os.path.exists(work):
        shutil.rmtree(work)
    os.makedirs(work)

    with zipfile.ZipFile(TEMPLATE, 'r') as z:
        z.extractall(f"{work}/extracted")

    slide1_path = f"{work}/extracted/ppt/slides/slide1.xml"
    content = open(slide1_path, encoding='utf-8').read()
    content = content.replace("Grade [#], Knowledge Unit [#]", grade_unit_text)
    open(slide1_path, 'w', encoding='utf-8').write(content)

    out_pptx = f"{work}/cover_only.pptx"
    base_dir = f"{work}/extracted"
    with zipfile.ZipFile(out_pptx, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(base_dir):
            for file in files:
                full = os.path.join(root, file)
                rel = os.path.relpath(full, base_dir)
                zf.write(full, rel)

    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", out_pptx, "--outdir", work],
        check=True, capture_output=True,
    )
    subprocess.run(
        ["pdftoppm", "-png", "-r", "200", "-f", "1", "-l", "1",
         f"{work}/cover_only.pdf", f"{work}/coverimg"],
        check=True, capture_output=True,
    )
    shutil.move(f"{work}/coverimg-01.png", out_png)
    shutil.rmtree(work)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit('Usage: python3 make_cover.py "Grade 4, Knowledge Unit 2" out.png')
    make_cover(sys.argv[1], sys.argv[2])
    print(f"Wrote {sys.argv[2]}")
