from pathlib import Path
import sys

import pypdfium2 as pdfium

sys.stdout.reconfigure(encoding="utf-8")

pdf_path = Path(sys.argv[1]).resolve()
output_dir = Path(sys.argv[2]).resolve()
scale = float(sys.argv[3]) if len(sys.argv) > 3 else 1.7
output_dir.mkdir(parents=True, exist_ok=True)

pdf = pdfium.PdfDocument(str(pdf_path))
for index in range(len(pdf)):
    page = pdf[index]
    bitmap = page.render(scale=scale)
    image = bitmap.to_pil()
    target = output_dir / f"page-{index + 1:02d}.png"
    image.save(target)
    print(target)
