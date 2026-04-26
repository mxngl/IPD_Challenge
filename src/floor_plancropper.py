from pathlib import Path
import json

import matplotlib.pyplot as plt
import numpy as np
from PIL import Image

try:
    import fitz  # PyMuPDF
except ImportError as exc:
    raise ImportError(
        "This cell needs PyMuPDF. Install it with `pip install pymupdf pillow`."
    ) from exc

FLOOR_PLAN_DIR = Path("../floor_plans")
CROPPED_DIR = FLOOR_PLAN_DIR / "cropped_png"
CROPPED_DIR.mkdir(parents=True, exist_ok=True)

pdf_paths = sorted(FLOOR_PLAN_DIR.glob("*.pdf"))
if not pdf_paths:
    raise FileNotFoundError(f"No PDFs found in {FLOOR_PLAN_DIR.resolve()}")

print("Available floor plans:")
for idx, pdf_path in enumerate(pdf_paths):
    print(f"  [{idx}] {pdf_path.name}")

selection = input("Enter the PDF index to crop [0]: ").strip()
pdf_index = int(selection) if selection else 0
pdf_path = pdf_paths[pdf_index]

page_selection = input("Enter the page number to load [0]: ").strip()
page_index = int(page_selection) if page_selection else 0
zoom_selection = input("Enter render zoom for the preview [2.0]: ").strip()
zoom = float(zoom_selection) if zoom_selection else 2.0

doc = fitz.open(pdf_path)
page = doc.load_page(page_index)
pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
image_np = np.array(image)
doc.close()

fig, ax = plt.subplots(figsize=(12, 12))
ax.imshow(image_np)
ax.set_title(
    "Click 2 opposite grid-extents corners: top-left first, bottom-right second",
    fontsize=12,
)
ax.axis("on")
plt.tight_layout()
points = plt.ginput(2, timeout=-1)
plt.close(fig)

if len(points) != 2:
    raise ValueError("Expected exactly 2 clicks: top-left and bottom-right corners.")

(x1, y1), (x2, y2) = points
left = int(round(min(x1, x2)))
right = int(round(max(x1, x2)))
top = int(round(min(y1, y2)))
bottom = int(round(max(y1, y2)))

cropped = image.crop((left, top, right, bottom))
png_path = CROPPED_DIR / f"{pdf_path.stem}_page_{page_index}_cropped.png"
json_path = CROPPED_DIR / f"{pdf_path.stem}_page_{page_index}_crop_points.json"
cropped.save(png_path)

json_path.write_text(
    json.dumps(
        {
            "pdf": str(pdf_path.resolve()),
            "page_index": page_index,
            "zoom": zoom,
            "clicked_points": [
                {"x": round(x1, 2), "y": round(y1, 2)},
                {"x": round(x2, 2), "y": round(y2, 2)},
            ],
            "crop_box": {
                "left": left,
                "top": top,
                "right": right,
                "bottom": bottom,
            },
            "png": str(png_path.resolve()),
        },
        indent=2,
    ),
    encoding="utf-8",
)

print(f"Saved cropped PNG to: {png_path.resolve()}")
print(f"Saved crop metadata to: {json_path.resolve()}")

fig, ax = plt.subplots(figsize=(10, 10))
ax.imshow(np.array(cropped))
ax.set_title(f"Cropped floor plan: {pdf_path.name}")
ax.axis("off")
plt.show()
