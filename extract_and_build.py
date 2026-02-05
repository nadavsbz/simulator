import os, re, zipfile
from openpyxl import load_workbook

XLSM = [f for f in os.listdir('.') if f.endswith('.xlsm')][0]
wb = load_workbook(XLSM, data_only=True)

os.makedirs("images", exist_ok=True)

# --- Extract images from sheet P ---
sheet = wb["P"]
for img in sheet._images:
    anchor = img.anchor._from
    row = anchor.row + 1
    qid = sheet.cell(row=row, column=12).value
    if not isinstance(qid, str):
        continue
    m = re.search(r"(\\d+)", qid)
    if not m:
        continue
    qnum = m.group(1)
    with open(f"images/q_{qnum}.png", "wb") as f:
        f.write(img._data())

# --- Build index.html ---
html = """<!DOCTYPE html>
<html lang="he"><head><meta charset="UTF-8">
<title>סימולטור מבחן</title></head><body>
<h1>סימולטור</h1>
<script>
const images = {};
"""
for f in os.listdir("images"):
    if f.startswith("q_"):
        num = f.split("_")[1].split(".")[0]
        html += f'images[{num}] = "images/{f}";\n'

html += """
document.body.innerHTML += "<p>תמונות זמינות: " + Object.keys(images).length + "</p>";
</script></body></html>
"""

with open("index.html", "w", encoding="utf8") as f:
    f.write(html)

# --- Zip result ---
with zipfile.ZipFile("site.zip", "w") as z:
    z.write("index.html")
    for f in os.listdir("images"):
        z.write("images/" + f)
