# Website for the documentation: https://pandoc.org/installing.html

import subprocess
from pathlib import Path
from bs4 import BeautifulSoup
import shutil

# ----------------------------
# Configuration
# ----------------------------
MD_DIR = Path(".")         # Current folder with Markdown files
CSS_FILE = "style.css"     # CSS file relative to this folder
SITE_TITLE = "The Scholar’s Tome"

# ----------------------------
# List Markdown files
# ----------------------------
md_files = list(MD_DIR.glob("*.md"))
if not md_files:
    print("No Markdown files found in the current folder.")
    exit()
else:
    print(f"Markdown files found: {[f.name for f in md_files]}")

# ----------------------------
# Convert Markdown → HTML using Pandoc
# ----------------------------
for md_file in md_files:
    html_file = md_file.with_suffix(".html")
    print(f"Converting {md_file.name} → {html_file.name}")

    # Updated template to match other pages
    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>{md_file.stem} | {SITE_TITLE}</title>
<link rel="stylesheet" href="../{CSS_FILE}"/>
</head>
<body>
<header>
    <h1>{SITE_TITLE}</h1>
    <nav>
        <ul>
            <li><a href="../index.html">Home</a></li>
            <li><a href="../data.html">Data</a></li>
            <li><a href="../howtos.html" class="active">How-Tos</a></li>
            <li><a href="../about.html">About</a></li>
        </ul>
    </nav>
</header>

<main>
    <section>
        {{CONTENT}}
    </section>
</main>

<footer>
    <p>© <span id="year"></span> {SITE_TITLE}</p>
</footer>
<script>
document.getElementById("year").textContent = new Date().getFullYear();
</script>
</body>
</html>"""

    # Save template temporarily
    template_file = MD_DIR / "pandoc_template.html"
    with open(template_file, "w", encoding="utf-8") as f:
        f.write(html_template.replace("{CONTENT}", "$body$"))

    # Run Pandoc
    subprocess.run([
        "pandoc",
        str(md_file),
        "-f", "markdown",
        "-t", "html",
        "-s",
        "-o", str(html_file),
        "--css", CSS_FILE,
        "--template", str(template_file),
        "--resource-path", str(MD_DIR)
    ])

    # Remove temporary template
    template_file.unlink()

# ----------------------------
# Fix <img> paths in all HTML files
# ----------------------------
html_files = list(MD_DIR.glob("*.html"))

for html_file in html_files:
    print(f"Fixing image paths in {html_file.name}...")
    with open(html_file, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    for img in soup.find_all("img"):
        src = img.get("src", "")
        if src.startswith("howtos/"):
            img["src"] = src[len("howtos/"):]
            print(f"  Updated src: {src} -> {img['src']}")

    with open(html_file, "w", encoding="utf-8") as f:
        f.write(str(soup))

# ----------------------------
# Move all Markdown files into "md" folder
# ----------------------------
md_folder = MD_DIR / "md"
md_folder.mkdir(exist_ok=True)

for md_file in md_files:
    dest = md_folder / md_file.name
    shutil.move(str(md_file), str(dest))
    print(f"Moved {md_file.name} → {md_folder}/")

print("✅ All Markdown files converted to HTML, images fixed, and Markdown moved to 'md' folder!")
