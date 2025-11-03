# Website for the documentation: https://pandoc.org/installing.html

import os
from pathlib import Path
from datetime import datetime
import markdown

# ----------------------------
# Configuration
# ----------------------------
MD_DIR = Path(".")  # Current folder with your Markdown files
CSS_FILE = "style.css"
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
# Convert Markdown → HTML in same folder
# ----------------------------
for md_file in md_files:
    html_file = md_file.with_suffix(".html")
    print(f"Converting {md_file.name} → {html_file.name}")

    # Read Markdown content
    with open(md_file, "r", encoding="utf-8") as f:
        md_content = f.read()

    # Convert Markdown to HTML (images, headings, links, lists)
    html_body = markdown.markdown(md_content, extensions=['extra'])

    # Wrap in basic HTML template
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>{md_file.stem}</title>
    <link rel="stylesheet" href="{CSS_FILE}">
</head>
<body>
<header>
    <h1>{SITE_TITLE}</h1>
    <nav>
        <ul>
            <li><a href="index.html">Home</a></li>
            <li><a href="data.html">Data</a></li>
            <li><a href="howtos.html" class="active">How-Tos</a></li>
            <li><a href="about.html">About</a></li>
        </ul>
    </nav>
</header>
<main>
<div id="content">
{html_body}
</div>
</main>
<footer>
<p>© {SITE_TITLE} {datetime.now().year}</p>
</footer>
</body>
</html>
"""

    # Save HTML file in same folder
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(html_content)

print("✅ All Markdown files converted to HTML. Image paths are preserved!")
