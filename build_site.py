"""Bundles site/ into docs/index.html (one self-contained file for GitHub Pages)."""
from pathlib import Path
root = Path(__file__).parent
html = (root / "site" / "index.html").read_text(encoding="utf-8")
for name in ["matcher.js", "sample-data.js"]:
    js = (root / "site" / name).read_text(encoding="utf-8")
    html = html.replace(f'<script src="{name}"></script>', f"<script>\n{js}\n</script>")
out = root / "docs"
out.mkdir(exist_ok=True)
(out / "index.html").write_text(html, encoding="utf-8")
print("wrote docs/index.html", len(html), "bytes")
