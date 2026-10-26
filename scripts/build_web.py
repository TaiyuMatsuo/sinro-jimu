"""data/checklist.json を埋め込んで Web 版チェックシートを生成する"""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
data = json.loads((ROOT / "data" / "checklist.json").read_text(encoding="utf-8"))
html = (ROOT / "web" / "template.html").read_text(encoding="utf-8")
out = ROOT / "output" / "私立高校_Web出願チェック.html"
out.write_text(html.replace("__DATA__", json.dumps(data, ensure_ascii=False)), encoding="utf-8")
print(out)
