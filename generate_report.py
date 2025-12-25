import json
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook

HTML_TEMPLATE = """<!doctype html>
<html><head><meta charset="utf-8"><title>Report</title></head>
<body>
<h1>Reporting Triple Output</h1>
<p>Generated at: {ts}</p>
<table border="1" cellpadding="6">
<tr><th>Item</th><th>Value</th></tr>
{rows}
</table>
</body></html>
"""

def main(out_dir: str = "out"):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    data = {
        "generated_at": datetime.utcnow().isoformat() + "Z",
        "metrics": [
            {"name": "total_tests", "value": 128},
            {"name": "passed", "value": 124},
            {"name": "failed", "value": 4},
        ],
    }

    # JSON
    (out / "report.json").write_text(json.dumps(data, indent=2), encoding="utf-8")

    # Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "metrics"
    ws.append(["name", "value"])
    for m in data["metrics"]:
        ws.append([m["name"], m["value"]])
    wb.save(out / "report.xlsx")

    # HTML
    rows = "\n".join([f"<tr><td>{m['name']}</td><td>{m['value']}</td></tr>" for m in data["metrics"]])
    html = HTML_TEMPLATE.format(ts=data["generated_at"], rows=rows)
    (out / "report.html").write_text(html, encoding="utf-8")

    print("Generated:", out / "report.json", out / "report.xlsx", out / "report.html")

if __name__ == "__main__":
    main()
