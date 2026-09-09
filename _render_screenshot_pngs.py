"""Render evaluation evidence .txt files and JSON to PNG for dissertation (C8)."""
from __future__ import annotations

import json
from pathlib import Path

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    import subprocess
    subprocess.check_call(["pip", "install", "pillow", "-q"])
    from PIL import Image, ImageDraw, ImageFont

BASE = Path(r"D:\Downloads\Coursework\SCI701_Research_Pathway\Project3_JML_Identity_Lifecycle")
SHOTS = BASE / "05_Evaluation" / "screenshots"
EVIDENCE = BASE / "03_Prototype_Lab" / "evidence"

FONT = ImageFont.load_default()
LINE_H = 16
PAD = 20
MAX_W = 1100


def text_to_png(text: str, out_path: Path, title: str = ""):
    lines = (title + "\n" + text if title else text).splitlines()
    h = PAD * 2 + LINE_H * (len(lines) + 2)
    img = Image.new("RGB", (MAX_W, max(h, 200)), "white")
    draw = ImageDraw.Draw(img)
    y = PAD
    if title:
        draw.text((PAD, y), title, fill="#1f4e79", font=FONT)
        y += LINE_H * 2
    for line in lines:
        draw.text((PAD, y), line[:130], fill="black", font=FONT)
        y += LINE_H
    img.save(out_path)
    print(f"Wrote {out_path.name}")


def render_from_txt():
    mapping = {
        "01_Entra_Users_List.txt": "01_Entra_Users_List.png",
        "02_Entra_Groups_List.txt": "02_Entra_Groups_List.png",
        "03_Joiner_Group_Memberships.txt": "03_Joiner_Group_Memberships.png",
        "04_Leaver_Disabled_Account.txt": "04_Leaver_Disabled_Account.png",
        "05_Privileged_Group_Members.txt": "05_Privileged_Group_Members.png",
        "06_PowerShell_Orphan_Report.txt": "06_PowerShell_Orphan_Report.png",
    }
    for txt, png in mapping.items():
        p = SHOTS / txt
        if p.exists():
            text_to_png(p.read_text(encoding="utf-8"), SHOTS / png, title=txt.replace(".txt", "").replace("_", " "))


def render_orphan_from_json():
    for f in sorted(EVIDENCE.glob("*Orphan*.json")):
        data = json.loads(f.read_text(encoding="utf-8-sig"))
        items = data if isinstance(data, list) else [data]
        lines = ["Orphan Privileged Account Report", "=" * 40]
        for item in items:
            for k, v in item.items():
                lines.append(f"{k}: {v}")
        text_to_png("\n".join(lines), SHOTS / "06_PowerShell_Orphan_Report.png", title="Lab Evidence - Orphan Report")


def render_metrics_chart():
    metrics = json.loads((BASE / "05_Evaluation" / "evaluation_summary.json").read_text(encoding="utf-8"))
    b = metrics["baseline_manual_minutes"]
    a = metrics["automated_minutes"]
    lines = [
        "Evaluation Metrics - Manual vs Automated (minutes)",
        "=" * 50,
        f"Joiner:  Manual {b['joiner']} min  ->  Automated {a['joiner_min']} min",
        f"Mover:   Manual {b['mover']} min  ->  Automated {a['mover_min']} min",
        f"Leaver:  Manual {b['leaver']} min  ->  Automated {a['leaver_min']} min",
        "",
        f"Evidence files: {metrics['jml_events_with_evidence']}",
        f"Orphan privileged detected: {metrics['orphan_privileged_detected']}",
        f"Tenant: niazkhanghostgmail.onmicrosoft.com",
    ]
    text_to_png("\n".join(lines), SHOTS / "00_Evaluation_Metrics_Summary.png")


if __name__ == "__main__":
    SHOTS.mkdir(parents=True, exist_ok=True)
    render_orphan_from_json()
    render_metrics_chart()
    render_from_txt()
    print("PNG render complete.")
