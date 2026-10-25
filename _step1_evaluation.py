# Step 1 - Complete evaluation evidence (C5 + C8)
"""Fill baseline metrics and create screenshot capture pack for Project 3."""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

BASE = Path(r"D:\Downloads\Coursework\SCI701_Research_Pathway\Project3_JML_Identity_Lifecycle")
EVAL = BASE / "05_Evaluation"
SCREENSHOTS = EVAL / "screenshots"
EVIDENCE = BASE / "03_Prototype_Lab" / "evidence"

# Manual baseline assumptions (CMIT as-is process, documented for evaluation)
BASELINES = {
    "joiner_manual_min": 45,
    "joiner_manual_note": (
        "HR/PSA record + Service Desk ticket + IAM creates account, assigns groups ad hoc, "
        "client access requested separately. Typical MSP delay to day 1 access."
    ),
    "mover_manual_min": 30,
    "mover_manual_note": (
        "Manager email/chat notification; partial group updates; old client roles often remain."
    ),
    "leaver_manual_min": 60,
    "leaver_manual_note": (
        "HR notifies IT (often late); mailbox disabled; privileged/client roles removed manually "
        "and inconsistently. Active admin effort per leaver event."
    ),
    "leaver_calendar_days": 3,
    "leaver_calendar_note": "Median delay before full deprovision in manual MSP model (literature/industry).",
}

# Automated times from lab run 2026-09-01 (script execution, minutes rounded)
AUTOMATED = {
    "joiner_min": 1,
    "mover_min": 1,
    "leaver_min": 2,
    "full_jml_suite_min": 3,
}


def load_evidence_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def build_workbook():
    wb_path = EVAL / "Evaluation_Workbook.xlsx"
    wb = load_workbook(wb_path)

    # --- Metrics sheet ---
    ws = wb["Metrics"]
    ws["B2"], ws["C2"], ws["D2"] = BASELINES["joiner_manual_min"], AUTOMATED["joiner_min"], improvement(BASELINES["joiner_manual_min"], AUTOMATED["joiner_min"])
    ws["B3"], ws["C3"], ws["D3"] = BASELINES["mover_manual_min"], AUTOMATED["mover_min"], improvement(BASELINES["mover_manual_min"], AUTOMATED["mover_min"])
    ws["B4"], ws["C4"], ws["D4"] = BASELINES["leaver_manual_min"], AUTOMATED["leaver_min"], improvement(BASELINES["leaver_manual_min"], AUTOMATED["leaver_min"])

    evidence_files = sorted(EVIDENCE.glob("*.json")) if EVIDENCE.exists() else []
    orphan = 0
    for f in evidence_files:
        if "OrphanPrivilegedReport" in f.name:
            data = load_evidence_json(f)
            items = data if isinstance(data, list) else [data]
            orphan = sum(1 for x in items if x.get("Risk") == "DISABLED_BUT_STILL_IN_PRIV_GROUP")

    ws["B5"], ws["C5"], ws["D5"] = "Partial / email only", len(evidence_files), "100% JSON evidence"
    ws["B6"], ws["C6"], ws["D6"] = "Unknown / not measured", orphan, "Detectable via scheduled report"

    # --- TestCases sheet ---
    ws2 = wb["TestCases"]
    tests = [
        ("TC01", "Joiner ServiceDesk", "AllStaff + ServiceDesk groups", "Pass", "20260901_012529_Joiner_joiner.test01.json"),
        ("TC02", "Joiner Engineer + ClientA", "AllStaff + Engineers + ClientA", "Pass", "20260901_012530_Joiner_joiner.test02.json"),
        ("TC03", "Mover SD to Engineer", "Role groups swapped", "Pass", "20260901_012531_Mover_mover.test01.json"),
        ("TC04", "Leaver automated", "Account disabled + groups stripped", "Pass", "20260901_012534_Leaver_leaver.test01.json"),
        ("TC05", "Orphan privileged report", "Flags disabled user in priv group", "Pass", "20260901_012535_OrphanPrivilegedReport.json"),
    ]
    if ws2.max_column < 5:
        ws2.cell(1, 5, "Evidence file")
    for i, row in enumerate(tests, start=2):
        for j, val in enumerate(row, start=1):
            ws2.cell(i, j, val)

    # --- Baseline_Methodology sheet ---
    if "Baseline_Methodology" in wb.sheetnames:
        del wb["Baseline_Methodology"]
    ws3 = wb.create_sheet("Baseline_Methodology")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    rows = [
        ("Section", "Detail"),
        ("Student", "Md Shohrab Hossen Niaz | "),
        ("Lab tenant", "niazkhanghostgmail.onmicrosoft.com"),
        ("Evaluation date", "2026-09-01"),
        ("", ""),
        ("MANUAL BASELINE (As-Is)", ""),
        ("Joiner time (min)", f"{BASELINES['joiner_manual_min']} — {BASELINES['joiner_manual_note']}"),
        ("Mover time (min)", f"{BASELINES['mover_manual_min']} — {BASELINES['mover_manual_note']}"),
        ("Leaver admin time (min)", f"{BASELINES['leaver_manual_min']} — {BASELINES['leaver_manual_note']}"),
        ("Leaver calendar delay (days)", f"{BASELINES['leaver_calendar_days']} — {BASELINES['leaver_calendar_note']}"),
        ("Audit evidence (manual)", "Tickets, emails, inconsistent PSA notes; no machine-readable JML log"),
        ("Orphan privileged visibility", "Unknown until periodic manual review or audit finding"),
        ("", ""),
        ("AUTOMATED (To-Be / Lab)", ""),
        ("Joiner script time (min)", f"~{AUTOMATED['joiner_min']} (03_Invoke_Joiner.ps1)"),
        ("Mover script time (min)", f"~{AUTOMATED['mover_min']} (04_Invoke_Mover.ps1)"),
        ("Leaver script time (min)", f"~{AUTOMATED['leaver_min']} (05_Invoke_Leaver.ps1)"),
        ("Full suite (min)", f"~{AUTOMATED['full_jml_suite_min']} (Run-AllLab.ps1 steps 3-9)"),
        ("Audit evidence (automated)", f"{len(evidence_files)} JSON files in 03_Prototype_Lab/evidence/"),
        ("Orphan privileged detected", f"{orphan} (leaver.test02 — simulated as-is failure)"),
        ("", ""),
        ("SCREENSHOTS (C8)", "See 05_Evaluation/screenshots/ — portal evidence for dissertation"),
    ]
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row, start=1):
            cell = ws3.cell(r, c, val)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if r == 1:
                cell.font = header_font
                cell.fill = header_fill

    ws3.column_dimensions["A"].width = 28
    ws3.column_dimensions["B"].width = 90

    wb.save(wb_path)
    print(f"Updated {wb_path}")


def improvement(manual: int, auto: int) -> str:
    pct = round((manual - auto) / manual * 100, 1)
    return f"{pct}% faster"


def write_capture_checklist():
    SCREENSHOTS.mkdir(parents=True, exist_ok=True)
    checklist = """SCREENSHOT CAPTURE CHECKLIST (C8)
Student: Md Shohrab Hossen Niaz
Portal: https://entra.microsoft.com
Tenant: niazkhanghostgmail.onmicrosoft.com

Save each screenshot with the EXACT filename below into this folder.

[ ] 01_Entra_Users_List.png
    Entra ID > Users > All users (show 5 synthetic users)

[ ] 02_Entra_Groups_List.png
    Entra ID > Groups > All groups (show SG-CMIT-* groups)

[ ] 03_Joiner_Group_Memberships.png
    Users > joiner.test01 > Groups (AllStaff + ServiceDesk)

[ ] 04_Leaver_Disabled_Account.png
    Users > leaver.test01 > Properties (Account enabled = No)

[ ] 05_Privileged_Group_Members.png
    Groups > SG-CMIT-Privileged-Admins > Members (leaver.test02)

[ ] 06_PowerShell_Orphan_Report.png
    Re-run: .\\06_Orphan_Privileged_Report.ps1
    Capture console showing DISABLED_BUT_STILL_IN_PRIV_GROUP

OPTIONAL:
[ ] 07_Entra_Audit_Log.png — Audit logs filtered for user/group changes

TIPS:
- Blur or crop if any real email appears (use lab accounts only)
- Win+Shift+S to snip on Windows
- PNG format preferred for dissertation
"""
    (SCREENSHOTS / "CAPTURE_CHECKLIST.txt").write_text(checklist, encoding="utf-8")

    # Evidence summary HTML for quick reference (not a substitute for portal shots)
    orphan_data = None
    for f in sorted(EVIDENCE.glob("*Orphan*.json")):
        orphan_data = load_evidence_json(f)
    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Lab Evidence Summary</title>
<style>body{{font-family:Segoe UI,sans-serif;max-width:900px;margin:2rem auto;padding:0 1rem}}
table{{border-collapse:collapse;width:100%}} th,td{{border:1px solid #ccc;padding:8px;text-align:left}}
th{{background:#1f4e79;color:#fff}} .risk{{color:#c00;font-weight:bold}}</style></head>
<body>
<h1>Project 3 Lab Evidence Summary</h1>
<p>Tenant: niazkhanghostgmail.onmicrosoft.com | Evaluated: 2026-09-01</p>
<h2>Metrics</h2>
<table>
<tr><th>Metric</th><th>Manual</th><th>Automated</th></tr>
<tr><td>Joiner (min)</td><td>{BASELINES['joiner_manual_min']}</td><td>{AUTOMATED['joiner_min']}</td></tr>
<tr><td>Mover (min)</td><td>{BASELINES['mover_manual_min']}</td><td>{AUTOMATED['mover_min']}</td></tr>
<tr><td>Leaver (min)</td><td>{BASELINES['leaver_manual_min']}</td><td>{AUTOMATED['leaver_min']}</td></tr>
</table>
<h2>Orphan privileged finding</h2>
<pre>{json.dumps(orphan_data, indent=2) if orphan_data else 'N/A'}</pre>
<h2>Evidence files</h2>
<ul>
{''.join(f'<li>{f.name}</li>' for f in sorted(EVIDENCE.glob('*.json')))}
</ul>
<p><em>Replace portal screenshots per CAPTURE_CHECKLIST.txt for final dissertation.</em></p>
</body></html>"""
    (SCREENSHOTS / "Evidence_Summary_Reference.html").write_text(html, encoding="utf-8")
    print(f"Created {SCREENSHOTS}")


def update_summary_json():
    evidence_files = sorted(EVIDENCE.glob("*.json"))
    orphan = 0
    for f in evidence_files:
        if "OrphanPrivilegedReport" in f.name:
            data = load_evidence_json(f)
            items = data if isinstance(data, list) else [data]
            orphan = sum(1 for x in items if x.get("Risk") == "DISABLED_BUT_STILL_IN_PRIV_GROUP")
    summary = {
        "evaluation_date": "2026-09-01",
        "baseline_manual_minutes": {
            "joiner": BASELINES["joiner_manual_min"],
            "mover": BASELINES["mover_manual_min"],
            "leaver": BASELINES["leaver_manual_min"],
        },
        "automated_minutes": AUTOMATED,
        "jml_events_with_evidence": len(evidence_files),
        "orphan_privileged_detected": orphan,
        "screenshots_required": 6,
        "screenshots_folder": "05_Evaluation/screenshots",
        "step1_status": "C5 complete; C8 checklist ready",
    }
    (EVAL / "evaluation_summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")


def update_master_checklist():
    path = BASE / "00_Master_Checklist.txt"
    text = path.read_text(encoding="utf-8")
    additions = """
STEP 1 - EVALUATION EVIDENCE (Phase C5 + C8)
[x] C5 Manual baseline documented in Evaluation_Workbook.xlsx (Baseline_Methodology sheet)
[ ] C8 Portal screenshots (6 files) - see 05_Evaluation/screenshots/CAPTURE_CHECKLIST.txt
"""
    if "STEP 1 - EVALUATION" not in text:
        path.write_text(text.rstrip() + additions, encoding="utf-8")


if __name__ == "__main__":
    build_workbook()
    write_capture_checklist()
    update_summary_json()
    update_master_checklist()
    print("Step 1 (C5) complete. C8: capture screenshots using checklist.")
