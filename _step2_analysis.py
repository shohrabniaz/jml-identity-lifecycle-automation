# Step 2 - Strengthen analysis (B1, B2, B3)
"""Update literature pack, as-is risk findings, and freeze to-be design."""
from __future__ import annotations

from datetime import date
from pathlib import Path

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

BASE = Path(r"D:\Downloads\Coursework\SCI701_Research_Pathway\Project3_JML_Identity_Lifecycle")
TODAY = date.today().isoformat()
STUDENT = "Md Shohrab Hossen Niaz"
SID = "1211589"
HEADER = f"{STUDENT} | Student ID {SID} | SCI700/701 Project 3 | Individual | Fictional MSP + lab-only"

ACADEMIC_SOURCES = [
    {
        "apa": (
            "Cameron, A., & Grewe, O. (2022). An overview of the digital identity lifecycle (v2). "
            "IDPro Body of Knowledge, 1(7). https://doi.org/10.55621/idpro.31"
        ),
        "search": "UniSC Library: search 'Cameron Grewe digital identity lifecycle' or DOI 10.55621/idpro.31",
        "relevance": (
            "Defines Joiner, Mover, and Leaver as the three principal workforce identity processes. "
            "Supports CMIT as-is analysis and the requirement for attestation, access removal, and "
            "audit when identities change state. Directly maps to Project 3 RQ1 and governance design."
        ),
        "quote_theme": "JML is the standard framework for workforce IAM lifecycle phases.",
    },
    {
        "apa": (
            "Hevner, A. R., March, S. T., Park, J., & Ram, S. (2004). Design science in information "
            "systems research. MIS Quarterly, 28(1), 75-105."
        ),
        "search": "UniSC Library: MIS Quarterly Hevner design science 2004",
        "relevance": (
            "Provides the design-science research methodology used in this project: artefact (JML "
            "automation prototype), evaluation against utility goals, and contribution to practice "
            "for SME MSPs. Frames RQ3 evaluation metrics as design-evaluation criteria."
        ),
        "quote_theme": "Design-science produces and evaluates IT artefacts to solve organisational problems.",
    },
    {
        "apa": (
            "Tuononen, H. (2023). Privileged access management model for a managed service provider "
            "(Master's thesis). JAMK University of Applied Sciences. http://theseus.fi/handle/10024/800216"
        ),
        "search": "UniSC Library: search 'Tuononen privileged access management MSP' or Theseus 10024/800216",
        "relevance": (
            "Examines PAM for MSPs administering multiple customer environments with highly privileged "
            "credentials. Supports the CMIT scenario where a single compromised MSP identity cascades "
            "into client tenants. Justifies RQ2 focus on timely privileged deprovisioning and audit trails."
        ),
        "quote_theme": "MSP privileged access requires secure, documented, and monitored administration models.",
    },
]

STANDARDS = [
    ("ASD Essential Eight — Restrict administrative privileges", "Baseline control for privileged access minimisation in Australian context."),
    ("ACSC — MSP security guidance", "Multi-tenant MSP identity risk and supplier cyber risk management."),
    ("ISO/IEC 27001:2022 A.5.15–A.5.18", "Identity lifecycle, access provisioning, and review controls."),
    ("Microsoft Entra Lifecycle Workflows", "Native cloud JML automation (contrast with Graph/PowerShell prototype)."),
    ("Microsoft Graph API", "Implementation mechanism for lab prototype scripts."),
]


def set_normal(doc):
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)


def add_title(doc, text, level=0):
    doc.add_heading(text, level=level)


def add_para(doc, text, bold=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    return p


def build_literature():
    doc = Document()
    set_normal(doc)
    add_title(doc, "Literature & Standards Pack — Project 3", 0)
    add_para(doc, HEADER)
    add_para(doc, f"Document date: {TODAY} | Step 2 (B1) — annotated bibliography added")

    add_title(doc, "1. Academic sources (UniSC library — search terms provided)", 1)
    for i, src in enumerate(ACADEMIC_SOURCES, 1):
        add_title(doc, f"1.{i} Source {i}", 2)
        add_para(doc, src["apa"])
        add_para(doc, f"UniSC search: {src['search']}", bold=True)
        add_para(doc, f"Relevance to Project 3: {src['relevance']}")
        add_para(doc, f"Key theme: {src['quote_theme']}")

    add_title(doc, "2. Standards and industry guidance", 1)
    for title, note in STANDARDS:
        add_title(doc, title, 2)
        add_para(doc, note)

    add_title(doc, "3. Synthesis for research questions", 1)
    add_para(doc, (
        "RQ1 (affordable JML automation): Cameron & Grewe (2022) define JML processes; "
        "this project demonstrates a low-cost Graph/PowerShell prototype on Entra ID Free "
        "rather than enterprise IGA platforms."
    ))
    add_para(doc, (
        "RQ2 (privileged deprovisioning): Tuononen (2023) and ACSC/ASD guidance establish "
        "MSP privileged-access risk; the prototype leaver script and orphan report operationalise "
        "timely removal and detection."
    ))
    add_para(doc, (
        "RQ3 (measuring effectiveness): Hevner et al. (2004) design-science evaluation; "
        "metrics include provisioning time, evidence coverage, and orphan privileged detection "
        "(see Evaluation_Workbook.xlsx)."
    ))

    path = BASE / "02_Literature" / "Literature_Standards_Pack.docx"
    doc.save(path)
    print(f"Wrote {path}")

    # Update source_list.txt
    lines = ["ACADEMIC (UniSC library)\n"]
    for i, s in enumerate(ACADEMIC_SOURCES, 1):
        lines.append(f"{i}. {s['apa']}")
        lines.append(f"   Search: {s['search']}\n")
    lines.append("\nSTANDARDS & INDUSTRY\n")
    for i, (t, _) in enumerate(STANDARDS, 1):
        lines.append(f"{i}. {t}")
    (BASE / "02_Literature" / "source_list.txt").write_text("\n".join(lines), encoding="utf-8")


def build_asis():
    doc = Document()
    set_normal(doc)
    add_title(doc, "As-Is JML Process Map & Risk Findings — Coastal Managed IT (Fictional MSP)", 0)
    add_para(doc, HEADER)
    add_para(doc, f"Document date: {TODAY} | Status: REVIEWED v1.0 (Step 2 — B2)")

    add_title(doc, "1. Scenario", 1)
    add_para(doc, (
        "Coastal Managed IT Pty Ltd (CMIT) is a fictional Adelaide MSP (~35 staff) administering IT "
        "for dozens of SME clients. Staff hold privileged access into client Microsoft tenants and RMM "
        "tools. HR/PSA updates are manual. There is no automated Joiner-Mover-Leaver (JML) workflow."
    ))

    add_title(doc, "2. Actors (swimlane roles)", 1)
    for actor in [
        "HR / Office Manager — PSA updates (often delayed)",
        "Service Desk — tickets, day-1 access requests",
        "IAM / Entra Admin — account and group changes (senior engineer dual-hatting)",
        "Team Lead / Engineer Manager — informal mover notifications",
        "Client Tenant Admin — customer-side approvals (inconsistent)",
        "Audit / Compliance — limited machine-readable evidence",
    ]:
        add_para(doc, f"• {actor}")

    add_title(doc, "3. As-is process summary", 1)
    processes = [
        ("Joiner", [
            "Offer accepted; PSA record created days later (sometimes after day 1).",
            "Ticket raised ad hoc; Entra account and groups chosen from memory.",
            "MFA configured late; client access added by urgent request without expiry.",
            "Failure modes: delayed day-1 access, over-provisioning, weak approval evidence.",
        ]),
        ("Mover", [
            "Role change announced in chat or standup.",
            "Some groups added; old client access often left untouched.",
            "Failure modes: privilege creep, stale client admin rights.",
        ]),
        ("Leaver", [
            "HR notifies IT late or after last day.",
            "Mailbox disabled eventually; privileged roles removed manually.",
            "Failure modes: orphaned privileged accounts, weak audit trail.",
        ]),
    ]
    for name, steps in processes:
        add_title(doc, name, 2)
        for s in steps:
            add_para(doc, f"• {s}")

    add_title(doc, "4. Swimlane (text representation)", 1)
    table = doc.add_table(rows=4, cols=6)
    hdr = ["Step", "HR/PSA", "Service Desk", "IAM Admin", "Privileged Approver", "Audit"]
    for i, h in enumerate(hdr):
        table.rows[0].cells[i].text = h
    rows = [
        ["Joiner", "Create PSA record", "Validate ticket", "Create account + groups", "Approve priv packs", "Email/ticket only"],
        ["Mover", "Update role (late)", "Confirm change", "Partial group update", "Rarely involved", "No standard log"],
        ["Leaver", "Mark terminated", "Priority ticket", "Manual disable", "Manual priv removal", "Incomplete evidence"],
    ]
    for r, row in enumerate(rows, 1):
        for c, val in enumerate(row):
            table.rows[r].cells[c].text = val

    add_title(doc, "5. Risk findings", 1)
    risks = [
        ("RF-01", "Orphan privileged accounts after leaver", "High", "High", "Lab orphan report detected leaver.test02 disabled but still in SG-CMIT-Privileged-Admins"),
        ("RF-02", "Privilege creep on mover", "Medium", "High", "Old client groups not removed; mover.test01 scenario"),
        ("RF-03", "Delayed joiner provisioning", "Medium", "Medium", "Manual ticket chain; 45 min baseline in evaluation"),
        ("RF-04", "Weak audit trail", "High", "High", "No machine-readable JML log in as-is; email/ticket only"),
        ("RF-05", "MSP cascade compromise", "High", "Medium", "Single MSP identity reaches multiple client tenants (Tuononen 2023; ACSC)"),
    ]
    rt = doc.add_table(rows=1, cols=5)
    for i, h in enumerate(["ID", "Risk", "Likelihood", "Impact", "Evidence / notes"]):
        rt.rows[0].cells[i].text = h
    for row in risks:
        cells = rt.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = v

    add_title(doc, "6. Lab validation of as-is risks (2026-09-01)", 1)
    add_para(doc, (
        "The lab deliberately simulated as-is leaver failure: leaver.test02 was disabled without "
        "removing privileged group membership. 06_Orphan_Privileged_Report.ps1 flagged "
        "DISABLED_BUT_STILL_IN_PRIV_GROUP — confirming RF-01 is measurable and reportable."
    ))

    path = BASE / "01_Process_Maps" / "AsIs_JML_Process_and_Risk_Findings.docx"
    doc.save(path)
    print(f"Wrote {path}")


def build_tobe_frozen():
    doc = Document()
    set_normal(doc)
    add_title(doc, "To-Be JML Target Design — Automated Lifecycle for CMIT Lab", 0)
    add_para(doc, HEADER)
    add_para(doc, f"Document date: {TODAY} | Status: FROZEN v1.0 FINAL (Step 2 — B3)")
    add_para(doc, "This document is the design baseline for implementation and dissertation Chapters 4–6.")

    add_title(doc, "1. Design goals (mapped to RQs)", 1)
    goals = [
        ("G1", "RQ1", "Affordable SME/MSP automation using Entra ID Free + Graph/PowerShell"),
        ("G2", "RQ2", "Timely, auditable privileged deprovision on leaver events"),
        ("G3", "RQ3", "Measurable effectiveness via timestamps, test cases, orphan report"),
        ("G4", "Ethics", "Lab-safe fictional tenants only; synthetic users"),
    ]
    gt = doc.add_table(rows=1, cols=3)
    gt.rows[0].cells[0].text = "Goal"
    gt.rows[0].cells[1].text = "RQ"
    gt.rows[0].cells[2].text = "Description"
    for g in goals:
        row = gt.add_row().cells
        row[0].text, row[1].text, row[2].text = g

    add_title(doc, "2. Target architecture", 1)
    add_para(doc, (
        "Trigger (simulated HR/PSA or approved request) → Automation engine (PowerShell/Graph) → "
        "Entra ID actions (enable/disable, group membership) → Evidence store (JSON per event + Entra audit logs)."
    ))

    add_title(doc, "3. Role packs (frozen)", 1)
    packs = [
        ("AllStaff", "SG-CMIT-AllStaff", "Every joiner"),
        ("ServiceDesk", "SG-CMIT-AllStaff + SG-CMIT-ServiceDesk", "Joiner ServiceDesk role pack"),
        ("Engineer", "SG-CMIT-AllStaff + SG-CMIT-Engineers", "Joiner Engineer role pack"),
        ("ClientA", "SG-ClientA-Support", "Optional client portfolio access"),
        ("Privileged", "SG-CMIT-Privileged-Admins", "Separate approval; stripped on leaver"),
    ]
    pt = doc.add_table(rows=1, cols=3)
    pt.rows[0].cells[0].text = "Pack"
    pt.rows[0].cells[1].text = "Groups"
    pt.rows[0].cells[2].text = "Use"
    for p in packs:
        row = pt.add_row().cells
        row[0].text, row[1].text, row[2].text = p

    add_title(doc, "4. JML workflows (frozen)", 1)
    add_title(doc, "Joiner", 2)
    add_para(doc, "Approved request → create/enable user → apply role pack groups → write JSON evidence.")
    add_title(doc, "Mover", 2)
    add_para(doc, "Approved change → remove old role groups → add new role pack → update department → evidence.")
    add_title(doc, "Leaver (critical path)", 2)
    add_para(doc, (
        "Leaver trigger → remove privileged/client/role groups → disable account → revoke sessions → "
        "evidence JSON → optional orphan privileged report."
    ))

    add_title(doc, "5. Implementation mapping (lab prototype)", 1)
    impl = [
        ("01_Create_Lab_Groups.ps1", "Provision security groups", "8 groups created"),
        ("02_Create_Synthetic_Users.ps1", "Lab users", "5 users"),
        ("03_Invoke_Joiner.ps1", "Joiner workflow", "TC01, TC02 Pass"),
        ("04_Invoke_Mover.ps1", "Mover workflow", "TC03 Pass"),
        ("05_Invoke_Leaver.ps1", "Leaver workflow", "TC04 Pass"),
        ("06_Orphan_Privileged_Report.ps1", "Governance detection", "TC05 Pass — 1 orphan"),
        ("Run-AllLab.ps1", "End-to-end test suite", "8 evidence files"),
    ]
    it = doc.add_table(rows=1, cols=3)
    it.rows[0].cells[0].text = "Script"
    it.rows[0].cells[1].text = "Design element"
    it.rows[0].cells[2].text = "Lab result (2026-09-01)"
    for row in impl:
        cells = it.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = v

    add_title(doc, "6. Out of scope (frozen)", 1)
    for item in [
        "Production HR/PSA integration",
        "Paid M365 licensing or Exchange mailboxes",
        "Entra ID P2 / Lifecycle Workflows (optional future work)",
        "Real client tenant federation / GDAP",
        "Azure paid resources (VMs, databases)",
    ]:
        add_para(doc, f"• {item}")

    add_title(doc, "7. Evaluation criteria (frozen)", 1)
    add_para(doc, "Joiner/Mover/Leaver time (manual vs automated); audit evidence coverage; orphan privileged detection.")

    path = BASE / "01_Process_Maps" / "ToBe_JML_Target_Design.docx"
    doc.save(path)
    print(f"Wrote {path}")

    # Also update shorter ToBe_JML_Process.docx with frozen stamp
    doc2 = Document(BASE / "01_Process_Maps" / "ToBe_JML_Process.docx")
    if doc2.paragraphs:
        doc2.paragraphs[0].text = "To-Be JML Process — FROZEN v1.0 FINAL"
    doc2.save(BASE / "01_Process_Maps" / "ToBe_JML_Process.docx")


def update_checklist():
    path = BASE / "00_Master_Checklist.txt"
    text = path.read_text(encoding="utf-8")
    block = """
STEP 2 - ANALYSIS (Phase B)
[x] B1 Literature pack — 3 academic sources + annotated bibliography
[x] B2 As-is process + risk findings reviewed (v1.0)
[x] B3 To-be target design frozen (v1.0 FINAL)
"""
    if "STEP 2 - ANALYSIS" not in text:
        path.write_text(text.rstrip() + block, encoding="utf-8")


def main():
    build_literature()
    build_asis()
    build_tobe_frozen()
    update_checklist()
    print("Step 2 complete.")


if __name__ == "__main__":
    main()
