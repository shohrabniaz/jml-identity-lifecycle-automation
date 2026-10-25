# Build Project 3 dissertation draft chapters (Phase D)
from __future__ import annotations

from datetime import date
from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

BASE = Path(r"D:\Downloads\Coursework\SCI701_Research_Pathway\Project3_JML_Identity_Lifecycle")
OUT = BASE / "06_Dissertation_Drafts"
TODAY = date.today().isoformat()
STUDENT = "Md Shohrab Hossen Niaz"
SID = ""
DOMAIN = "niazkhanghostgmail.onmicrosoft.com"
TENANT = "a76a68a8-6243-4b0e-b11c-98e815e91820"


def set_style(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)


def h(doc: Document, text: str, level: int = 1) -> None:
    doc.add_heading(text, level=level)


def p(doc: Document, text: str) -> None:
    doc.add_paragraph(text)


def bullets(doc: Document, items: list[str]) -> None:
    for item in items:
        doc.add_paragraph(item, style="List Bullet")


def build_full_dissertation() -> Path:
    doc = Document()
    set_style(doc)

    title = doc.add_heading(
        "Identity Lifecycle Automation and Joiner-Mover-Leaver (JML) Governance for an SME MSP",
        0,
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph(
        f"{STUDENT}\n"
        f"Draft date: {TODAY}\n"
        "Fictional organisation: Coastal Managed IT Pty Ltd (CMIT) — lab-only"
    )
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # --- Ch 1 ---
    h(doc, "1 Introduction")
    h(doc, "1.1 Background", 2)
    p(
        doc,
        "Managed service providers (MSPs) administer IT for many small-business clients and often "
        "hold privileged access across multiple Microsoft tenants and operational tools. Compromise "
        "of a single MSP identity can cascade into many customer environments (Australian Cyber "
        "Security Centre [ACSC], n.d.). Identity lifecycle processes—Joiner, Mover and Leaver "
        "(JML)—are therefore a critical control surface for SME MSPs.",
    )
    h(doc, "1.2 Problem statement", 2)
    p(
        doc,
        "Coastal Managed IT Pty Ltd (CMIT) is a fictional Adelaide MSP (~35 staff) used as the "
        "design scenario for this project. CMIT lacks an automated JML process. Access is "
        "provisioned late, changed inconsistently and deprovisioned slowly, creating orphaned "
        "accounts and standing privileged access with weak audit evidence.",
    )
    h(doc, "1.3 Research questions", 2)
    bullets(
        doc,
        [
            "RQ1: How can JML processes be automated affordably for an SME MSP with multi-tenant responsibilities?",
            "RQ2: What controls ensure timely, auditable deprovisioning of privileged access?",
            "RQ3: How is the effectiveness of lifecycle automation measured?",
        ],
    )
    h(doc, "1.4 Scope and boundaries", 2)
    bullets(
        doc,
        [
            "In scope: as-is/to-be process design; PowerShell/Microsoft Graph prototype on Entra ID Free; governance/audit model; evaluation metrics.",
            "Out of scope: production HR/PSA integration; paid M365 licensing; Entra ID P2 Lifecycle Workflows/PIM/GDAP; real client tenants.",
            "Ethics: fictional CMIT; synthetic users only; isolated lab tenant; no employer/client production data.",
        ],
    )
    h(doc, "1.5 Contribution", 2)
    p(
        doc,
        "This project contributes a design-science artefact: an affordable JML automation prototype "
        "with role packs, JSON audit evidence and orphan-privileged detection, evaluated against "
        "timeliness, coverage and auditability criteria suitable for SME MSP practice.",
    )

    # --- Ch 2 ---
    h(doc, "2 Literature Review")
    h(doc, "2.1 Identity lifecycle and JML", 2)
    p(
        doc,
        "Cameron and Grewe (2022) frame workforce identity through Joiner, Mover and Leaver "
        "processes, emphasising attestation, access removal and audit when identity state changes. "
        "This project adopts that JML framing for CMIT staff and simulated client-portfolio access.",
    )
    h(doc, "2.2 MSP privileged access risk", 2)
    p(
        doc,
        "Tuononen (2023) analyses privileged access management for MSPs that administer multiple "
        "customer environments with highly privileged credentials. ACSC MSP guidance similarly "
        "highlights supplier cyber risk when MSPs hold standing access into customer systems. "
        "These sources justify RQ2’s focus on timely privileged deprovisioning.",
    )
    h(doc, "2.3 Standards and controls", 2)
    p(
        doc,
        "ASD Essential Eight (restrict administrative privileges), ISO/IEC 27001:2022 A.5.15–A.5.18 "
        "(identity and access themes), and Microsoft Graph documentation inform control design. "
        "Entra Lifecycle Workflows are noted as a native alternative that often requires premium "
        "licensing—important for RQ1 affordability.",
    )
    h(doc, "2.4 Design science", 2)
    p(
        doc,
        "Hevner et al. (2004) provide the methodological foundation: build and evaluate an IT "
        "artefact to solve an organisational problem. Evaluation criteria map to RQ3.",
    )

    # --- Ch 3 ---
    h(doc, "3 Methodology")
    p(
        doc,
        "Design-science process used in this project:",
    )
    bullets(
        doc,
        [
            "As-is mapping of CMIT JML processes and risk findings",
            "To-be design of role packs, workflows and evidence model",
            "Prototype implementation in an isolated Microsoft Entra ID Free lab tenant using PowerShell and Microsoft Graph",
            "Evaluation using baseline (estimated manual) versus automated metrics, test cases TC01–TC05, and orphan report",
        ],
    )
    p(
        doc,
        f"Lab tenant: {DOMAIN} (Tenant ID {TENANT}). Cost model: Entra ID Free only—no paid Azure "
        "resources, VMs or M365 Business licenses.",
    )

    # --- Ch 4 ---
    h(doc, "4 As-Is Analysis")
    h(doc, "4.1 Process summary", 2)
    p(
        doc,
        "As-is Joiner depends on delayed PSA/HR records and ad-hoc tickets; groups are assigned "
        "from memory; MFA and client packs are inconsistent. Movers are notified informally; old "
        "client access often remains. Leavers are notified late; privileged roles are removed "
        "manually and inconsistently with weak machine-readable evidence.",
    )
    h(doc, "4.2 Risk findings", 2)
    bullets(
        doc,
        [
            "RF-01 Orphan privileged accounts after leaver (High)",
            "RF-02 Privilege creep on mover (Medium)",
            "RF-03 Delayed joiner provisioning (Medium)",
            "RF-04 Weak audit trail (High)",
            "RF-05 MSP cascade compromise via standing privileged identity (High)",
        ],
    )
    p(
        doc,
        "Lab validation of RF-01: leaver.test02 was deliberately disabled without removing "
        "SG-CMIT-Privileged-Admins membership; the orphan report flagged "
        "DISABLED_BUT_STILL_IN_PRIV_GROUP.",
    )

    # --- Ch 5 ---
    h(doc, "5 Prototype Design")
    h(doc, "5.1 Architecture", 2)
    p(
        doc,
        "Trigger (approved request) → Automation engine (PowerShell/Graph) → Entra actions "
        "(enable/disable, group membership) → Evidence store (JSON + Entra audit logs).",
    )
    h(doc, "5.2 Role packs (frozen)", 2)
    bullets(
        doc,
        [
            "AllStaff: SG-CMIT-AllStaff",
            "ServiceDesk: AllStaff + SG-CMIT-ServiceDesk",
            "Engineer: AllStaff + SG-CMIT-Engineers",
            "ClientA (optional): SG-ClientA-Support",
            "Privileged: SG-CMIT-Privileged-Admins (stripped on leaver; monitored by orphan report)",
        ],
    )
    h(doc, "5.3 Script mapping", 2)
    bullets(
        doc,
        [
            "01_Create_Lab_Groups.ps1 — security groups",
            "02_Create_Synthetic_Users.ps1 — five lab users",
            "03_Invoke_Joiner.ps1 — role-pack assignment",
            "04_Invoke_Mover.ps1 — role-pack diff",
            "05_Invoke_Leaver.ps1 — strip / disable / revoke",
            "06_Orphan_Privileged_Report.ps1 — governance detection",
            "Run-AllLab.ps1 — end-to-end suite",
        ],
    )

    # --- Ch 6 ---
    h(doc, "6 Implementation")
    p(
        doc,
        f"Implementation used tenant {DOMAIN}. Configuration is held in 00_Config.ps1 with "
        "Assert-LabTenant to prevent execution against the wrong directory. Connect-Lab.ps1 "
        "handles Graph authentication (browser or device code).",
    )
    p(
        doc,
        "Synthetic users: joiner.test01, joiner.test02, mover.test01, leaver.test01, leaver.test02. "
        "Evidence JSON files are written under 03_Prototype_Lab/evidence/ with RequestId, Action, "
        "User, Groups and TimestampUtc.",
    )
    p(
        doc,
        "Technical challenge: Graph sign-in reliability in non-interactive shells. Mitigations "
        "included a dedicated connect helper, device-code fallback and CurrentUser context scope.",
    )

    # --- Ch 7 ---
    h(doc, "7 Evaluation")
    h(doc, "7.1 Metrics", 2)
    table = doc.add_table(rows=1, cols=4)
    hdr = table.rows[0].cells
    hdr[0].text = "Metric"
    hdr[1].text = "As-Is (manual)"
    hdr[2].text = "To-Be (automated)"
    hdr[3].text = "Improvement"
    for row in [
        ("Joiner time (min)", "45 (estimated)", "1 (lab)", "~97.8% faster"),
        ("Mover time (min)", "30 (estimated)", "1 (lab)", "~96.7% faster"),
        ("Leaver time (min)", "60 (estimated)", "2 (lab)", "~96.7% faster"),
        ("JML events with evidence", "Partial/email", "8 JSON files", "Machine-readable"),
        ("Orphan privileged detected", "Unknown", "1", "Reportable"),
    ]:
        cells = table.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = v
    h(doc, "7.2 Test cases", 2)
    bullets(
        doc,
        [
            "TC01 Joiner ServiceDesk — Pass",
            "TC02 Joiner Engineer + ClientA — Pass",
            "TC03 Mover SD→Engineer — Pass",
            "TC04 Automated leaver — Pass",
            "TC05 Orphan privileged report — Pass",
        ],
    )
    h(doc, "7.3 Honesty on baselines", 2)
    p(
        doc,
        "Manual baselines are documented estimates of typical MSP effort for comparison, not timed "
        "production tickets. Automated times are from the lab run on 2026-09-01. This is stated as "
        "a limitation under RQ3.",
    )

    # --- Ch 8 ---
    h(doc, "8 Discussion")
    h(doc, "8.1 Answering the research questions", 2)
    p(
        doc,
        "RQ1: Affordable automation is feasible for an SME MSP using Entra ID Free and Graph "
        "scripting with standardised role packs, without enterprise IGA licensing.",
    )
    p(
        doc,
        "RQ2: Timely privileged deprovisioning is achieved by transactional leaver automation "
        "(strip groups, disable, revoke) plus orphan detection for residual risk.",
    )
    p(
        doc,
        "RQ3: Effectiveness can be measured via timeliness, coverage, audit evidence count and "
        "orphan detection, supported by TC01–TC05.",
    )
    h(doc, "8.2 Implications for practice", 2)
    p(
        doc,
        "SME MSPs can start with free-tier Graph automation and governance reports before "
        "investing in premium Lifecycle Workflows or PIM/GDAP. Role packs reduce ad-hoc "
        "over-provisioning; evidence JSON improves audit readiness.",
    )
    h(doc, "8.3 Limitations", 2)
    bullets(
        doc,
        [
            "Lab-only fictional tenant; no HR/PSA integration",
            "Estimated manual baselines",
            "No Entra P2 Lifecycle Workflows / PIM / GDAP",
            "Client access simulated via groups, not real multi-tenant GDAP",
        ],
    )
    h(doc, "8.4 Future work", 2)
    bullets(
        doc,
        [
            "HR/PSA or webhook triggers",
            "Entra Lifecycle Workflows if licensed",
            "Just-in-time privileged access (PIM/GDAP) for real MSP multi-tenant models",
            "Larger evaluation samples and timed baseline protocols",
        ],
    )

    # --- Ch 9 ---
    h(doc, "9 Conclusion")
    p(
        doc,
        "This design-science project designed and evaluated an affordable JML automation "
        "prototype for a fictional SME MSP. The artefact demonstrates measurable reductions in "
        "lifecycle handling time, machine-readable audit evidence, and detection of orphan "
        "privileged accounts. Within ethics and free-tier constraints, the work answers RQ1–RQ3 "
        "and provides a practical starting point for research-informed DevSecOps identity "
        "governance in SME MSP contexts.",
    )

    # References
    h(doc, "References")
    refs = [
        "Australian Cyber Security Centre. (n.d.). How to manage your security when engaging a managed service provider. https://www.cyber.gov.au/",
        "Cameron, A., & Grewe, O. (2022). An overview of the digital identity lifecycle (v2). IDPro Body of Knowledge, 1(7). https://doi.org/10.55621/idpro.31",
        "Hevner, A. R., March, S. T., Park, J., & Ram, S. (2004). Design science in information systems research. MIS Quarterly, 28(1), 75–105.",
        "International Organization for Standardization. (2022). ISO/IEC 27001:2022 Information security, cybersecurity and privacy protection — Information security management systems — Requirements.",
        "Microsoft. (n.d.). Microsoft Graph users overview. https://learn.microsoft.com/graph/api/resources/users-overview",
        "Tuononen, H. (2023). Privileged access management model for a managed service provider (Master’s thesis). JAMK University of Applied Sciences. http://theseus.fi/handle/10024/800216",
    ]
    for r in refs:
        p(doc, r)

    path = OUT / "Project3_Dissertation_Draft.docx"
    doc.save(path)
    return path


def build_chapter_files() -> None:
    """Short chapter stubs for modular editing."""
    chapters = {
        "Ch01_Introduction.docx": (
            "Chapter 1 — Introduction",
            "Background on MSP identity risk; CMIT problem; RQ1–RQ3; scope; contribution. "
            "See full draft for complete text.",
        ),
        "Ch02_Literature.docx": (
            "Chapter 2 — Literature Review",
            "Cameron & Grewe (2022); Tuononen (2023); Essential Eight; ISO A.5.15–18; Hevner (2004).",
        ),
        "Ch03_Methodology.docx": (
            "Chapter 3 — Methodology",
            "Design-science steps; Entra Free lab; PowerShell/Graph; evaluation criteria.",
        ),
        "Ch04_AsIs_Analysis.docx": (
            "Chapter 4 — As-Is Analysis",
            "Joiner/Mover/Leaver failure modes; RF-01 to RF-05; lab orphan validation.",
        ),
        "Ch05_Design.docx": (
            "Chapter 5 — Prototype Design",
            "Architecture; role packs; script mapping 01–06.",
        ),
        "Ch07_Evaluation.docx": (
            "Chapter 7 — Evaluation",
            "Metrics table; TC01–TC05; baseline honesty; orphan finding.",
        ),
        "Ch08_Discussion.docx": (
            "Chapter 8 — Discussion",
            "RQ answers; practice implications; limitations; future work.",
        ),
        "Ch09_Conclusion.docx": (
            "Chapter 9 — Conclusion",
            "Summary of artefact, results and career/research contribution.",
        ),
    }
    for name, (title, body) in chapters.items():
        d = Document()
        set_style(d)
        h(d, title, 0)
        p(d, f"{STUDENT} | {SID} | {TODAY}")
        p(d, body)
        p(d, "Full narrative: Project3_Dissertation_Draft.docx")
        d.save(OUT / name)


def update_checklist() -> None:
    path = BASE / "00_Master_Checklist.txt"
    text = path.read_text(encoding="utf-8")
    block = """
STEP 3 - DISSERTATION WRITE-UP (Phase D)
[x] D1 Introduction + Methodology drafted
[x] D2 Current-state + Design drafted
[x] D3 Implementation + Evaluation drafted
[x] D4 Discussion answering RQ1-RQ3 drafted
[x] D5 References + full dissertation draft
[ ] Final polish / supervisor feedback / Turnitin submission
"""
    if "STEP 3 - DISSERTATION" not in text:
        path.write_text(text.rstrip() + block, encoding="utf-8")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    full = build_full_dissertation()
    build_chapter_files()
    update_checklist()
    # Refresh Ch06 implementation draft pointer
    d6 = Document()
    set_style(d6)
    h(d6, "Chapter 6 — Implementation", 0)
    p(d6, f"{STUDENT} | {SID} | Lab: {DOMAIN}")
    p(
        d6,
        "See Project3_Dissertation_Draft.docx Chapter 6 for the integrated narrative. "
        "Evidence: 03_Prototype_Lab/evidence/; screenshots in 05_Evaluation/screenshots/.",
    )
    d6.save(OUT / "Ch06_Implementation_Draft.docx")
    print(f"Wrote {full}")
    print("Chapter stubs written; checklist updated.")


if __name__ == "__main__":
    main()
