"""Project 3 — submission report (<=2000 words) and 5-minute slides."""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PRgb
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt as PPt

OUT = Path(r"D:\Downloads\Coursework\SCI701_Research_Pathway\Project3_JML_Identity_Lifecycle\08_Submission")
NAVY = RGBColor(0x1F, 0x4E, 0x79)
PN = PRgb(0x1F, 0x4E, 0x79)
GOLD = PRgb(0xC4, 0x59, 0x11)
WHITE = PRgb(0xFF, 0xFF, 0xFF)
SLATE = PRgb(0x33, 0x33, 0x33)

SECTIONS: list[tuple[str, list[str]]] = [
 (
 "Abstract",
 [
 "Managed service providers (MSPs) hold privileged access across many client environments, so weak identity lifecycle control is a high-impact risk. This design-science project examines how a small MSP can automate joiner–mover–leaver (JML) processes without enterprise identity-governance licensing. Using a fictional Adelaide MSP, Coastal Managed IT (CMIT), the study mapped manual JML failure modes, designed role-pack workflows, and prototyped them in a free Microsoft Entra ID tenant with PowerShell and Microsoft Graph. Evaluation compared estimated manual effort with laboratory timings and tested audit evidence and orphan privileged-account detection. Automated joiner and mover events completed in about one minute, and leaver events in about two minutes, against estimated manual baselines of 45, 30 and 60 minutes. Eight machine-readable evidence records were produced, and a deliberate failed leaver was detected by an orphan report. The prototype shows that affordable JML automation is feasible for an SME MSP, provided claims about time savings are bounded by the laboratory design and estimated baselines.",
 ],
 ),
 (
 "1. Introduction",
 [
 "MSPs are attractive targets because a single administrator identity can reach many customer tenants (Australian Cyber Security Centre [ACSC], n.d.). Catalogue Project 3 addresses this risk for an SME of about 35 staff: onboarding, role change and offboarding are manual, so access is granted late, changed inconsistently and removed slowly. The result is orphaned accounts and standing privileged access.",
 "The project treats Coastal Managed IT Pty Ltd (CMIT) as a fictional case. No employer or client production tenant was used. Three questions guide the work. RQ1 asks how JML can be automated affordably when an MSP has multi-tenant responsibilities. RQ2 asks which controls make privileged deprovisioning timely and auditable. RQ3 asks how effectiveness should be measured. The expected outcomes are a working prototype, evidence of faster deprovisioning and fewer undetected orphans, and a governance model suited to a small MSP.",
 ],
 ),
 (
 "2. Literature and standards",
 [
 "Workforce identity is commonly organised as joiner, mover and leaver processes. Cameron and Grewe (2022) describe these as the points at which accounts, group memberships and entitlements are created, changed and removed, and they stress that access must be attested and withdrawn when employment status changes. That framing fits CMIT: the failure is not the absence of accounts, but the absence of a repeatable control at each transition.",
 "MSP privilege raises the stakes. Tuononen (2023) shows that service providers administer customer management planes with highly privileged credentials, so a breach can damage both the customer and the provider’s liability position. ACSC guidance on engaging an MSP makes the same point from the customer side: standing supplier access must be governed, not assumed safe because the supplier is trusted (ACSC, n.d.). ASD’s Essential Eight control to restrict administrative privileges, and ISO/IEC 27001:2022 controls A.5.16 to A.5.18 on identity, authentication information and access rights, supply the control objectives. They do not prescribe a product.",
 "Product choice matters for cost. Microsoft Entra lifecycle workflows can automate joiner and leaver tasks, but they sit in a premium governance tier that many 35-person MSPs will not buy first (Microsoft, n.d.-a). Microsoft Graph can create users, manage group membership and disable accounts on the free directory tier (Microsoft, n.d.-b). Hevner et al. (2004) justify treating that scripting approach as a design-science artefact: build something useful, then evaluate it against explicit criteria rather than only describing the problem. The criteria used here are timeliness, coverage of role changes, and auditability.",
 ],
 ),
 (
 "3. Methodology",
 [
 "The study followed four design-science steps. First, the as-is JML process was mapped for HR/PSA, service desk, identity administration and audit, and risks were rated. Second, a to-be design specified role packs and the order of leaver actions. Third, the design was implemented in an isolated Entra ID Free tenant using PowerShell. Fourth, five test cases compared the prototype with the as-is failure modes.",
 "The laboratory contained eight security groups and five synthetic users. Scripts refused to run unless the connected tenant matched a configured tenant identifier, which limited the chance of changing the wrong directory. Manual timings were not taken from live tickets. They are stated estimates of typical MSP effort—45 minutes to join, 30 to move and 60 to complete a leaver—including ticket handling and inconsistent privileged cleanup. Automated timings are the laboratory run on 1 September 2026. This distinction is deliberate: overstating the baseline would inflate the result.",
 ],
 ),
 (
 "4. As-is risks and to-be design",
 [
 "The as-is joiner path depended on a late personnel record, an ad-hoc ticket and groups chosen from memory. Day-one access was slow, and urgent client access was often added without an expiry. Movers were announced in chat; new groups were added and old client groups left in place, which is privilege creep. Leavers were notified late. Mailboxes might be disabled, but privileged and client-admin roles were removed inconsistently, and the audit trail was email rather than a record of membership changes. Five risks followed: orphan privileged accounts, privilege creep, delayed joiners, weak evidence, and cascade compromise of client environments.",
 "The to-be design replaces memory with role packs. Every joiner receives an all-staff group plus either a service-desk or engineer pack. Client portfolio access is a separate optional group. Privileged administration is a distinct group that leaver processing must remove. A joiner applies the pack and writes evidence. A mover removes both role packs and adds only the new one, which is a simple diff rather than an additive change. A leaver removes privileged, client and role groups, disables the account, revokes sign-in sessions where the API allows it, and writes a JSON record containing a request identifier, actor, timestamp and groups touched. A separate report lists members of the privileged group and flags any disabled account still present. That report is the control for the failure the automation is meant to prevent.",
 "Governance is intentionally small. Each event stores a request identifier, the action, the user, the groups touched and a UTC timestamp. That is enough for a supervisor or an auditor to reconstruct one change without reading a mailbox. Privileged packs are outside the default joiner path, which is a lightweight separation of duties: day-one access and admin access are not the same decision. The design does not integrate a human-resources system, does not assign licences, and does not use partner GDAP or Privileged Identity Management. Those are real MSP needs, but they require either production data or paid licensing. They were left out so the artefact stayed inside the free tenant and the ethics boundary.",
 ],
 ),
 (
 "5. Implementation and evaluation",
 [
 "Implementation produced the groups and five test users, then executed the workflows. A service-desk joiner received the all-staff and service-desk groups. An engineer joiner received all-staff, engineers and a simulated Client A group. A mover was taken from the service-desk pack to the engineer pack. An automated leaver was disabled and stripped of role groups. To test detection rather than only the happy path, a second leaver was disabled but deliberately left in the privileged group, reproducing the as-is mistake.",
 "All five test cases passed. The orphan report returned one finding: the disabled privileged user, classified as still entitled. Eight JSON evidence files recorded group creation, user creation, two joiners, one mover, one leaver, the simulated failure and the orphan report. Automated handling took about one minute for joiner and mover events and about two minutes for the leaver, against the estimates above. The relative reductions are large—about 98% for joiner and about 97% for mover and leaver—but they describe script execution against estimated manual effort, not a controlled time-and-motion study of CMIT staff.",
 "Coverage was demonstrated by the group diffs: the mover did not keep both role packs, and the automated leaver did not remain enabled. Auditability improved from partial email evidence to a file per event. What the run does not show is elapsed business time from resignation to ticket, because no HR system triggered the script. The measured gain is therefore the identity-administration step, which is exactly where orphaned privilege is created. A calendar delay of several days, which the as-is map treats as common, would still dominate end-to-end leaver risk even if the script itself takes two minutes. The prototype removes the technical excuse for that delay. It does not, by itself, force HR to notify IT on the last day.",
 ],
 ),
 (
 "6. Discussion",
 [
 "RQ1 is answered with a qualified yes. An SME MSP can automate JML on Entra ID Free by standardising role packs and calling Graph. That is enough to remove the ad-hoc pattern. It is not a full identity-governance suite: there is no access review campaign, no separation-of-duties policy engine and no just-in-time elevation. For a 35-person firm, starting here is more realistic than buying Entra ID Governance before the process exists. Premium lifecycle workflows remain a later option once the role packs are stable.",
 "RQ2 is answered by ordering and evidence, not by a single disable flag. Disablement alone left the simulated leaver inside the privileged group. The control that matters is the combination of membership removal, session revocation and a detective report. Tuononen’s (2023) argument that MSP privilege must be monitored is visible in that result: prevention failed in the simulation, and detection still worked. ISO-style access-rights removal is only met if both succeed.",
 "RQ3 needs more than a percentage. Timeliness, coverage and auditability are necessary, but the honest metric set also includes residual orphan count. A fast leaver script that does not check the privileged group would look successful and still fail Essential Eight intent. Publishing the estimated nature of the baseline is part of the measurement method. A marker, a client or a future employer should be able to see what was timed and what was assumed.",
 "Two alternative explanations should be kept in view. First, the time reduction may shrink once approvals, licence assignment and endpoint build are included; those steps were out of scope and are often slower than the directory change. Second, a single successful laboratory path does not prove the scripts survive concurrent operators, naming collisions or a half-failed Graph call. Idempotent re-runs were handled for existing users and groups, but there was no fault-injection test. The main limitations are therefore the fictional single tenant, estimated manual baselines, group-based simulation of client access rather than GDAP, and no human-resources trigger. Future work should time a sanitised ticket sample, add approval before privileged packs, and test just-in-time admin only if licensing is available. Those limits do not erase the finding. They bound it: the artefact is a credible first control for an SME MSP, not a completed privileged-access programme.",
 ],
 ),
 (
 "7. Conclusion",
 [
 "CMIT’s manual JML process creates delayed access, privilege creep and orphaned admin rights. A free-tier Graph prototype, built as a design-science artefact, applied role packs for joiners and movers, stripped and disabled leavers, and detected a deliberately orphaned privileged account. Against stated manual estimates, administration time fell from tens of minutes to about one or two minutes, with eight evidence records. The work shows that an SME MSP can improve JML governance before buying premium identity services, if it is willing to measure residual privilege and not only speed.",
 ],
 ),
]

REFS = [
 "Australian Cyber Security Centre. (n.d.). How to manage your security when engaging a managed service provider. https://www.cyber.gov.au/business-government/supplier-cyber-risk-management/managed-service-providers/how-to-manage-your-security-when-engaging-a-managed-service-provider",
 "Australian Signals Directorate. (n.d.). Essential Eight. https://www.cyber.gov.au/resources-business-and-government/essential-cyber-security/essential-eight",
 "Cameron, A., & Grewe, O. (2022). An overview of the digital identity lifecycle (v2). IDPro Body of Knowledge, 1(7). https://doi.org/10.55621/idpro.31",
 "Hevner, A. R., March, S. T., Park, J., & Ram, S. (2004). Design science in information systems research. MIS Quarterly, 28(1), 75–105. https://doi.org/10.2307/25148625",
 "International Organization for Standardization. (2022). ISO/IEC 27001:2022 Information security, cybersecurity and privacy protection — Information security management systems — Requirements.",
 "Microsoft. (n.d.-a). What are lifecycle workflows? https://learn.microsoft.com/entra/id-governance/what-are-lifecycle-workflows",
 "Microsoft. (n.d.-b). Use the Microsoft Graph API. https://learn.microsoft.com/graph/use-the-api",
 "Tuononen, H. (2023). Privileged access management model for a managed service provider (Master’s thesis, JAMK University of Applied Sciences). Theseus. http://theseus.fi/handle/10024/800216",
]


def words(text: str) -> int:
 return len(text.split)


def body_count -> int:
 return sum(words(para) for _, paras in SECTIONS for para in paras)


def shade_cell(cell, hex_color: str) -> None:
 tc = cell._tc
 tcPr = tc.get_or_add_tcPr
 shd = tcPr.makeelement(
 qn("w:shd"),
 {qn("w:fill"): hex_color, qn("w:val"): "clear"},
 )
 tcPr.append(shd)


def build_report(count: int) -> Path:
 doc = Document
 normal = doc.styles["Normal"]
 normal.font.name = "Calibri"
 normal.font.size = Pt(11)
 normal.paragraph_format.space_after = Pt(8)
 normal.paragraph_format.line_spacing = 1.15
 for i in range(1, 4):
 doc.styles[f"Heading {i}"].font.color.rgb = NAVY

 t = doc.add_paragraph
 t.alignment = WD_ALIGN_PARAGRAPH.CENTER
 r = t.add_run(
 "Identity Lifecycle Automation and Joiner–Mover–Leaver Governance\nfor an SME Managed Service Provider"
 )
 r.bold = True
 r.font.size = Pt(16)
 r.font.color.rgb = NAVY
 r.font.name = "Calibri"

 meta = doc.add_paragraph
 meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
 mr = meta.add_run(
 "Md Shohrab Hossen Niaz | \n"
 "\n"
 "\n"
 "Fictional case: Coastal Managed IT Pty Ltd | Laboratory only\n"
 f"Word count (excluding title block and reference list): {count}"
 )
 mr.font.size = Pt(11)
 mr.font.name = "Calibri"

 for title, paras in SECTIONS:
 level = 1 if title != "Abstract" else 1
 doc.add_heading(title, level=level)
 for para in paras:
 doc.add_paragraph(para)

 doc.add_heading("References", level=1)
 for ref in REFS:
 rp = doc.add_paragraph(ref)
 rp.paragraph_format.left_indent = Cm(1.25)
 rp.paragraph_format.first_line_indent = Cm(-1.25)

 path = OUT / "Project3_Report.docx"
 doc.save(path)
 return path


def add_notes(slide, text: str) -> None:
 slide.notes_slide.notes_text_frame.text = text


def set_run(paragraph, text, size, bold=False, color=SLATE, font="Calibri"):
 paragraph.clear
 run = paragraph.add_run
 run.text = text
 run.font.size = PPt(size)
 run.font.bold = bold
 run.font.color.rgb = color
 run.font.name = font
 return run


def add_bullets(shape, lines: list[str]) -> None:
 tf = shape.text_frame
 tf.word_wrap = True
 for i, line in enumerate(lines):
 p = tf.paragraphs[0] if i == 0 else tf.add_paragraph
 p.level = 0
 p.space_after = PPt(10)
 set_run(p, line, 20, color=SLATE)


def build_slides -> Path:
 prs = Presentation
 prs.slide_width = Inches(13.333)
 prs.slide_height = Inches(7.5)
 blank = prs.slide_layouts[6]

 def bar(slide):
 shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
 shape.fill.solid
 shape.fill.fore_color.rgb = GOLD
 shape.line.fill.background

 def footer(slide, n):
 box = slide.shapes.add_textbox(Inches(0.5), Inches(7.1), Inches(10), Inches(0.3))
 p = box.text_frame.paragraphs[0]
 set_run(p, "Niaz | Project 3 | JML for an SME MSP | Lab only", 12, color=PRgb(0x66, 0x66, 0x66))
 nb = slide.shapes.add_textbox(Inches(12.2), Inches(7.1), Inches(0.8), Inches(0.3))
 set_run(nb.text_frame.paragraphs[0], str(n), 12, color=PRgb(0x66, 0x66, 0x66))

 def title_block(slide, text):
 box = slide.shapes.add_textbox(Inches(0.55), Inches(0.35), Inches(12.2), Inches(0.7))
 set_run(box.text_frame.paragraphs[0], text, 32, bold=True, color=PN)

 # 1 title
 s = prs.slides.add_slide(blank)
 bg = s.shapes.add_shape(1, 0, 0, Inches(13.333), Inches(7.5))
 bg.fill.solid
 bg.fill.fore_color.rgb = PN
 bg.line.fill.background
 accent = s.shapes.add_shape(1, 0, Inches(4.55), Inches(13.333), Inches(0.08))
 accent.fill.solid
 accent.fill.fore_color.rgb = GOLD
 accent.line.fill.background
 tb = s.shapes.add_textbox(Inches(0.7), Inches(1.6), Inches(12), Inches(2.4))
 tf = tb.text_frame
 tf.word_wrap = True
 set_run(tf.paragraphs[0], "Identity lifecycle automation for an SME MSP", 36, True, WHITE)
 p2 = tf.add_paragraph
 set_run(p2, "Joiner–mover–leaver governance · a free-tier Entra prototype", 22, color=PRgb(0xD6, 0xE3, 0xF0))
 meta = s.shapes.add_textbox(Inches(0.7), Inches(5.0), Inches(11), Inches(1.4))
 mf = meta.text_frame
 set_run(mf.paragraphs[0], "Md Shohrab Hossen Niaz | ", 20, True, WHITE)
 p3 = mf.add_paragraph
 set_run(p3, " · ", 16, color=PRgb(0xD6, 0xE3, 0xF0))
 p4 = mf.add_paragraph
 set_run(p4, "Fictional case: Coastal Managed IT · laboratory only", 16, color=PRgb(0xD6, 0xE3, 0xF0))
 add_notes(
 s,
 "Good morning. I am Shohrab Niaz, . In the next five minutes I will brief you on my project: identity lifecycle automation for a small managed service provider. The case organisation, Coastal Managed IT, is fictional, about thirty-five staff, and the work was done only in a free Microsoft Entra laboratory. I will cover the risk, the method, what the prototype actually does, and the result I am prepared to defend.",
 )

 # 2 problem
 s = prs.slides.add_slide(blank)
 bar(s)
 title_block(s, "Why this problem")
 body = s.shapes.add_textbox(Inches(0.55), Inches(1.3), Inches(7.6), Inches(5.4))
 add_bullets(
 body,
 [
 "An MSP identity can reach many client tenants.",
 "CMIT (~35 staff) has no automated JML process.",
 "Joiners wait; movers keep old access; leavers stay privileged.",
 "The audit trail is email and memory, not a record.",
 "One missed leaver is a standing admin credential.",
 ],
 )
 card = s.shapes.add_shape(1, Inches(8.5), Inches(1.5), Inches(4.2), Inches(4.6))
 card.fill.solid
 card.fill.fore_color.rgb = PN
 card.line.fill.background
 ct = s.shapes.add_textbox(Inches(8.75), Inches(1.8), Inches(3.8), Inches(4.1))
 ctf = ct.text_frame
 ctf.word_wrap = True
 set_run(ctf.paragraphs[0], "The risk", 18, True, GOLD)
 q = ctf.add_paragraph
 set_run(q, "Compromise of one MSP admin cascades into customer environments.", 20, color=WHITE)
 src = ctf.add_paragraph
 set_run(src, "ACSC MSP guidance; Essential Eight — restrict admin privileges.", 14, color=PRgb(0xD6, 0xE3, 0xF0))
 footer(s, 2)
 add_notes(
 s,
 "Managed service providers are high-value targets. One administrator identity can reach many client tenants, which is why the Australian Cyber Security Centre treats MSP access as a supplier risk, and why Essential Eight says administrative privilege must be restricted. In Coastal Managed IT there is no automated joiner, mover or leaver process. New starters wait for an ad-hoc ticket. People who change role keep their old client groups. People who leave are sometimes disabled, but not always removed from privileged groups. The only record is email. That is how standing admin access survives an offboarding. The prototype is aimed at that specific failure, not at replacing an entire IT service desk.",
 )

 # 3 approach
 s = prs.slides.add_slide(blank)
 bar(s)
 title_block(s, "Questions and method")
 body = s.shapes.add_textbox(Inches(0.55), Inches(1.25), Inches(12.2), Inches(5.5))
 add_bullets(
 body,
 [
 "RQ1 How can a small MSP automate JML without premium licensing?",
 "RQ2 What makes privileged deprovisioning timely and auditable?",
 "RQ3 How should effectiveness be measured, not just described?",
 "Method Design science: map the as-is process, build the artefact, evaluate it.",
 "Choice PowerShell and Microsoft Graph on Entra ID Free — not Lifecycle Workflows.",
 "Boundary Synthetic users only. Scripts check the tenant ID before they run.",
 ],
 )
 footer(s, 3)
 add_notes(
 s,
 "I set three questions. Can a small MSP automate this without a premium licence? What control makes privileged removal both fast and auditable? And how should we measure success without exaggerating it? The method is design science: map the current process, build an artefact, then evaluate it. I deliberately did not use Entra lifecycle workflows. They are a real Microsoft feature, but they usually need paid identity governance. Affordability is the first research question, so the tool was PowerShell and Microsoft Graph on Entra ID Free. Ethics stayed tight: synthetic users only, and every script checks the tenant identifier before it changes anything.",
 )

 # 4 prototype
 s = prs.slides.add_slide(blank)
 bar(s)
 title_block(s, "What the prototype does")
 cols = [
 ("Joiner", "Apply a role pack:\nall-staff, plus service desk\nor engineer. Optional\nclient group."),
 ("Mover", "Remove the old pack,\nthen add the new one.\nDo not leave both."),
 ("Leaver", "Strip privilege and\nclient groups, disable\nthe account, revoke\nsessions, write evidence."),
 ("Detect", "Report any disabled\nuser still in the\nprivileged group."),
 ]
 for i, (head, text) in enumerate(cols):
 left = 0.45 + i * 3.2
 card = s.shapes.add_shape(1, Inches(left), Inches(1.5), Inches(3.0), Inches(4.8))
 card.fill.solid
 card.fill.fore_color.rgb = PRgb(0xF4, 0xF7, 0xFB)
 card.line.color.rgb = PN
 hb = s.shapes.add_textbox(Inches(left + 0.15), Inches(1.7), Inches(2.7), Inches(0.6))
 set_run(hb.text_frame.paragraphs[0], head, 22, True, PN)
 bb = s.shapes.add_textbox(Inches(left + 0.15), Inches(2.5), Inches(2.7), Inches(3.4))
 bb.text_frame.word_wrap = True
 set_run(bb.text_frame.paragraphs[0], text, 16, color=SLATE)
 footer(s, 4)
 add_notes(
 s,
 "The design replaces memory with four actions. A joiner receives a role pack: all staff, plus either service desk or engineer, and only then an optional client group. A mover does not add groups on top of old ones. The script removes both role packs and applies only the new one, which is how privilege creep is stopped. A leaver is the critical path: strip privileged and client groups, disable the account, revoke sessions, and write a JSON evidence file. The fourth box is detection. I tested a bad leaver on purpose. I disabled an account and left it in the privileged group, because a disable flag alone does not remove admin membership. The report is there to catch that residue.",
 )

 # 5 results
 s = prs.slides.add_slide(blank)
 bar(s)
 title_block(s, "Laboratory results")
 rows = [
 ("Event", "Manual estimate", "Lab run"),
 ("Joiner", "45 min", "~1 min"),
 ("Mover", "30 min", "~1 min"),
 ("Leaver", "60 min", "~2 min"),
 ("Evidence records", "Email / partial", "8 JSON files"),
 ("Orphan privileged account", "Not detected", "1 found, by design"),
 ("Test cases TC01–TC05", "—", "All passed"),
 ]
 table_shape = s.shapes.add_table(7, 3, Inches(0.7), Inches(1.4), Inches(12), Inches(4.6))
 table = table_shape.table
 table.columns[0].width = Inches(4.2)
 table.columns[1].width = Inches(3.9)
 table.columns[2].width = Inches(3.9)
 for r, row in enumerate(rows):
 for c, val in enumerate(row):
 cell = table.cell(r, c)
 cell.text = val
 for p in cell.text_frame.paragraphs:
 p.alignment = PP_ALIGN.LEFT
 for run in p.runs:
 run.font.size = PPt(18)
 run.font.name = "Calibri"
 run.font.bold = r == 0
 run.font.color.rgb = WHITE if r == 0 else SLATE
 cell.fill.solid
 cell.fill.fore_color.rgb = PN if r == 0 else (PRgb(0xF4, 0xF7, 0xFB) if r % 2 else WHITE)
 note = s.shapes.add_textbox(Inches(0.7), Inches(6.2), Inches(12), Inches(0.6))
 set_run(
 note.text_frame.paragraphs[0],
 "Manual figures are estimates of typical effort, not timed production tickets. Lab figures are the 1 September 2026 run.",
 14,
 color=PRgb(0x66, 0x66, 0x66),
 )
 footer(s, 5)
 add_notes(
 s,
 "These are the numbers I will defend, including the caveat. Against my estimates of typical manual effort — forty-five minutes to join, thirty to move, and sixty to finish a leaver — the laboratory run took about one minute, one minute, and two minutes. I am not claiming a stopwatch study of a live service desk. The manual figures are estimates, and I say that in the report. What was measured is the directory change. All five test cases passed. Eight JSON files record who was changed and which groups moved. The result I care about most is the last row. One disabled user was still in the privileged group, and the report found them. Speed without that check would have hidden the original risk.",
 )

 # 6 close
 s = prs.slides.add_slide(blank)
 bar(s)
 title_block(s, "What I conclude")
 body = s.shapes.add_textbox(Inches(0.55), Inches(1.3), Inches(12.2), Inches(4.4))
 add_bullets(
 body,
 [
 "A small MSP can start JML automation on the free directory tier.",
 "Speed is not enough: removal plus an orphan check is the control.",
 "Limits: one lab tenant, no HR trigger, no GDAP, estimated baselines.",
 "Next: sanitised ticket timings, then just-in-time admin only if licensed.",
 ],
 )
 end = s.shapes.add_textbox(Inches(0.55), Inches(5.9), Inches(12), Inches(0.8))
 set_run(end.text_frame.paragraphs[0], "Questions", 28, True, GOLD)
 footer(s, 6)
 add_notes(
 s,
 "My conclusion is narrow, and that is intentional. A small MSP can start joiner, mover and leaver automation on the free directory tier, before it buys a premium identity product. The control that matters is not only speed, and not only disabling the account. It is removing privilege and detecting anything left behind. I do not claim this is production GDAP, an HR integration, or a finished privileged-access programme. Those are the next steps, and only if the licence and the data boundary allow them. I am happy to take questions.",
 )

 path = OUT / "Project3_Presentation.pptx"
 prs.save(path)
 return path


def main -> None:
 OUT.mkdir(parents=True, exist_ok=True)
 count = body_count
 print(f"BODY_WORDS {count}")
 if count > 2000:
 raise SystemExit(f"Over limit: {count}")
 report = build_report(count)
 slides = build_slides
 print(report)
 print(slides)


if __name__ == "__main__":
 main
