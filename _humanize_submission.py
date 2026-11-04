"""Humanise the submission report and build proof images from real lab files."""
from __future__ import annotations

import json
import shutil
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor as PRgb
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt as PPt

ROOT = Path(r"D:\Downloads\Coursework\SCI701_Research_Pathway\Project3_JML_Identity_Lifecycle")
HAND = ROOT / "Submission_Ready"
EVID = ROOT / "03_Prototype_Lab" / "evidence"
SHOTS = ROOT / "05_Evaluation" / "screenshots"
NAVY = RGBColor(0x1F, 0x4E, 0x79)

SECTIONS: list[tuple[str, list[str]]] = [
 (
 "Abstract",
 [
 "I looked at whether a small managed service provider can automate joiner, mover and leaver access without buying a full identity-governance licence. The company in the scenario, Coastal Managed IT, is fictional and has about 35 staff. I mapped how that work is usually done by hand, then built a PowerShell and Microsoft Graph prototype in a free Microsoft Entra ID tenant. In the laboratory run, a joiner and a mover each took about a minute, and a leaver about two minutes. I had estimated the same jobs by hand at 45, 30 and 60 minutes. The scripts wrote JSON evidence, and a separate report found one disabled account that was still in the privileged group. Those time figures are the script run against my estimates. They are not a stopwatch study of a live service desk.",
 ],
 ),
 (
 "1. Introduction",
 [
 "A managed service provider is a useful target. One administrator account can reach many customer tenants, which is why the Australian Cyber Security Centre treats MSP access as a supplier risk (ACSC, n.d.). Project 3 asks what happens when a firm of about 35 people still does joiner, mover and leaver work manually. Access turns up late, role changes are inconsistent, and people who have left can keep privileged access.",
 "I used Coastal Managed IT Pty Ltd as a fictional case only. I did not touch an employer tenant or a real client tenant. The questions I worked from were: how can this be automated at a cost a small MSP can justify; what actually makes privileged removal timely and checkable; and how should I measure that, rather than only describe it?",
 ],
 ),
 (
 "2. What I used from the literature",
 [
 "Cameron and Grewe (2022) describe workforce identity as joiner, mover and leaver. Accounts and group memberships are created, changed, and then removed. The weak point at Coastal Managed IT is not that accounts cannot be created. It is that nobody repeats the same check at each of those three moments.",
 "Tuononen (2023) looked at privileged access for a service provider that administers customer environments. A mistake there is not only a technical fault. It can become a liability problem. ACSC guidance says the same thing from the customer side: standing supplier access has to be governed (ACSC, n.d.). The Essential Eight control on restricting administrative privileges, and ISO/IEC 27001:2022 controls A.5.16 to A.5.18, tell me what the outcome should be. They do not tell me which product to buy.",
 "Microsoft can automate joiner and leaver work with Entra lifecycle workflows, but that feature sits in a paid governance tier (Microsoft, n.d.-a). Graph can create users, change groups and disable accounts on the free directory (Microsoft, n.d.-b). I followed Hevner et al. (2004) in a plain way: build something a small MSP could actually run, then test it on time, on whether the right groups changed, and on whether I could show the evidence afterwards.",
 ],
 ),
 (
 "3. How I did the work",
 [
 "I did this in four steps. I wrote down the manual process and the risks. I designed role packs and the order a leaver should follow. I built that in an isolated Entra ID Free tenant with PowerShell. I then ran five test cases.",
 "The lab had security groups for all staff, service desk, engineers, privileged admins, and two pretend client groups, plus five synthetic users: joiner.test01, joiner.test02, mover.test01, leaver.test01 and leaver.test02. The tenant was niazkhanghostgmail.onmicrosoft.com. Each script checked the tenant ID before it changed anything, so I would not hit the wrong directory. I did not time live tickets. The manual numbers — 45 minutes to join, 30 to move, 60 to finish a leaver — are my estimates of ticket handling and messy privileged cleanup. The automated numbers are from the run on 1 September 2026. I am keeping that split visible so the result is not inflated. I come from DevOps work, so a script with a file at the end was a better fit for me than clicking through the portal and hoping I remembered the groups.",
 ],
 ),
 (
 "4. What was going wrong, and what I designed",
 [
 "In the manual joiner path, the personnel record is late, the ticket is ad hoc, and someone picks groups from memory. Day-one access is slow, and urgent client access gets added with no end date. A mover is often a chat message: new groups go on, old client groups stay. That is privilege creep. A leaver is often notified late. The mailbox might be switched off, but privileged groups are not always removed, and the only record is email. The risks I cared about were orphaned admin accounts, privilege creep, slow joiners, weak evidence, and one MSP identity reaching many clients.",
 "The design uses role packs instead of memory. A joiner gets the all-staff group and either service desk or engineer. Client access is a separate group. Privileged admin is not part of the normal joiner pack. A mover removes both role packs and adds only the new one. A leaver is supposed to strip privileged, client and role groups, disable the account, revoke sessions where Graph allows it, and write a JSON file with a request ID, the user, the time and the groups. A second script lists the privileged group and flags any disabled account still in it. That second script is there because switching an account off does not, by itself, take the person out of an admin group.",
 "I kept the governance small on purpose. Day-one access and admin access are different decisions. I did not connect a human-resources system, assign licences, or use GDAP or Privileged Identity Management. A real MSP needs those. They also need either production data or a paid licence, and both were outside this project.",
 ],
 ),
 (
 "5. What the laboratory run showed",
 [
 "I created the groups and five test users, then ran the cases. joiner.test01 received SG-CMIT-AllStaff and SG-CMIT-ServiceDesk (request JOIN-20260901012528). joiner.test02 received the engineer pack and the Client A group. mover.test01 was moved to the engineer pack. leaver.test01 was disabled. For the last case I disabled leaver.test02 and left that account in SG-CMIT-Privileged-Admins, which is the manual mistake I was trying to catch.",
 "The five test cases passed in the sense I defined them. The orphan report returned one row: Leaver Test02, account disabled, risk DISABLED_BUT_STILL_IN_PRIV_GROUP. The JSON files from that run are in the appendix. Joiner and mover handling took about a minute each, and the leaver about two minutes, against the estimates above. That is roughly a 97 to 98 percent reduction, but only for the directory step I scripted.",
 "One result needs to be read carefully. The leaver record for leaver.test01 shows the account disabled, and it also shows that the user was not a member of the privileged or client groups at that moment. So I can claim the account was disabled. I should not claim that this particular user was stripped of a long list of groups they did not hold. The stronger proof of the control is the orphan report, not that one leaver file. There is also no clock time from a resignation to the script, because nothing in HR starts it. A late notice from HR would still dominate the real delay. The script removes the excuse that the directory change itself takes an hour. It does not make HR send the notice on the last day.",
 ],
 ),
 (
 "6. Discussion",
 [
 "On the first question, a small MSP can start this on Entra ID Free. Role packs plus Graph are enough to stop the worst ad-hoc pattern. This is not a full identity-governance product. There is no access review campaign and no just-in-time elevation. For a 35-person firm I would still start here, and look at lifecycle workflows later if the packs settle and the licence is justified (Microsoft, n.d.-a).",
 "On the second question, disabling the account is not the control. leaver.test02 was disabled and still privileged. The useful control is removal of the group, plus a report that finds anyone left behind. That matches Tuononen’s (2023) point that MSP privilege has to be monitored, and it is what Essential Eight is aiming at when it says administrative access should be restricted.",
 "On the third question, a percentage is not enough. I needed time, the right group changes, a file I can show, and a count of leftover privileged accounts. A fast script that never checks the privileged group would look good and still miss the original risk. Saying that the manual times are estimates is part of the result, not a footnote I hope nobody reads.",
 "Two limits matter. The time saving will look smaller once approvals, licences and a laptop build are included, because I did not automate those. And one clean laboratory path does not prove the scripts cope with two people running them at once, or with Graph failing halfway. The other limits are the fictional tenant, client access modelled as groups rather than GDAP, and no HR trigger. Next I would time a sanitised sample of real tickets, and only then look at just-in-time admin if a licence exists. This is a first control, not a finished privileged-access programme.",
 ],
 ),
 (
 "7. Conclusion",
 [
 "Manual joiner, mover and leaver work at a small MSP leaves access late, messy, and sometimes still privileged after someone has gone. I built a free-tier Graph prototype that applies role packs, disables a leaver, writes a JSON record, and finds a disabled account left in the privileged group. Against my own estimates, the directory step fell from tens of minutes to about one or two minutes. A small MSP can improve this before it buys a premium identity product, as long as it measures who is still privileged and does not only measure speed.",
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


def font(size: int, bold: bool = False):
 path = Path(r"C:\Windows\Fonts\consola.ttf")
 bold_path = Path(r"C:\Windows\Fonts\consolab.ttf")
 chosen = bold_path if bold and bold_path.exists else path
 return ImageFont.truetype(str(chosen), size)


def terminal_png(title: str, body: str, dest: Path) -> None:
 body = body.replace("\t", " ").rstrip + "\n"
 lines = body.splitlines or [""]
 f = font(18)
 fb = font(16, True)
 line_h = 26
 width = 1680
 height = 70 + line_h * (len(lines) + 1) + 24
 img = Image.new("RGB", (width, height), "#0c0c0c")
 draw = ImageDraw.Draw(img)
 draw.rectangle((0, 0, width, 42), fill="#1f4e79")
 draw.text((16, 10), title, font=fb, fill="#ffffff")
 y = 58
 for line in lines:
 draw.text((18, y), line[:190], font=f, fill="#d4d4d4")
 y += line_h
 dest.parent.mkdir(parents=True, exist_ok=True)
 img.save(dest)


def build_proof -> None:
 out = HAND / "Appendix" / "C_Laboratory_Proof"
 users = (SHOTS / "01_Entra_Users_List.txt").read_text(encoding="utf-8")
 groups = (SHOTS / "02_Entra_Groups_List.txt").read_text(encoding="utf-8")
 joiner = (SHOTS / "03_Joiner_Group_Memberships.txt").read_text(encoding="utf-8")
 leaver = (SHOTS / "04_Leaver_Disabled_Account.txt").read_text(encoding="utf-8")
 priv = (SHOTS / "05_Privileged_Group_Members.txt").read_text(encoding="utf-8")
 orphan = (
 "DisplayName UPN AccountEnabled Department JobTitle Risk\n"
 "Leaver Test02 leaver.test02@niazkhanghostgmail.onmicrosoft.com False Engineers Privileged Admin DISABLED_BUT_STILL_IN_PRIV_GROUP\n"
 "\n"
 "Source: evidence\\20260901_012535_OrphanPrivilegedReport.json\n"
 "Same finding on the 19:00 re-run the same day."
 )
 join_json = json.dumps(json.loads((EVID / "20260901_012529_Joiner_joiner.test01.json").read_text(encoding="utf-8-sig")), indent=2)
 orphan_json = json.dumps(json.loads((EVID / "20260901_012535_OrphanPrivilegedReport.json").read_text(encoding="utf-8-sig")), indent=2)

 shots = [
 ("01_Lab_users.png", "PowerShell output — synthetic users — 1 September 2026", users),
 ("02_Lab_groups.png", "PowerShell output — security groups — 1 September 2026", groups),
 ("03_Joiner_membership.png", "PowerShell output — joiner.test01 groups — 1 September 2026", joiner),
 ("04_Leaver_disabled.png", "PowerShell output — leaver.test01 disabled — 1 September 2026", leaver),
 ("05_Privileged_members.png", "PowerShell output — privileged group — 1 September 2026", priv),
 ("06_Orphan_report.png", "PowerShell output — orphan privileged report — 1 September 2026", orphan),
 ("07_Joiner_evidence_json.png", "Evidence file 20260901_012529_Joiner_joiner.test01.json", join_json),
 ("08_Orphan_evidence_json.png", "Evidence file 20260901_012535_OrphanPrivilegedReport.json", orphan_json),
 ]
 for name, title, body in shots:
 terminal_png(title, body, out / name)

 note = Document
 note.styles["Normal"].font.name = "Calibri"
 note.styles["Normal"].font.size = Pt(11)
 note.add_heading("Laboratory proof", 0)
 note.add_paragraph(
 "These images are captures of the PowerShell and JSON output from the Entra ID Free laboratory "
 "on 1 September 2026, tenant niazkhanghostgmail.onmicrosoft.com. They are not screenshots of the "
 "Entra portal. I could not sign in to the portal again while preparing this folder, so I have not "
 "invented portal pictures. The same records are in Appendix B as the original JSON files."
 )
 for name, title, _ in shots:
 note.add_paragraph(f"{name} — {title}")
 note.save(out / "What_these_images_are.docx")


def build_report(count: int) -> None:
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
 "Identity lifecycle automation for a small managed service provider\n"
 "Joiner, mover and leaver access without a premium licence"
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
 doc.add_heading(title, level=1)
 for para in paras:
 doc.add_paragraph(para)
 doc.add_heading("References", level=1)
 for ref in REFS:
 rp = doc.add_paragraph(ref)
 rp.paragraph_format.left_indent = Cm(1.25)
 rp.paragraph_format.first_line_indent = Cm(-1.25)
 doc.save(HAND / "01_Project3_Report.docx")


def add_notes(slide, text: str) -> None:
 slide.notes_slide.notes_text_frame.text = text


def set_run(paragraph, text, size, bold=False, color=PRgb(0x33, 0x33, 0x33)):
 paragraph.clear
 run = paragraph.add_run
 run.text = text
 run.font.size = PPt(size)
 run.font.bold = bold
 run.font.color.rgb = color
 run.font.name = "Calibri"


def add_bullets(shape, lines: list[str]) -> None:
 tf = shape.text_frame
 tf.word_wrap = True
 for i, line in enumerate(lines):
 p = tf.paragraphs[0] if i == 0 else tf.add_paragraph
 p.space_after = PPt(10)
 set_run(p, line, 20)


def build_slides -> None:
 prs = Presentation
 prs.slide_width = Inches(13.333)
 prs.slide_height = Inches(7.5)
 blank = prs.slide_layouts[6]
 pn = PRgb(0x1F, 0x4E, 0x79)
 gold = PRgb(0xC4, 0x59, 0x11)
 white = PRgb(0xFF, 0xFF, 0xFF)

 def bar(slide):
 shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
 shape.fill.solid
 shape.fill.fore_color.rgb = gold
 shape.line.fill.background

 def footer(slide, n):
 box = slide.shapes.add_textbox(Inches(0.5), Inches(7.1), Inches(10), Inches(0.3))
 set_run(box.text_frame.paragraphs[0], "Niaz | Project 3 | Lab only", 12, color=PRgb(0x66, 0x66, 0x66))
 nb = slide.shapes.add_textbox(Inches(12.2), Inches(7.1), Inches(0.8), Inches(0.3))
 set_run(nb.text_frame.paragraphs[0], str(n), 12, color=PRgb(0x66, 0x66, 0x66))

 def title_block(slide, text):
 box = slide.shapes.add_textbox(Inches(0.55), Inches(0.35), Inches(12.2), Inches(0.7))
 set_run(box.text_frame.paragraphs[0], text, 32, bold=True, color=pn)

 s = prs.slides.add_slide(blank)
 bg = s.shapes.add_shape(1, 0, 0, Inches(13.333), Inches(7.5))
 bg.fill.solid
 bg.fill.fore_color.rgb = pn
 bg.line.fill.background
 tb = s.shapes.add_textbox(Inches(0.7), Inches(1.8), Inches(12), Inches(3.2))
 tf = tb.text_frame
 tf.word_wrap = True
 set_run(tf.paragraphs[0], "Joiner, mover and leaver access\nfor a small MSP", 36, True, white)
 p2 = tf.add_paragraph
 set_run(p2, "A free Entra lab, not a paid identity product", 22, color=PRgb(0xD6, 0xE3, 0xF0))
 meta = s.shapes.add_textbox(Inches(0.7), Inches(5.2), Inches(11), Inches(1.2))
 set_run(meta.text_frame.paragraphs[0], "Md Shohrab Hossen Niaz | ", 20, True, white)
 add_notes(s, "I’m Shohrab Niaz. This is a short briefing on automating joiner, mover and leaver access for a small managed service provider. The company is fictional. The work was done in a free Entra lab.")

 s = prs.slides.add_slide(blank)
 bar(s)
 title_block(s, "The problem I started from")
 body = s.shapes.add_textbox(Inches(0.55), Inches(1.4), Inches(12), Inches(5))
 add_bullets(body, [
 "One MSP admin account can reach many client tenants.",
 "Joiners wait on an ad-hoc ticket.",
 "Movers keep old access. That is privilege creep.",
 "Leavers get disabled, and sometimes stay privileged.",
 "The record is usually an email, not a file I can check.",
 ])
 footer(s, 2)
 add_notes(s, "The risk is simple. One MSP login can sit inside a lot of customer tenants. In the scenario, new starters wait, people who change role keep old groups, and people who leave are not always taken out of admin groups. I wanted a process a 35-person firm could actually run.")

 s = prs.slides.add_slide(blank)
 bar(s)
 title_block(s, "What I asked, and how I tested it")
 body = s.shapes.add_textbox(Inches(0.55), Inches(1.4), Inches(12), Inches(5))
 add_bullets(body, [
 "Can a small MSP automate this without a premium licence?",
 "What makes privileged removal checkable, not just fast?",
 "How do I measure it without overstating the result?",
 "I mapped the manual process, built the scripts, then ran five tests.",
 "PowerShell and Graph on Entra ID Free. Synthetic users only.",
 ])
 footer(s, 3)
 add_notes(s, "I did not use Entra lifecycle workflows. They usually need a paid licence, and cost was the first question. Every script checks the tenant ID before it changes anything.")

 s = prs.slides.add_slide(blank)
 bar(s)
 title_block(s, "What the scripts do")
 body = s.shapes.add_textbox(Inches(0.55), Inches(1.4), Inches(12), Inches(5))
 add_bullets(body, [
 "Joiner: all-staff, plus service desk or engineer.",
 "Mover: take the old pack off, then add the new one.",
 "Leaver: disable the account and write a JSON record.",
 "Then check the privileged group for anyone disabled but still a member.",
 "I left leaver.test02 in that group on purpose.",
 ])
 footer(s, 4)
 add_notes(s, "Switching an account off does not remove the group. I disabled leaver.test02 and left them in the privileged group so the report had a real miss to find.")

 s = prs.slides.add_slide(blank)
 bar(s)
 title_block(s, "What I can actually show")
 rows = [
 ("Check", "Result"),
 ("Joiner / mover / leaver time", "About 1 min, 1 min and 2 min"),
 ("My manual estimates", "45, 30 and 60 minutes"),
 ("joiner.test01 groups", "All-staff and service desk"),
 ("Orphan found", "leaver.test02, disabled, still privileged"),
 ("Evidence", "JSON files from 1 September 2026"),
 ]
 table = s.shapes.add_table(6, 2, Inches(0.7), Inches(1.45), Inches(12), Inches(4.4)).table
 table.columns[0].width = Inches(5.2)
 table.columns[1].width = Inches(6.8)
 for r, row in enumerate(rows):
 for c, val in enumerate(row):
 cell = table.cell(r, c)
 cell.text = val
 for p in cell.text_frame.paragraphs:
 for run in p.runs:
 run.font.size = PPt(18)
 run.font.name = "Calibri"
 run.font.bold = r == 0
 run.font.color.rgb = white if r == 0 else PRgb(0x33, 0x33, 0x33)
 cell.fill.solid
 cell.fill.fore_color.rgb = pn if r == 0 else white
 note = s.shapes.add_textbox(Inches(0.7), Inches(6.15), Inches(12), Inches(0.7))
 set_run(note.text_frame.paragraphs[0], "Manual times are my estimates. The proof images are the script output, not portal screenshots.", 14, color=PRgb(0x66, 0x66, 0x66))
 footer(s, 5)
 add_notes(s, "Please treat the manual times as estimates. What I measured is the script. The orphan result is the one I would defend first: a disabled account still sitting in the privileged group, and the report finding it.")

 s = prs.slides.add_slide(blank)
 bar(s)
 title_block(s, "Where I landed")
 body = s.shapes.add_textbox(Inches(0.55), Inches(1.4), Inches(12), Inches(4.6))
 add_bullets(body, [
 "A small MSP can start this on the free directory tier.",
 "Speed is not the control. Finding leftover admin access is.",
 "This is not GDAP, not an HR feed, and not a live ticket study.",
 "I am happy to take questions.",
 ])
 footer(s, 6)
 add_notes(s, "That is as far as I am willing to push the claim. The directory step got much faster, and the orphan check worked. I have not built a full privileged-access programme. Questions?")
 dest = HAND / "02_Project3_Presentation.pptx"
 prs.save(dest)
 shutil.copy2(dest, ROOT / "Speaker_Only" / "Presentation_with_speaker_notes.pptx")


def rewrite_readme -> None:
 doc = Document
 doc.styles["Normal"].font.name = "Calibri"
 doc.styles["Normal"].font.size = Pt(11)
 doc.add_heading("Project submission", 0)
 doc.add_paragraph(
 "Md Shohrab Hossen Niaz | \n"
 "Project 3 — joiner, mover and leaver automation for a small MSP\n"
 "Coastal Managed IT is fictional. The laboratory tenant was niazkhanghostgmail.onmicrosoft.com."
 )
 doc.add_heading("In this folder", level=1)
 for line in [
 "01 report. Word count is on the title page. References are extra.",
 "02 presentation. About five minutes.",
 "Appendix A. The metrics workbook.",
 "Appendix B. The original JSON from the 1 September 2026 run.",
 "Appendix C. Pictures of that same PowerShell and JSON output. These are not Entra portal screenshots.",
 "Appendix D. The scripts. No passwords.",
 ]:
 doc.add_paragraph(line, style="List Bullet")
 doc.save(HAND / "00_Read_Me_First.docx")


def main -> None:
 count = sum(words(p) for _, paras in SECTIONS for p in paras)
 print("BODY_WORDS", count)
 if count > 2000:
 raise SystemExit(count)
 build_report(count)
 build_proof
 build_slides
 rewrite_readme
 print("updated", HAND)


if __name__ == "__main__":
 main
