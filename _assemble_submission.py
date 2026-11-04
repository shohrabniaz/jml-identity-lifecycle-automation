"""Assemble a clean Canvas/Turnitin hand-in folder and zip."""
from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor

ROOT = Path(r"D:\Downloads\Coursework\SCI701_Research_Pathway\Project3_JML_Identity_Lifecycle")
SRC = ROOT / "08_Submission"
HANDIN = ROOT / "Submission_Ready"
SPEAKER = ROOT / "Speaker_Only"

NAVY = RGBColor(0x1F, 0x4E, 0x79)


def reset(path: Path) -> None:
 if path.exists:
 shutil.rmtree(path)
 path.mkdir(parents=True)


def copy(src: Path, dest: Path) -> None:
 dest.parent.mkdir(parents=True, exist_ok=True)
 shutil.copy2(src, dest)


def contents -> None:
 doc = Document
 style = doc.styles["Normal"]
 style.font.name = "Calibri"
 style.font.size = Pt(11)
 title = doc.add_heading("Submission contents", 0)
 title.runs[0].font.color.rgb = NAVY
 doc.add_paragraph(
 "Md Shohrab Hossen Niaz | \n"
 " | Project 3\n"
 "Identity Lifecycle Automation and Joiner–Mover–Leaver Governance for an SME MSP\n"
 "Fictional case: Coastal Managed IT Pty Ltd | Laboratory only"
 )
 doc.add_heading("Upload these files", level=1)
 rows = [
 ("01_Project3_Report.docx", "Written report. 1,808 words excluding the title block and reference list (limit 2,000)."),
 ("02_Project3_Presentation.pptx", "Five-minute presentation. Six slides."),
 ("Appendix/", "Evaluation workbook, laboratory evidence, figures, and prototype scripts. Not required if the portal allows only the report and slides."),
 ]
 for name, note in rows:
 p = doc.add_paragraph
 run = p.add_run(name)
 run.bold = True
 p.add_run(" — " + note)

 doc.add_heading("What the appendix shows", level=1)
 for line in [
 "Evaluation workbook: manual estimates versus laboratory timings, and test cases TC01–TC05.",
 "Evidence JSON: joiner, mover, leaver, failed-leaver simulation, and the orphan privileged-account report.",
 "Figures: users, groups, memberships, disabled leaver, privileged group, and orphan report.",
 "Scripts: the PowerShell prototype. No passwords are included.",
 ]:
 doc.add_paragraph(line, style="List Bullet")

 doc.add_heading("Boundaries", level=1)
 doc.add_paragraph(
 "The organisation is fictional. Users are synthetic. The manual timings are estimates of typical effort, "
 "not a timed production study. The laboratory used Microsoft Entra ID Free only."
 )
 doc.save(HANDIN / "00_Read_Me_First.docx")


def main -> None:
 reset(HANDIN)
 reset(SPEAKER)

 copy(SRC / "Project3_Report.docx", HANDIN / "01_Project3_Report.docx")
 copy(SRC / "Project3_Presentation.pptx", HANDIN / "02_Project3_Presentation.pptx")
 copy(SRC / "Project3_Cue_Card.docx", SPEAKER / "Cue_Card_Do_Not_Upload.docx")
 copy(SRC / "Project3_Presentation.pptx", SPEAKER / "Presentation_with_speaker_notes.pptx")

 copy(
 ROOT / "05_Evaluation" / "Evaluation_Workbook.xlsx",
 HANDIN / "Appendix" / "A_Evaluation" / "Evaluation_Workbook.xlsx",
 )
 evidence = HANDIN / "Appendix" / "B_Evidence"
 for path in sorted((ROOT / "03_Prototype_Lab" / "evidence").glob("*.json")):
 if "CreateGroups" in path.name:
 continue
 copy(path, evidence / path.name)
 figures = HANDIN / "Appendix" / "C_Figures"
 for path in sorted((ROOT / "05_Evaluation" / "screenshots").glob("*.png")):
 copy(path, figures / path.name)
 scripts = HANDIN / "Appendix" / "D_Prototype_Scripts"
 for name in [
 "00_Config.ps1",
 "01_Create_Lab_Groups.ps1",
 "02_Create_Synthetic_Users.ps1",
 "03_Invoke_Joiner.ps1",
 "04_Invoke_Mover.ps1",
 "05_Invoke_Leaver.ps1",
 "06_Orphan_Privileged_Report.ps1",
 ]:
 copy(ROOT / "03_Prototype_Lab" / name, scripts / name)

 contents
 print(HANDIN)


if __name__ == "__main__":
 main
