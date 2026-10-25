# JML Identity Lifecycle Automation

**Title:** Identity Lifecycle Automation and Joiner-Mover-Leaver (JML) Governance for an SME MSP  
**Scenario:** Fictional MSP — Coastal Managed IT Pty Ltd (CMIT) — **lab-only**

## Research questions

1. How can JML processes be automated affordably for an SME MSP with multi-tenant responsibilities?
2. What controls ensure timely, auditable deprovisioning of privileged access?
3. How is the effectiveness of lifecycle automation measured?

## Methodology

Design-science (Hevner et al., 2004): as-is mapping → to-be role packs → PowerShell/Microsoft Graph prototype on **Entra ID Free** → evaluation (timeliness, coverage, auditability).

## Lab results (2026-09-01)

| Metric | Manual (est.) | Automated (lab) |
|--------|---------------|-----------------|
| Joiner | 45 min | 1 min |
| Mover | 30 min | 1 min |
| Leaver | 60 min | 2 min |
| Evidence JSON files | — | 8 |
| Orphan privileged detected | — | 1 |
| Test cases TC01–TC05 | — | Pass |

**Cost model:** Entra ID Free only — no paid Azure VMs, no Pay-as-you-go, no paid M365 licenses.

## Repository layout

```
01_Process_Maps/          As-is / to-be / risk register
02_Literature/            Standards & annotated sources
03_Prototype_Lab/         PowerShell scripts + evidence JSON
04_Governance_Audit/      Governance & audit model
05_Evaluation/            Workbook + screenshots
06_Dissertation_Drafts/   Full dissertation draft + chapters
07_Lab_Tenant_Setup/      Setup notes (credentials excluded from git)
```

## Prototype scripts (`03_Prototype_Lab`)

1. `01_Create_Lab_Groups.ps1`
2. `02_Create_Synthetic_Users.ps1`
3. `03_Invoke_Joiner.ps1`
4. `04_Invoke_Mover.ps1`
5. `05_Invoke_Leaver.ps1`
6. `06_Orphan_Privileged_Report.ps1`
7. `Run-AllLab.ps1` / `CLICK_TO_RUN_LAB.bat`

Configure `00_Config.ps1` with your lab TenantId and domain before running. Never commit real passwords or private credential files.

## Dissertation

Open `06_Dissertation_Drafts/Project3_Dissertation_Draft.docx` for the full draft (Chapters 1–9 + references).

## Ethics

- Fictional organisation only  
- Synthetic lab users only  
- No production employer/client data  
- Tenant-ID guardrails in scripts  

## Licence / academic use

Laboratory artefact. Not a production IAM product.
