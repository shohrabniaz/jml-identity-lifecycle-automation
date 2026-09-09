PROJECT 3 PROTOTYPE LAB
=======================
Safety: Lab tenant only. Set TENANT_ID before running anything.
Install once (PowerShell 7 recommended):
  Install-Module Microsoft.Graph -Scope CurrentUser

Connect:
  Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All","Directory.ReadWrite.All","AuditLog.Read.All"

Scripts:
  01_Create_Lab_Groups.ps1
  02_Create_Synthetic_Users.ps1
  03_Invoke_Joiner.ps1
  04_Invoke_Mover.ps1
  05_Invoke_Leaver.ps1
  06_Orphan_Privileged_Report.ps1

Run order for first lab day: 01 -> 02 -> 03 -> 05 -> 06
