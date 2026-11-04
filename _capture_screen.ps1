param([Parameter(Mandatory=$true)][string]$OutFile)
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Fg {
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
 [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow;
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr pid);
 [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool attach);
 [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
}
"@
$target = Get-Process chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like "*Entra*" } | Select-Object -First 1
if (-not $target) { throw "Entra Chrome window not found" }
$hwnd = $target.MainWindowHandle
$fg = [Fg]::GetForegroundWindow
$fgThread = [Fg]::GetWindowThreadProcessId($fg, [IntPtr]::Zero)
$targetThread = [Fg]::GetWindowThreadProcessId($hwnd, [IntPtr]::Zero)
[Fg]::AttachThreadInput($fgThread, $targetThread, $true) | Out-Null
[Fg]::ShowWindow($hwnd, 3) | Out-Null
[Fg]::BringWindowToTop($hwnd) | Out-Null
[Fg]::SetForegroundWindow($hwnd) | Out-Null
[Fg]::AttachThreadInput($fgThread, $targetThread, $false) | Out-Null
Start-Sleep -Milliseconds 1200
$b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bmp = New-Object System.Drawing.Bitmap $b.Width, $b.Height
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($b.X, $b.Y, 0, 0, $b.Size)
$bmp.Save($OutFile, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose; $bmp.Dispose
Write-Output "saved $($target.MainWindowTitle)"
