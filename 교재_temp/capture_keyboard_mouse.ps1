$ErrorActionPreference = 'Stop'

$outDir = Join-Path $PSScriptRoot 'keyboard_mouse_images'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class Win32Capture {
  [StructLayout(LayoutKind.Sequential)]
  public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
'@

function Wait-MainWindow([System.Diagnostics.Process]$process, [int]$seconds = 12) {
  for ($i = 0; $i -lt ($seconds * 5); $i++) {
    $process.Refresh()
    if ($process.MainWindowHandle -ne 0) { return $process.MainWindowHandle }
    Start-Sleep -Milliseconds 200
  }
  return [IntPtr]::Zero
}

function Capture-Window([IntPtr]$handle, [string]$path) {
  if ($handle -eq [IntPtr]::Zero) { throw "창 핸들을 찾지 못했습니다: $path" }
  [Win32Capture]::ShowWindow($handle, 3) | Out-Null
  [Win32Capture]::SetForegroundWindow($handle) | Out-Null
  Start-Sleep -Milliseconds 900
  $r = New-Object Win32Capture+RECT
  [Win32Capture]::GetWindowRect($handle, [ref]$r) | Out-Null
  $w = $r.Right - $r.Left
  $h = $r.Bottom - $r.Top
  if ($w -lt 100 -or $h -lt 100) { throw "캡처할 창 크기가 올바르지 않습니다: $w x $h" }
  $bmp = New-Object System.Drawing.Bitmap($w, $h)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen($r.Left, $r.Top, 0, 0, $bmp.Size)
  $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose()
  $bmp.Dispose()
}

$notepad = Start-Process notepad.exe -ArgumentList (Join-Path $PSScriptRoot 'input_practice.txt') -PassThru
$notepadHandle = Wait-MainWindow $notepad
Capture-Window $notepadHandle (Join-Path $outDir '01_input_notepad.png')

$charmap = Start-Process charmap.exe -PassThru
$charmapHandle = Wait-MainWindow $charmap
Capture-Window $charmapHandle (Join-Path $outDir '02_special_charmap.png')

$osk = Start-Process "$env:WINDIR\System32\osk.exe" -PassThru
$oskHandle = Wait-MainWindow $osk
if ($oskHandle -ne [IntPtr]::Zero) {
  Capture-Window $oskHandle (Join-Path $outDir '03_shortcuts_osk.png')
}

$explorer = Start-Process explorer.exe -ArgumentList $PSScriptRoot -PassThru
Start-Sleep -Seconds 2
$explorerHandle = [IntPtr]::Zero
foreach ($ep in (Get-Process explorer)) {
  $ep.Refresh()
  if ($ep.MainWindowTitle -eq (Split-Path -Leaf $PSScriptRoot)) {
    $explorerHandle = $ep.MainWindowHandle
    break
  }
}
if ($explorerHandle -ne [IntPtr]::Zero) {
  Capture-Window $explorerHandle (Join-Path $outDir '04_mouse_explorer.png')
}

Get-ChildItem -LiteralPath $outDir | Select-Object Name,Length
