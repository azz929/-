# ==============================================================
#  أداة تحويل الترميز - Convert-To-1256.ps1
# --------------------------------------------------------------
#  محرر VBA في Access يقرأ ملفات .bas بترميز النظام (ANSI)،
#  بينما ملفات المشروع محفوظة بترميز UTF-8 لعرضها بشكل سليم
#  على GitHub. هذا السكربت يحوّل الملفات إلى ترميز Windows-1256
#  (العربية) داخل مجلد dist لاستيرادها في Access دون تشويه النصوص.
#
#  طريقة التشغيل (من داخل مجلد المشروع):
#    powershell -ExecutionPolicy Bypass -File tools\Convert-To-1256.ps1
# ==============================================================

$src = Join-Path $PSScriptRoot "..\src"
$dst = Join-Path $PSScriptRoot "..\dist"

New-Item -ItemType Directory -Force -Path $dst | Out-Null

$enc1256 = [System.Text.Encoding]::GetEncoding(1256)

Get-ChildItem -Path $src -Filter *.bas | ForEach-Object {
    $text = [System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::UTF8)
    $out  = Join-Path $dst $_.Name
    [System.IO.File]::WriteAllText($out, $text, $enc1256)
    Write-Host ("OK  " + $_.Name)
}

Write-Host ""
Write-Host "تم التحويل. استورد الملفات من مجلد dist داخل محرر VBA في Access."
