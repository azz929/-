# ==============================================================
#  Convert-Modules.ps1 - تجهيز وحدات VBA للاستيراد الآمن
# --------------------------------------------------------------
#  يحوّل ملفات src\*.bas إلى ملفات dist\*.bas خالية تماماً من
#  الأحرف غير اللاتينية (ASCII فقط):
#    - كل نص عربي داخل علامتي اقتباس يتحول إلى استدعاء
#      U("....") برموز يونيكود سداسية تُفك وقت التشغيل
#    - التعليقات العربية تُحذف من النسخة المحوّلة
#  وبذلك تُستورد الوحدات بشكل سليم على أي جهاز Windows
#  مهما كانت لغة النظام، بلا أي إعدادات إضافية.
#
#  التشغيل:
#    powershell -ExecutionPolicy Bypass -File tools\Convert-Modules.ps1
# ==============================================================

param(
    [string]$SrcDir,
    [string]$DstDir
)

$ErrorActionPreference = "Stop"

if (-not $SrcDir) { $SrcDir = Join-Path $PSScriptRoot "..\src" }
if (-not $DstDir) { $DstDir = Join-Path $PSScriptRoot "..\dist" }

function Test-HasNonAscii([string]$s) {
    foreach ($c in $s.ToCharArray()) {
        if ([int]$c -gt 127) { return $true }
    }
    return $false
}

function Convert-VbaLine([string]$line) {
    $sb = New-Object System.Text.StringBuilder
    $i = 0
    $n = $line.Length
    while ($i -lt $n) {
        $ch = $line[$i]
        if ($ch -eq "'") {
            # تعليق: يُحذف إذا احتوى أحرفاً غير لاتينية
            $rest = $line.Substring($i)
            if (-not (Test-HasNonAscii $rest)) { [void]$sb.Append($rest) }
            break
        }
        elseif ($ch -eq '"') {
            # نص بين علامتي اقتباس (مع دعم "" كعلامة اقتباس داخلية)
            $j = $i + 1
            $lit = New-Object System.Text.StringBuilder
            while ($j -lt $n) {
                if ($line[$j] -eq '"') {
                    if (($j + 1) -lt $n -and $line[$j + 1] -eq '"') {
                        [void]$lit.Append('"')
                        $j += 2
                    } else {
                        break
                    }
                } else {
                    [void]$lit.Append($line[$j])
                    $j++
                }
            }
            $s = $lit.ToString()
            if (Test-HasNonAscii $s) {
                $codes = ($s.ToCharArray() | ForEach-Object { ([int]$_).ToString("X4") }) -join " "
                [void]$sb.Append('U("' + $codes + '")')
            } else {
                [void]$sb.Append('"' + $s.Replace('"', '""') + '"')
            }
            $i = $j + 1
        }
        else {
            [void]$sb.Append($ch)
            $i++
        }
    }
    return $sb.ToString().TrimEnd()
}

New-Item -ItemType Directory -Force -Path $DstDir | Out-Null

Get-ChildItem -Path $SrcDir -Filter *.bas | ForEach-Object {
    $text  = [System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::UTF8)
    $lines = $text -split "`r`n|`n"
    $out   = ($lines | ForEach-Object { Convert-VbaLine $_ }) -join "`r`n"
    [System.IO.File]::WriteAllText((Join-Path $DstDir $_.Name), $out, [System.Text.Encoding]::ASCII)
    Write-Host ("OK  " + $_.Name)
}

Write-Host ""
Write-Host "Done. Import the files from the 'dist' folder into Access."
