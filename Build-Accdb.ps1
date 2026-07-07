# ==============================================================
#  Build-Accdb.ps1  -  إنشاء ملف المكاتبات.accdb جاهزاً بالكامل
# --------------------------------------------------------------
#  يتطلب: Windows + Microsoft Access (Office) مثبّت على الجهاز.
#
#  طريقة التشغيل (من داخل مجلد المشروع):
#    powershell -ExecutionPolicy Bypass -File Build-Accdb.ps1
#
#  ماذا يفعل السكربت؟
#    1) يجهّز ملفات الوحدات بصيغة آمنة الترميز (ASCII)
#    2) ينشئ قاعدة بيانات جديدة (المكاتبات.accdb)
#    3) يستورد وحدات VBA تلقائياً
#    4) يشغّل إجراء البناء لإنشاء الجداول والاستعلامات
#       والنماذج والتقرير
#    والنتيجة: ملف .accdb جاهز للاستخدام مباشرة.
# ==============================================================

$ErrorActionPreference = "Stop"

$root    = $PSScriptRoot
$srcDir  = Join-Path $root "src"
$distDir = Join-Path $root "dist"
$dbPath  = Join-Path $root "المكاتبات.accdb"

# ----- 0) فحوصات أولية -----
if (Test-Path $dbPath) {
    Write-Host "الملف موجود مسبقاً: $dbPath" -ForegroundColor Yellow
    Write-Host "احذفه أو انقله ثم أعد تشغيل السكربت (حماية من الكتابة فوق بياناتك)." -ForegroundColor Yellow
    exit 1
}

# ----- 1) تجهيز الوحدات بصيغة آمنة الترميز (ASCII فقط) -----
# النصوص العربية تتحول إلى رموز يونيكود تُفك وقت التشغيل،
# فتعمل بشكل سليم على أي جهاز مهما كانت لغة النظام.
& (Join-Path $root "tools\Convert-Modules.ps1") -SrcDir $srcDir -DstDir $distDir | Out-Null
Write-Host "1/4  تم تجهيز الوحدات (ترميز آمن)" -ForegroundColor Green

# ----- 2) تشغيل Access وإنشاء قاعدة البيانات -----
try {
    $app = New-Object -ComObject Access.Application
} catch {
    Write-Host "لم يتم العثور على Microsoft Access على هذا الجهاز." -ForegroundColor Red
    Write-Host "ثبّت Microsoft Office (بما فيه Access) ثم أعد المحاولة." -ForegroundColor Red
    exit 1
}

try {
    $app.AutomationSecurity = 1      # السماح بتشغيل الأكواد أثناء البناء
    $app.Visible = $true
    $app.NewCurrentDatabase($dbPath)
    Write-Host "2/4  تم إنشاء قاعدة البيانات" -ForegroundColor Green

    # ----- 3) استيراد وحدات VBA -----
    $acModule = 5
    $modules = @("basMain", "basBuildTables", "basBuildQueries",
                 "basBuildForms", "basBuildReport", "basHandlers")
    foreach ($m in $modules) {
        $app.LoadFromText($acModule, $m, (Join-Path $distDir "$m.bas"))
    }
    Write-Host "3/4  تم استيراد وحدات الأكواد" -ForegroundColor Green

    # ----- 4) بناء الجداول والاستعلامات والنماذج والتقرير -----
    $app.Run("BuildAllSilent")
    Write-Host "4/4  تم بناء الجداول والنماذج والاستعلامات والتقرير" -ForegroundColor Green

    $app.CloseCurrentDatabase()
}
finally {
    $app.Quit(1)   # acQuitSaveAll
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}

Write-Host ""
Write-Host "اكتمل البناء بنجاح!" -ForegroundColor Cyan
Write-Host "الملف الجاهز: $dbPath" -ForegroundColor Cyan
Write-Host "افتحه بنقرة مزدوجة وستظهر الواجهة الرئيسية تلقائياً." -ForegroundColor Cyan
