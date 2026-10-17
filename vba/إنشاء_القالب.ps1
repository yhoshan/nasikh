# ============================================================
#  إنشاء قالب  ناسخ.dotm  تلقائيًا من ملف الماكرو ناسخ.bas
#  (يُشغَّل مرة واحدة على جهاز ويندوز فيه Word — ثم تنشر الـ .dotm)
#
#  التشغيل: انقر يمينًا على هذا الملف ← Run with PowerShell
#  أو من PowerShell:   powershell -ExecutionPolicy Bypass -File إنشاء_القالب.ps1
# ============================================================

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$bas  = Join-Path $here "ناسخ.bas"
$out  = Join-Path $here "ناسخ.dotm"

if (-not (Test-Path $bas)) {
    Write-Host "لم يُعثر على ملف الماكرو: $bas" -ForegroundColor Red
    exit 1
}

Write-Host "جارٍ فتح Word لإنشاء القالب..." -ForegroundColor Cyan
$word = New-Object -ComObject Word.Application
$word.Visible = $false

try {
    # أنشئ قالبًا جديدًا
    $doc = $word.Documents.Add()
    # استورد وحدة الماكرو
    $vbProj = $doc.VBProject
    $vbProj.VBComponents.Import($bas) | Out-Null

    # احفظ كقالب ممكّن للماكرو (.dotm = نوع 15)
    $wdFormatXMLTemplateMacroEnabled = 15
    $doc.SaveAs2($out, $wdFormatXMLTemplateMacroEnabled)
    $doc.Close()
    Write-Host "تم إنشاء القالب بنجاح:" -ForegroundColor Green
    Write-Host "  $out" -ForegroundColor Green
    Write-Host ""
    Write-Host "انشر هذا الملف (ناسخ.dotm) على تليجرام/تويتر." -ForegroundColor Yellow
    Write-Host "واطلب من المستخدمين اتباع (طريقة_التثبيت.txt)." -ForegroundColor Yellow
}
catch {
    Write-Host "خطأ: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "تأكد من تفعيل الوصول إلى نموذج VBA في Word:" -ForegroundColor Yellow
    Write-Host "  File > Options > Trust Center > Trust Center Settings" -ForegroundColor Yellow
    Write-Host "  > Macro Settings > فعّل: Trust access to the VBA project object model" -ForegroundColor Yellow
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
