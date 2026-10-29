# Export the report deck to PDF with the installed PowerPoint (ppSaveAsPDF = 32).
# usage: powershell -File export_pdf.ps1
$folder = Split-Path -Parent $PSScriptRoot
$deck = Join-Path $folder "Machine Learning Capstone Project Report.pptx"
$pdf = Join-Path $folder "Machine Learning Capstone Project Report.pdf"
$ppt = New-Object -ComObject PowerPoint.Application
try {
    $pres = $ppt.Presentations.Open($deck, $true, $false, $false)
    $pres.SaveAs($pdf, 32)
    $pres.Close()
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
Write-Output "saved $pdf"
