# PPT의 각 장을 PNG로 내보낸다 (폰용 대본 PDF의 슬라이드 그림용).
# 사용: powershell -ExecutionPolicy Bypass -File render_png.ps1 -Pptx <pptx 경로> -OutDir <폴더>
# PowerPoint가 설치돼 있어야 한다. 사용자가 PowerPoint를 쓰는 중이면 끄지 않는다.
param([string]$Pptx, [string]$OutDir)
$Pptx = (Resolve-Path $Pptx).Path
if (Test-Path $OutDir) { Remove-Item -Recurse -Force $OutDir -Confirm:$false }
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path
$app = New-Object -ComObject PowerPoint.Application
$userHadOpen = $app.Presentations.Count -gt 0
$pres = $null
try {
    $pres = $app.Presentations.Open($Pptx, $true, $false, $false)
    $i = 1
    foreach ($slide in $pres.Slides) {
        $slide.Export((Join-Path $OutDir ("slide{0:D2}.png" -f $i)), "PNG", 1600, 900)
        $i++
    }
} finally {
    if ($pres) { $pres.Close() }
    if (-not $userHadOpen) { $app.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
Write-Output "saved $($i - 1) slides to $OutDir"
