# 발표 자료 마무리: 이징 진입 애니메이션 추가 + Pretendard 포함 저장
#   입력  slides/_draft.pptx   (build_deck.py 결과, 모핑 전환 포함)
#   출력  slides/credit_presentation.pptx
# 도형 이름이 "anim{n}_..." 인 것만 애니메이션을 넣는다. 같은 n은 함께(약간의 시차), n 순서대로 차례로.
param(
    [string]$Draft = (Join-Path $PSScriptRoot "_draft.pptx"),
    [string]$Out = (Join-Path $PSScriptRoot "credit_presentation.pptx")
)
$ASCEND = 39            # 떠오르기: 페이드 + 위로 살짝 이동
$WITH_PREVIOUS = 2
$AFTER_PREVIOUS = 3

$app = New-Object -ComObject PowerPoint.Application
# 사용자가 이미 PowerPoint를 쓰고 있으면 같은 창에 붙는다 → 그때는 PowerPoint를 끄지 않는다
$userHadOpen = $app.Presentations.Count -gt 0
# 결과 파일이 PowerPoint에 열려 있으면 덮어쓸 수 없으니 멈춘다 (사용자 작업 보호)
foreach ($p in $app.Presentations) {
    if ($p.FullName -ieq $Out) { throw "결과 파일이 PowerPoint에 열려 있습니다. 닫은 뒤 다시 실행하세요: $Out" }
}
$pres = $null
try {
    $pres = $app.Presentations.Open($Draft, $true, $false, $false)
    foreach ($slide in $pres.Slides) {
        $items = @()
        foreach ($sh in $slide.Shapes) {
            if ($sh.Name -match '^anim(\d+)_') { $items += [pscustomobject]@{ Order = [int]$Matches[1]; Shape = $sh } }
        }
        $first = $true
        foreach ($grp in ($items | Group-Object Order | Sort-Object { [int]$_.Name })) {
            $k = 0
            foreach ($it in $grp.Group) {
                $trigger = if ($k -eq 0) { $AFTER_PREVIOUS } else { $WITH_PREVIOUS }
                $eff = $slide.TimeLine.MainSequence.AddEffect($it.Shape, $ASCEND, 0, $trigger)
                $eff.Timing.Duration = 0.7
                $eff.Timing.SmoothStart = 0
                $eff.Timing.SmoothEnd = -1          # 끝부분 감속 (ease-out)
                $eff.Timing.TriggerDelayTime = if ($first) { 0.45 } elseif ($k -eq 0) { 0.0 } else { 0.08 * $k }
                $first = $false
                $k++
            }
        }
    }
    # 25 = ppSaveAsDefault? 사용하지 않음. 24 = ppSaveAsOpenXMLPresentation, -1 = 글꼴 포함
    $pres.SaveAs($Out, 24, -1)
    "saved $Out"
} finally {
    if ($pres) { $pres.Close() }          # 내가 연 파일만 닫는다
    if (-not $userHadOpen) { $app.Quit() } # 원래 PowerPoint가 꺼져 있었을 때만 종료
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
