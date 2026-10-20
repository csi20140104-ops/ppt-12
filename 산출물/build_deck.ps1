# 사용법: powershell -NoProfile -File build_deck.ps1 -Config 상걸리.json
# 템플릿 PPT를 복사한 뒤, 설정(JSON)의 문구 치환 / 토지이용표 / 도면·계산서 이미지를 채워 새 PPT를 만든다.
param([Parameter(Mandatory)][string]$Config)
$ErrorActionPreference = "Stop"
$cfg = Get-Content -LiteralPath $Config -Raw -Encoding utf8 | ConvertFrom-Json
$base = Split-Path -Parent (Resolve-Path -LiteralPath $Config)
Copy-Item -LiteralPath $cfg.template -Destination $cfg.output -Force

$pp = New-Object -ComObject PowerPoint.Application
$p = $pp.Presentations.Open($cfg.output, 0, 0, 0)

function ReplaceAll($tr, $find, $repl) {
    $n = 0
    while ($true) {
        $r = $tr.Replace($find, $repl, 0, 0, 0)
        if ($null -eq $r) { break }
        $n++; if ($n -gt 50) { break }
    }
    return $n
}

# 1) 문구 치환 (모든 슬라이드의 글상자·표)
$log = @()
foreach ($rep in $cfg.replacements) {
    $cnt = 0
    foreach ($sl in $p.Slides) {
        foreach ($sh in $sl.Shapes) {
            if ($sh.HasTextFrame -and $sh.TextFrame.HasText) { $cnt += ReplaceAll $sh.TextFrame.TextRange $rep.find $rep.replace }
            elseif ($sh.HasTable) {
                $tb = $sh.Table
                for ($r = 1; $r -le $tb.Rows.Count; $r++) { for ($c = 1; $c -le $tb.Columns.Count; $c++) {
                    $tr = $tb.Cell($r, $c).Shape.TextFrame.TextRange
                    if ($tr.Text.Length -gt 0) { $cnt += ReplaceAll $tr $rep.find $rep.replace }
                } }
            }
        }
    }
    $log += ("replace '{0}' -> '{1}' : {2} hits" -f $rep.find, $rep.replace, $cnt)
}

# 2) 토지이용계획 표 (슬라이드 4, 3행부터 데이터)
if ($cfg.landTable) {
    $tb = ($p.Slides.Item($cfg.landTable.slide).Shapes | Where-Object { $_.HasTable } | Select-Object -First 1).Table
    $row = 3
    foreach ($lr in $cfg.landTable.rows) {
        $tb.Cell($row, 2).Shape.TextFrame.TextRange.Text = $lr.parcel
        $tb.Cell($row, 3).Shape.TextFrame.TextRange.Text = $lr.area
        $tb.Cell($row, 4).Shape.TextFrame.TextRange.Text = $lr.pct
        $row++
    }
    $log += "land table rows: $($cfg.landTable.rows.Count)"
}

# 3) 이미지 교체: 슬라이드의 기존 그림을 지우고 프레임 안에 맞춰 넣음
$fx = 59; $fy = 162; $fw = 613; $fh = 341
foreach ($im in $cfg.images) {
    $sl = $p.Slides.Item($im.slide)
    foreach ($sh in @($sl.Shapes)) { if ($sh.Type -eq 13) { $sh.Delete() } }
    $k = 0
    foreach ($f in $im.files) {
        $slots = @($im.files).Count
        $w = $fw / $slots; $x0 = $fx + $w * $k
        if ($slots -gt 1) { $w = $w - 10 }
        $pic = $sl.Shapes.AddPicture((Join-Path $base $f), 0, -1, 0, 0)
        $r = [Math]::Min($w / $pic.Width, $fh / $pic.Height)
        $pic.LockAspectRatio = -1
        $pic.Width = $pic.Width * $r
        $pic.Left = $x0 + ($w - $pic.Width) / 2
        $pic.Top = $fy + ($fh - $pic.Height) / 2
        $pic.ZOrder(1)
        $pic.Name = "Picture $($k + 2)"
        $k++
    }
}
$log += "images replaced on slides: " + (($cfg.images | ForEach-Object { $_.slide }) -join ",")

$p.Save()
if ($cfg.previewSlides) { foreach ($i in $cfg.previewSlides) { $p.Slides.Item($i).Export((Join-Path $base "preview_s$i.png"), "PNG", 1400, 1050) } }
$p.Close()
$log

# 4) 협의의견 표: 회신 문서 추출 -> JSON -> 표 슬라이드 생성
if ($cfg.consult) {
    $here = Split-Path -Parent $MyInvocation.MyCommand.Path
    $json = Join-Path (Split-Path -Parent $cfg.output) "consult_rows.json"
    $ovArg = @()
    if ($cfg.consult.fixes) {
        $ovPath = Join-Path (Split-Path -Parent $cfg.output) "consult_overrides.json"
        @{ _fixes = $cfg.consult.fixes } | ConvertTo-Json | Set-Content -LiteralPath $ovPath -Encoding utf8
        $ovArg = @($ovPath)
    }
    python -I (Join-Path $here "consult_extract.py") $cfg.consult.folder $json @ovArg
    powershell -NoProfile -File (Join-Path $here "fill_consult.ps1") -Deck $cfg.output -Json $json -Period $cfg.consult.period
}
"saved $($cfg.output)"
