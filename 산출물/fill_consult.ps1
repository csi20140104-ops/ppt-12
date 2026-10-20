# 협의의견 표 슬라이드 자동 생성
# 사용: powershell -NoProfile -File fill_consult.ps1 -Deck 덱.pptx -Json consult_rows.json -Period "2026년04월06일~2026년05월18일" [-Submitted "개발행위허가 신청"]
# - 덱에서 '협의기관' 머리말이 있는 표 슬라이드(첫 장=요약 포함, 둘째 장=표만)를 서식 원본으로 삼아 페이지를 다시 만든다.
# - 행은 PowerPoint가 계산한 실제 행 높이로 쪽 나눔한다(넘치면 다음 장).
param(
    [Parameter(Mandatory)][string]$Deck,
    [Parameter(Mandatory)][string]$Json,
    [string]$Period = "",
    [int]$ReplyCount = -1
)
$ErrorActionPreference = "Stop"
$data = Get-Content -LiteralPath $Json -Raw -Encoding utf8 | ConvertFrom-Json
$rows = @($data.rows)
if ($ReplyCount -lt 0) { $ReplyCount = $rows.Count + @($data.skipped | Where-Object { $_[1] -like "별도의견*" }).Count }
$BOTTOM = 498   # 표 하단 한계(pt). 쪽번호·의안번호 위

$pp = New-Object -ComObject PowerPoint.Application
$p = $pp.Presentations.Open($Deck, 0, 0, 0)

function Get-Table($sl) { foreach ($sh in $sl.Shapes) { if ($sh.HasTable) { return $sh } } return $null }

# 협의 표 슬라이드 찾기
$idx = @()
for ($i = 1; $i -le $p.Slides.Count; $i++) {
    $t = Get-Table $p.Slides.Item($i)
    if ($t -and $t.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text -like "협의기관*") { $idx += $i }
}
if ($idx.Count -lt 2) { throw "협의 표 슬라이드(요약+표, 표)를 2장 이상 찾지 못함: $($idx -join ',')" }
$first = $idx[0]; $last = $idx[-1]
"기존 협의 슬라이드: $($idx -join ',')"

function Set-Para($tr2, $text, $kind) {
    # kind: item(○ 항목) / sub(가. 항목)
}

function Fill-Cell($cell, [string[]]$paras, $subIdx) {
    $tf = $cell.Shape.TextFrame
    $tf.TextRange.Text = ($paras -join "`r")
    $tf.TextRange.Font.Bold = 0
    $tf2 = $cell.Shape.TextFrame2
    for ($k = 1; $k -le $paras.Count; $k++) {
        $pf = $tf2.TextRange.Paragraphs($k).ParagraphFormat
        if ($subIdx -contains ($k - 1)) { $pf.LeftIndent = 24; $pf.FirstLineIndent = -13 }
        else { $pf.LeftIndent = 12; $pf.FirstLineIndent = -12 }
        $pf.SpaceWithin = 1.0
    }
}

# 의견 항목 -> 문단 목록 ("○ 본문", "가. 하위")
function To-Paras($items) {
    $paras = New-Object System.Collections.ArrayList
    $sub = New-Object System.Collections.ArrayList
    foreach ($it in $items) {
        $parts = [regex]::Split($it, "\s+(?=[가-하]\.\s)")
        $head = $parts[0] -replace "^[가-하]\.\s*", ""
        [void]$paras.Add("○ " + $head.Trim())
        for ($j = 1; $j -lt $parts.Count; $j++) { [void]$paras.Add($parts[$j].Trim()); [void]$sub.Add($paras.Count - 1) }
    }
    return @{ paras = @($paras); sub = @($sub) }
}

# 서식 원본 슬라이드 복제 (새 페이지는 마지막 협의 슬라이드 뒤에 쌓고, 원본은 나중에 삭제)
$tmpl1 = $p.Slides.Item($first)
$tmpl2 = $p.Slides.Item($idx[1])
$origCount = $idx.Count
$insertAt = $last
$pages = @()

function New-Page($tmpl, $pos) {
    $dup = $tmpl.Duplicate()
    $dup.MoveTo($pos)
    $tbs = Get-Table $dup
    $tb = $tbs.Table
    while ($tb.Rows.Count -gt 2) { $tb.Rows.Item($tb.Rows.Count).Delete() }   # 머리행 + 서식행만 남김
    $tb.Cell(2, 1).Shape.TextFrame.TextRange.Text = ""
    $tb.Cell(2, 2).Shape.TextFrame.TextRange.Text = ""
    $tb.Cell(2, 3).Shape.TextFrame.TextRange.Text = ""
    $tb.Cell(2, 4).Shape.TextFrame.TextRange.Text = ""
    return $dup
}

function Table-Height($tb) { $h = 0; for ($r = 1; $r -le $tb.Rows.Count; $r++) { $h += $tb.Rows.Item($r).Height }; return $h }

function Put-Row($tb, $rowNo, $r) {
    $tb.Cell($rowNo, 1).Shape.TextFrame.TextRange.Text = $r.dept
    $o = To-Paras @($r.opinion); Fill-Cell $tb.Cell($rowNo, 2) $o.paras $o.sub
    $a = To-Paras @($r.action);  Fill-Cell $tb.Cell($rowNo, 3) $a.paras $a.sub
    $tb.Cell($rowNo, 4).Shape.TextFrame.TextRange.Text = ""
}

$pos = $last + 1
$pageNo = 0
$cur = $null; $curTb = $null; $curShape = $null; $rowNo = 1
$n = 0
foreach ($r in $rows) {
    $n++
    $needNew = ($null -eq $cur)
    if (-not $needNew) {
        # 다음 행을 시험 삽입해 넘치면 새 페이지
        $curTb.Rows.Add() | Out-Null
        $rowNo++
        Put-Row $curTb $rowNo $r
        if (($curShape.Top + (Table-Height $curTb)) -gt $BOTTOM) {
            $curTb.Rows.Item($rowNo).Delete(); $rowNo--
            $needNew = $true
        }
    }
    if ($needNew) {
        $pageNo++
        $tmpl = if ($pageNo -eq 1) { $tmpl1 } else { $tmpl2 }
        $cur = New-Page $tmpl $pos; $pos++
        $curShape = Get-Table $cur; $curTb = $curShape.Table; $rowNo = 2
        Put-Row $curTb 2 $r
        $h = $curShape.Top + (Table-Height $curTb)
        if ($h -gt $BOTTOM) { Write-Warning "행이 한 페이지를 넘음: $($r.dept) (하단 $([int]$h)pt) — 문장 요약/글자 크기 조정 필요" }
        $pages += $cur
    }
}

# 첫 페이지 요약 문구
if ($pages.Count -gt 0) {
    $sum = @(("협의기간 : " + $Period), ("협의부서 : " + $ReplyCount + "개 부서"), ("제출의견 : " + $ReplyCount + "건"))
    $k = 0
    foreach ($sh in $pages[0].Shapes) {
        if ($sh.HasTextFrame -and $sh.TextFrame.HasText -and $sh.TextFrame.TextRange.Text -match "^(협의기간|협의부서|제출의견)") {
            $sh.TextFrame.TextRange.Text = $sum[$k]; $k++
        }
    }
}

# 원본 협의 슬라이드 삭제 (뒤에서부터)
for ($i = $origCount - 1; $i -ge 0; $i--) { $p.Slides.Item($idx[$i]).Delete() }

$p.Save()
"새 협의 슬라이드 $($pages.Count)장 생성 (행 $($rows.Count)개, 회신 ${ReplyCount}개 부서)"
$start = $idx[0]
for ($i = 0; $i -lt $pages.Count; $i++) { $p.Slides.Item($start + $i).Export((Join-Path (Split-Path $Deck) "consult_p$($i + 1).png"), "PNG", 1400, 1050) }
$p.Close()
