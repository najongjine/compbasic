$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutRoot = Join-Path $Root "examples"
$AssetRoot = Join-Path $Root "assets"
New-Item -ItemType Directory -Force -Path $OutRoot, $AssetRoot | Out-Null

function New-SampleImage([string]$Path, [string]$Title, [string]$SubTitle, [string]$Color1, [string]$Color2) {
    Add-Type -AssemblyName System.Drawing
    $bmp = New-Object Drawing.Bitmap 900, 520
    $g = [Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $rect = New-Object Drawing.Rectangle 0, 0, 900, 520
    $brush = New-Object Drawing.Drawing2D.LinearGradientBrush $rect, ([Drawing.ColorTranslator]::FromHtml($Color1)), ([Drawing.ColorTranslator]::FromHtml($Color2)), 25
    $g.FillRectangle($brush, $rect)
    $pen = New-Object Drawing.Pen ([Drawing.Color]::FromArgb(230, 255, 255, 255)), 8
    $g.DrawRectangle($pen, 32, 32, 836, 456)
    $fontTitle = New-Object Drawing.Font "Malgun Gothic", 46, ([Drawing.FontStyle]::Bold)
    $fontSub = New-Object Drawing.Font "Malgun Gothic", 24, ([Drawing.FontStyle]::Regular)
    $white = New-Object Drawing.SolidBrush ([Drawing.Color]::White)
    $shadow = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb(80, 0, 0, 0))
    $g.DrawString($Title, $fontTitle, $shadow, 75, 151)
    $g.DrawString($Title, $fontTitle, $white, 70, 146)
    $g.DrawString($SubTitle, $fontSub, $shadow, 75, 252)
    $g.DrawString($SubTitle, $fontSub, $white, 70, 247)
    $bmp.Save($Path, [Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose()
    $bmp.Dispose()
}

function New-Hwp {
    if ($null -eq $script:SharedHwp) {
        $script:SharedHwp = New-Object -ComObject HWPFrame.HwpObject
        try { $script:SharedHwp.XHwpWindows.Item(0).Visible = $false } catch {}
        try { $script:SharedHwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule") | Out-Null } catch {}
    }
    $hwp = $script:SharedHwp
    try { $hwp.Run("FileNew") | Out-Null } catch {}
    return $hwp
}

function Try-Run($hwp, [string]$command) {
    try { $hwp.Run($command) | Out-Null } catch {}
}

function Put-Text($hwp, [string]$text) {
    $hwp.HAction.GetDefault("InsertText", $hwp.HParameterSet.HInsertText.HSet) | Out-Null
    $hwp.HParameterSet.HInsertText.Text = $text
    $hwp.HAction.Execute("InsertText", $hwp.HParameterSet.HInsertText.HSet) | Out-Null
}

function New-Line($hwp, [int]$count = 1) {
    for ($i = 0; $i -lt $count; $i++) {
        $hwp.Run("BreakPara") | Out-Null
    }
}

function Put-Line($hwp, [string]$text = "") {
    if ($text.Length -gt 0) { Put-Text $hwp $text }
    New-Line $hwp
}

function Put-Title($hwp, [string]$text) {
    Try-Run $hwp "ParagraphShapeAlignCenter"
    Try-Run $hwp "CharShapeBold"
    Put-Line $hwp $text
    Try-Run $hwp "CharShapeBold"
    Try-Run $hwp "ParagraphShapeAlignLeft"
    New-Line $hwp
}

function Put-Table($hwp, [string[][]]$rows) {
    $rowCount = $rows.Count
    $colCount = $rows[0].Count
    $act = $hwp.CreateAction("TableCreate")
    $set = $act.CreateSet()
    $act.GetDefault($set) | Out-Null
    $set.SetItem("Rows", $rowCount) | Out-Null
    $set.SetItem("Cols", $colCount) | Out-Null
    $set.SetItem("WidthType", 2) | Out-Null
    $set.SetItem("HeightType", 1) | Out-Null
    $act.Execute($set) | Out-Null

    for ($r = 0; $r -lt $rowCount; $r++) {
        for ($c = 0; $c -lt $colCount; $c++) {
            if ($r -eq 0) { Try-Run $hwp "CharShapeBold" }
            Put-Text $hwp $rows[$r][$c]
            if ($r -eq 0) { Try-Run $hwp "CharShapeBold" }
            if (!($r -eq $rowCount - 1 -and $c -eq $colCount - 1)) {
                $hwp.Run("TableRightCell") | Out-Null
            }
        }
    }
    $hwp.Run("MoveDocEnd") | Out-Null
    New-Line $hwp 2
}

function Insert-Picture($hwp, [string]$path, [int]$widthMm = 90, [int]$heightMm = 52) {
    $w = $hwp.MiliToHwpUnit($widthMm)
    $h = $hwp.MiliToHwpUnit($heightMm)
    $hwp.InsertPicture($path, $false, 1, $false, $false, 0, $w, $h) | Out-Null
}

function Save-Hwpx($hwp, [string]$fileName) {
    $out = Join-Path $OutRoot $fileName
    if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
    $hwp.SaveAs($out, "HWPX", "") | Out-Null
}

function Build-ImageInvitation {
    $img = Join-Path $AssetRoot "sample_concert.png"
    New-SampleImage $img "우리 동네 작은 음악회" "이미지 배치 연습" "#2f6f73" "#d9a441"

    $hwp = New-Hwp
    Put-Title $hwp "우리 동네 작은 음악회"
    Put-Line $hwp "초대합니다."
    Put-Line $hwp "일시: 2026년 7월 3일 오후 3시"
    Put-Line $hwp "장소: 주민센터 강당"
    New-Line $hwp
    Put-Line $hwp "1. 글자처럼 취급 예시"
    Insert-Picture $hwp (Resolve-Path $img).Path 85 49
    New-Line $hwp 2
    Put-Line $hwp "2. 어울림 배치 연습: 아래 그림을 선택하고 개체 속성에서 어울림으로 바꿔 봅니다."
    Insert-Picture $hwp (Resolve-Path $img).Path 65 38
    Put-Line $hwp "그림 주변에 글자가 어떻게 놓이는지 확인합니다."
    New-Line $hwp
    Put-Line $hwp "3. 글 뒤로 배치 연습: 아래 그림을 선택하고 글 뒤로를 적용해 봅니다."
    Insert-Picture $hwp (Resolve-Path $img).Path 95 55
    Put-Line $hwp "배경 그림 위에 제목과 날짜를 올리는 연습을 합니다."
    Save-Hwpx $hwp "06_이미지배치_초대장.hwpx"
}

function Build-ContactTable {
    $hwp = New-Hwp
    Put-Title $hwp "우리 반 연락망"
    Put-Line $hwp "아래는 실제 한글 표로 만든 연락망입니다."
    New-Line $hwp
    Put-Table $hwp @(
        @("번호", "이름", "전화번호", "비고"),
        @("1", "홍길동", "010-1111-2222", "반장"),
        @("2", "김영희", "010-3333-4444", "연락 가능"),
        @("3", "이민수", "010-5555-6666", "문자 선호"),
        @("4", "박정자", "010-7777-8888", "")
    )
    Put-Line $hwp "따라 하기"
    Put-Line $hwp "1. Ctrl+N,T로 5줄 4칸 표를 만듭니다."
    Put-Line $hwp "2. 첫 줄에는 번호, 이름, 전화번호, 비고를 입력합니다."
    Put-Line $hwp "3. 첫 줄을 굵게 만들고 L에서 테두리와 배경을 바꿉니다."
    Put-Line $hwp "4. 셀 너비와 높이를 같게 맞춥니다."
    Save-Hwpx $hwp "07_표기본_연락망.hwpx"
}

function Build-ScheduleTable {
    $img = Join-Path $AssetRoot "sample_class.png"
    New-SampleImage $img "6월 교육" "표 안 그림 넣기" "#455a8a" "#4c9f70"

    $hwp = New-Hwp
    Put-Title $hwp "6월 교육 일정표"
    Put-Line $hwp "아래는 실제 한글 표와 실제 삽입 그림으로 만든 일정표 예시입니다."
    Put-Line $hwp "수강생은 F7을 눌러 용지 방향을 가로로 바꾼 뒤 따라 만듭니다."
    New-Line $hwp
    Put-Table $hwp @(
        @("교육명", "월", "화", "수", "목"),
        @("컴퓨터 기초", "마우스", "키보드", "한글", "복습"),
        @("한글 문서", "저장", "글자 모양", "표 만들기", "그림 넣기"),
        @("실습", "찾기", "바꾸기", "인쇄", "발표"),
        @("메모", "천천히", "반복", "같이 만들기", "완성")
    )
    Put-Line $hwp "표 안 그림 넣기 예시"
    Insert-Picture $hwp (Resolve-Path $img).Path 70 40
    New-Line $hwp
    Put-Line $hwp "따라 하기"
    Put-Line $hwp "1. 첫 줄 셀을 선택해 셀 합치기를 합니다."
    Put-Line $hwp "2. 요일 칸을 선택하고 L에서 배경색을 넣습니다."
    Put-Line $hwp "3. 표 안 칸에 커서를 두고 입력 > 그림으로 이미지를 넣습니다."
    Save-Hwpx $hwp "08_표응용_일정표.hwpx"
}

$script:SharedHwp = $null
Build-ImageInvitation
Build-ContactTable
Build-ScheduleTable
if ($null -ne $script:SharedHwp) {
    $script:SharedHwp.Quit() | Out-Null
}

Write-Host "Rebuilt 06-08 with Hanword COM: real pictures and real tables."

