$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$assetDir = Join-Path $PSScriptRoot "assets_robot_vacuum"
$outPath = Join-Path $PSScriptRoot "로봇청소기_사용법_최종실습.pptx"

$purchaseUrl = "https://www.mi.com/kr/product/xiaomi-robot-vacuum-x20-plus/buy/"
$productUrl = "https://www.mi.com/kr/product/xiaomi-robot-vacuum-x20-plus/"
$priceUrl = "https://prod.danawa.com/info/?pcode=58836050"
$videoUrl = "https://www.youtube.com/watch?v=0wtwbwLYSvs"
$supportUrl = "https://www.mi.com/kr/support/faq/details/KA-125830/"

$imgHero = Join-Path $assetDir "4df7f0df3f8e694bcce24763a3c40fb2.png"
$imgApp = Join-Path $assetDir "0702fc0873589afa74f258d08d2f7e2f.png"
$imgMop = Join-Path $assetDir "e9427c1b774ba4f9ae0a56d1713bbdda.png"
$imgBrush = Join-Path $assetDir "x20plus_station.png"

$msoTrue = -1
$msoFalse = 0
$ppLayoutBlank = 12
$ppMouseClick = 1
$msoTextOrientationHorizontal = 1
$msoShapeRoundedRectangle = 5
$msoShapeRectangle = 1
$msoShapeOval = 9
$msoShapeLine = 9
$msoAlignCenter = 2
$msoAlignLeft = 1
$msoAnchorMiddle = 3
$msoAnchorTop = 1

function RGB($r, $g, $b) {
    return [int]($r + ($g * 256) + ($b * 65536))
}

function Add-Text($slide, $text, $x, $y, $w, $h, $size = 24, $color = $(RGB 28 34 42), $bold = $false, $align = $msoAlignLeft) {
    $shape = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $x, $y, $w, $h)
    $shape.TextFrame.TextRange.Text = $text
    $shape.TextFrame.MarginLeft = 8
    $shape.TextFrame.MarginRight = 8
    $shape.TextFrame.MarginTop = 4
    $shape.TextFrame.MarginBottom = 4
    $shape.TextFrame.WordWrap = $msoTrue
    $shape.TextFrame.TextRange.Font.Name = "맑은 고딕"
    $shape.TextFrame.TextRange.Font.Size = [single]$size
    $shape.TextFrame.TextRange.Font.Color.RGB = $color
    if ($bold) { $shape.TextFrame.TextRange.Font.Bold = $msoTrue }
    $shape.TextFrame.TextRange.ParagraphFormat.Alignment = $align
    return $shape
}

function Add-Title($slide, $title, $subtitle = "") {
    Add-Text $slide $title 48 34 820 54 28 $(RGB 18 24 31) $true | Out-Null
    if ($subtitle) {
        Add-Text $slide $subtitle 52 84 780 28 12 $(RGB 82 91 105) $false | Out-Null
    }
    $line = $slide.Shapes.AddShape($msoShapeRectangle, 52, 122, 84, 4)
    $line.Fill.ForeColor.RGB = RGB 30 132 121
    $line.Line.Visible = $msoFalse
}

function Add-Panel($slide, $x, $y, $w, $h, $fill = $(RGB 247 249 250), $line = $(RGB 224 229 232)) {
    $shape = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $x, $y, $w, $h)
    $shape.Fill.ForeColor.RGB = $fill
    $shape.Line.ForeColor.RGB = $line
    $shape.Adjustments.Item(1) = 0.08
    return $shape
}

function Add-Bullets($slide, $items, $x, $y, $w, $h, $size = 18) {
    $text = [string]::Join("`r", $items)
    $shape = Add-Text $slide $text $x $y $w $h $size $(RGB 37 45 55) $false
    $shape.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = $msoTrue
    $shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 7
    return $shape
}

function Add-LinkButton($slide, $label, $url, $x, $y, $w, $h, $fill = $(RGB 30 132 121)) {
    $btn = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $x, $y, $w, $h)
    $btn.Fill.ForeColor.RGB = $fill
    $btn.Line.Visible = $msoFalse
    $btn.Adjustments.Item(1) = 0.18
    $btn.TextFrame.TextRange.Text = $label
    $btn.TextFrame.TextRange.Font.Name = "맑은 고딕"
    $btn.TextFrame.TextRange.Font.Size = 15
    $btn.TextFrame.TextRange.Font.Bold = $msoTrue
    $btn.TextFrame.TextRange.Font.Color.RGB = RGB 255 255 255
    $btn.TextFrame.TextRange.ParagraphFormat.Alignment = $msoAlignCenter
    $btn.TextFrame.VerticalAnchor = $msoAnchorMiddle
    $btn.ActionSettings.Item($ppMouseClick).Hyperlink.Address = $url
    return $btn
}

function Add-PictureFit($slide, $path, $x, $y, $w, $h) {
    $pic = $slide.Shapes.AddPicture($path, $msoFalse, $msoTrue, $x, $y, -1, -1)
    $scale = [Math]::Min($w / $pic.Width, $h / $pic.Height)
    $pic.Width = $pic.Width * $scale
    $pic.Height = $pic.Height * $scale
    $pic.Left = $x + (($w - $pic.Width) / 2)
    $pic.Top = $y + (($h - $pic.Height) / 2)
    return $pic
}

function Add-Step($slide, $num, $title, $body, $x, $y, $w, $h) {
    Add-Panel $slide $x $y $w $h $(RGB 255 255 255) $(RGB 219 226 229) | Out-Null
    $circle = $slide.Shapes.AddShape($msoShapeOval, $x + 16, $y + 18, 34, 34)
    $circle.Fill.ForeColor.RGB = RGB 30 132 121
    $circle.Line.Visible = $msoFalse
    $circle.TextFrame.TextRange.Text = "$num"
    $circle.TextFrame.TextRange.Font.Name = "맑은 고딕"
    $circle.TextFrame.TextRange.Font.Size = 16
    $circle.TextFrame.TextRange.Font.Bold = $msoTrue
    $circle.TextFrame.TextRange.Font.Color.RGB = RGB 255 255 255
    $circle.TextFrame.TextRange.ParagraphFormat.Alignment = $msoAlignCenter
    $circle.TextFrame.VerticalAnchor = $msoAnchorMiddle
    Add-Text $slide $title ($x + 62) ($y + 16) ($w - 78) 26 17 $(RGB 18 24 31) $true | Out-Null
    Add-Text $slide $body ($x + 62) ($y + 47) ($w - 82) ($h - 56) 12.5 $(RGB 76 86 99) $false | Out-Null
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $msoTrue
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

try {
    # Slide 1
    $s = $pres.Slides.Add(1, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 243 246 247
    Add-PictureFit $s $imgHero 548 58 340 400 | Out-Null
    Add-Text $s "로봇 청소기`r사용법" 56 96 430 118 38 $(RGB 18 24 31) $true | Out-Null
    Add-Text $s "가성비 모델 예시: Xiaomi 로봇청소기 X20+`r구매 링크, 제품 사진, 시연 영상, 장단점, 사용 순서 포함" 60 230 440 78 17 $(RGB 70 80 92) | Out-Null
    Add-LinkButton $s "공식 구매 링크 열기" $purchaseUrl 60 338 190 44 | Out-Null
    Add-LinkButton $s "시연 영상 보기" $videoUrl 268 338 164 44 $(RGB 43 92 158) | Out-Null
    Add-Text $s "최종실습 PPT" 64 462 240 24 13 $(RGB 113 124 137) | Out-Null

    # Slide 2
    $s = $pres.Slides.Add(2, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 255 255
    Add-Title $s "제품 소개" "Xiaomi 로봇청소기 X20+"
    Add-PictureFit $s $imgApp 526 148 350 300 | Out-Null
    Add-Bullets $s @(
        "흡입 청소와 회전 물걸레 청소를 함께 지원",
        "올인원 베이스 스테이션으로 먼지 비움과 물걸레 세척을 자동화",
        "LDS 레이저 내비게이션으로 집 구조를 스캔하고 경로를 계획",
        "Xiaomi Home 앱으로 구역 청소, 예약 청소, 금지 구역 설정 가능"
    ) 76 166 390 250 18 | Out-Null
    Add-Text $s "핵심 컨셉: 바닥 정리만 잘해두면 매일 반복되는 청소를 자동화해 주는 생활가전" 78 430 760 38 17 $(RGB 30 132 121) $true | Out-Null

    # Slide 3
    $s = $pres.Slides.Add(3, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 245 247 247
    Add-Title $s "구매 정보" "가성비 모델 예시와 확인 링크"
    Add-Panel $s 66 154 390 270 $(RGB 255 255 255) $(RGB 222 229 231) | Out-Null
    Add-Text $s "추천 예시" 94 184 300 28 19 $(RGB 18 24 31) $true | Out-Null
    Add-Text $s "Xiaomi 로봇청소기 X20+`r공식/가격비교 기준 419,000원대 확인`r※ 가격은 판매처와 시점에 따라 변동" 94 225 320 118 22 $(RGB 37 45 55) $true | Out-Null
    Add-LinkButton $s "공식 구매 페이지" $purchaseUrl 94 356 160 42 | Out-Null
    Add-LinkButton $s "다나와 가격 확인" $priceUrl 270 356 160 42 $(RGB 43 92 158) | Out-Null
    Add-PictureFit $s $imgHero 560 120 280 360 | Out-Null

    # Slide 4
    $s = $pres.Slides.Add(4, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 255 255
    Add-Title $s "장점" "일상 청소를 줄여주는 기능"
    Add-PictureFit $s $imgMop 72 154 360 265 | Out-Null
    Add-Bullets $s @(
        "6,000Pa 흡입력으로 먼지와 머리카락 제거에 유리",
        "180rpm 회전 물걸레로 생활 얼룩 관리 가능",
        "자동 먼지 비움, 물걸레 세척, 공기 건조로 손이 덜 감",
        "앱에서 방별 청소, 예약, 금지 구역을 세밀하게 설정",
        "카펫 구역에서는 물걸레를 들어 올리는 기능 지원"
    ) 500 154 370 260 17 | Out-Null

    # Slide 5
    $s = $pres.Slides.Add(5, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 245 247 247
    Add-Title $s "단점과 주의점" "사기 전, 쓰기 전 확인할 부분"
    Add-Panel $s 70 154 820 290 $(RGB 255 255 255) $(RGB 222 229 231) | Out-Null
    Add-Bullets $s @(
        "바닥에 전선, 양말, 얇은 매트가 많으면 청소가 중단될 수 있음",
        "물걸레는 찌든 때 제거보다 매일 관리용에 가까움",
        "스테이션 설치 공간과 콘센트 위치를 미리 확보해야 함",
        "먼지봉투, 필터, 브러시, 물걸레 패드 같은 소모품 교체 비용 발생",
        "처음 맵을 만들 때는 문을 열고 장애물을 치워야 정확도가 좋아짐"
    ) 110 192 740 220 19 | Out-Null

    # Slide 6
    $s = $pres.Slides.Add(6, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 255 255
    Add-Title $s "사용법" "처음 설치부터 청소 시작까지"
    Add-Step $s 1 "설치" "스테이션을 벽 가까이 두고 좌우 공간과 콘센트를 확보한다." 54 150 270 115
    Add-Step $s 2 "앱 연결" "Xiaomi Home 앱에서 기기를 추가하고 Wi-Fi에 연결한다." 345 150 270 115
    Add-Step $s 3 "지도 생성" "첫 청소 전 바닥 물건을 치우고 전체 맵을 만든다." 636 150 270 115
    Add-Step $s 4 "청소 설정" "방 이름, 금지 구역, 카펫 구역, 흡입력과 물량을 설정한다." 54 302 270 115
    Add-Step $s 5 "청소 실행" "전체 청소, 구역 청소, 예약 청소 중 필요한 방식을 선택한다." 345 302 270 115
    Add-Step $s 6 "관리" "먼지봉투, 물통, 필터, 브러시, 센서를 주기적으로 확인한다." 636 302 270 115

    # Slide 7
    $s = $pres.Slides.Add(7, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 245 247 247
    Add-Title $s "시연 동영상" "유튜브 링크로 사용 장면 확인"
    Add-Panel $s 70 156 420 245 $(RGB 255 255 255) $(RGB 222 229 231) | Out-Null
    Add-Text $s "시연 영상에서 확인할 것" 102 186 340 32 21 $(RGB 18 24 31) $true | Out-Null
    Add-Bullets $s @(
        "실제 바닥에서 이동하는 방식",
        "앱 지도와 구역 청소 화면",
        "물걸레 청소와 스테이션 복귀 과정",
        "청소 전 바닥 정리가 왜 필요한지"
    ) 102 232 340 126 16 | Out-Null
    Add-LinkButton $s "YouTube 시연 영상 열기" $videoUrl 102 354 214 42 $(RGB 205 53 59) | Out-Null
    Add-PictureFit $s $imgBrush 540 150 340 260 | Out-Null
    Add-Text $s $videoUrl 84 438 800 26 12 $(RGB 75 86 100) | Out-Null

    # Slide 8
    $s = $pres.Slides.Add(8, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 255 255
    Add-Title $s "출처와 마무리" "실습 파일 안의 링크는 클릭해서 열 수 있습니다"
    Add-Text $s "핵심 정리" 78 156 240 28 22 $(RGB 18 24 31) $true | Out-Null
    Add-Bullets $s @(
        "로봇 청소기는 바닥 정리와 앱 설정을 잘할수록 만족도가 올라감",
        "흡입, 물걸레, 자동 관리 기능을 생활 패턴에 맞춰 예약하면 효과적",
        "구매 전에는 가격, AS, 소모품, 설치 공간을 함께 확인"
    ) 78 200 450 140 17 | Out-Null
    Add-Text $s "자료 링크" 580 156 240 28 22 $(RGB 18 24 31) $true | Out-Null
    Add-LinkButton $s "제품 공식 소개" $productUrl 580 206 170 40 | Out-Null
    Add-LinkButton $s "공식 구매 링크" $purchaseUrl 580 258 170 40 | Out-Null
    Add-LinkButton $s "FAQ/지원 문서" $supportUrl 580 310 170 40 $(RGB 43 92 158) | Out-Null
    Add-LinkButton $s "가격 비교" $priceUrl 580 362 170 40 $(RGB 88 96 105) | Out-Null
    Add-Text $s "사진: Xiaomi 공식 제품 페이지 / 영상: YouTube" 78 452 720 28 13 $(RGB 113 124 137) | Out-Null

    $pres.SaveAs($outPath)
}
finally {
    $pres.Close()
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}

Write-Output $outPath
