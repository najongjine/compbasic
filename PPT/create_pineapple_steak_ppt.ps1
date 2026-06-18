$ErrorActionPreference = "Stop"

$assetDir = Join-Path $PSScriptRoot "assets_pineapple_steak"
$outPath = Join-Path $PSScriptRoot "파인애플_스테이크_요리만들기.pptx"

$videoUrl = "https://www.youtube.com/watch?v=mdxLYB1v2j0"
$thumb = Join-Path $assetDir "maxresdefault.jpg"

$msoTrue = -1
$msoFalse = 0
$ppLayoutBlank = 12
$ppMouseClick = 1
$msoTextOrientationHorizontal = 1
$msoShapeRoundedRectangle = 5
$msoShapeRectangle = 1
$msoShapeOval = 9
$msoAlignCenter = 2
$msoAlignLeft = 1
$msoAnchorMiddle = 3

function RGB($r, $g, $b) {
    return [int]($r + ($g * 256) + ($b * 65536))
}

function Add-Text($slide, $text, $x, $y, $w, $h, $size = 22, $color = $(RGB 30 33 36), $bold = $false, $align = $msoAlignLeft) {
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
    Add-Text $slide $title 48 32 820 54 29 $(RGB 25 28 31) $true | Out-Null
    if ($subtitle) {
        Add-Text $slide $subtitle 52 84 780 30 13 $(RGB 92 82 73) $false | Out-Null
    }
    $line = $slide.Shapes.AddShape($msoShapeRectangle, 52, 122, 88, 4)
    $line.Fill.ForeColor.RGB = RGB 186 82 38
    $line.Line.Visible = $msoFalse
}

function Add-Panel($slide, $x, $y, $w, $h, $fill = $(RGB 255 248 240), $line = $(RGB 232 215 197)) {
    $shape = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $x, $y, $w, $h)
    $shape.Fill.ForeColor.RGB = $fill
    $shape.Line.ForeColor.RGB = $line
    $shape.Adjustments.Item(1) = 0.08
    return $shape
}

function Add-Bullets($slide, $items, $x, $y, $w, $h, $size = 18) {
    $shape = Add-Text $slide ([string]::Join("`r", $items)) $x $y $w $h $size $(RGB 47 43 39)
    $shape.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = $msoTrue
    $shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 7
    return $shape
}

function Add-LinkButton($slide, $label, $url, $x, $y, $w, $h, $fill = $(RGB 186 82 38)) {
    $btn = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $x, $y, $w, $h)
    $btn.Fill.ForeColor.RGB = $fill
    $btn.Line.Visible = $msoFalse
    $btn.Adjustments.Item(1) = 0.16
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

function Add-PictureFill($slide, $path, $x, $y, $w, $h) {
    $pic = $slide.Shapes.AddPicture($path, $msoFalse, $msoTrue, $x, $y, -1, -1)
    $scale = [Math]::Max($w / $pic.Width, $h / $pic.Height)
    $pic.Width = $pic.Width * $scale
    $pic.Height = $pic.Height * $scale
    $pic.Left = $x + (($w - $pic.Width) / 2)
    $pic.Top = $y + (($h - $pic.Height) / 2)
    return $pic
}

function Add-StepCard($slide, $num, $title, $body, $x, $y, $w, $h) {
    Add-Panel $slide $x $y $w $h $(RGB 255 255 255) $(RGB 229 214 198) | Out-Null
    $circle = $slide.Shapes.AddShape($msoShapeOval, $x + 16, $y + 18, 34, 34)
    $circle.Fill.ForeColor.RGB = RGB 186 82 38
    $circle.Line.Visible = $msoFalse
    $circle.TextFrame.TextRange.Text = "$num"
    $circle.TextFrame.TextRange.Font.Name = "맑은 고딕"
    $circle.TextFrame.TextRange.Font.Size = 16
    $circle.TextFrame.TextRange.Font.Bold = $msoTrue
    $circle.TextFrame.TextRange.Font.Color.RGB = RGB 255 255 255
    $circle.TextFrame.TextRange.ParagraphFormat.Alignment = $msoAlignCenter
    $circle.TextFrame.VerticalAnchor = $msoAnchorMiddle
    Add-Text $slide $title ($x + 62) ($y + 16) ($w - 78) 26 17 $(RGB 29 29 28) $true | Out-Null
    Add-Text $slide $body ($x + 62) ($y + 48) ($w - 80) ($h - 56) 12.5 $(RGB 86 76 68) | Out-Null
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $msoTrue
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

try {
    # 1
    $s = $pres.Slides.Add(1, $ppLayoutBlank)
    Add-PictureFill $s $thumb 0 0 960 540 | Out-Null
    $overlay = $s.Shapes.AddShape($msoShapeRectangle, 0, 0, 430, 540)
    $overlay.Fill.ForeColor.RGB = RGB 24 20 17
    $overlay.Fill.Transparency = 0.08
    $overlay.Line.Visible = $msoFalse
    Add-Text $s "파인애플`r스테이크 만들기" 52 114 330 118 38 $(RGB 255 255 255) $true | Out-Null
    Add-Text $s "영상 기반 요리 실습 PPT`rOutdoor Cooking The Juiciest Steaks with Pineapples!" 58 258 332 70 17 $(RGB 255 232 205) | Out-Null
    Add-LinkButton $s "원본 영상 보기" $videoUrl 60 370 168 44 | Out-Null

    # 2
    $s = $pres.Slides.Add(2, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 251 246
    Add-Title $s "요리 소개" "야외에서 두툼한 스테이크와 파인애플을 함께 굽는 메뉴"
    Add-PictureFill $s $thumb 538 144 330 230 | Out-Null
    Add-Bullets $s @(
        "고기의 묵직한 풍미와 파인애플의 단맛, 산미를 함께 살리는 요리",
        "파인애플은 고기 잡내를 줄이고 달콤한 소스 느낌을 더함",
        "영상은 야외 조리 분위기가 강하므로 발표용으로 시각적 효과가 좋음",
        "실습에서는 팬, 그릴팬, 에어프라이어 등으로 응용 가능"
    ) 76 166 390 220 18 | Out-Null
    Add-Text $s "주의: 이 PPT는 영상 제목과 장면을 바탕으로 수업용 조리 흐름을 재구성한 자료입니다." 78 428 780 30 14 $(RGB 137 91 56) $true | Out-Null

    # 3
    $s = $pres.Slides.Add(3, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 255 255
    Add-Title $s "준비 재료" "2~3인분 기준 예시"
    Add-Panel $s 70 156 370 270 $(RGB 255 248 240) $(RGB 232 215 197) | Out-Null
    Add-Text $s "주재료" 104 188 260 30 22 $(RGB 29 29 28) $true | Out-Null
    Add-Bullets $s @(
        "두툼한 소고기 스테이크용 2~3장",
        "파인애플 1개 또는 슬라이스 파인애플",
        "소금, 후추",
        "올리브오일 또는 식용유",
        "버터, 마늘, 허브 선택"
    ) 104 232 290 140 17 | Out-Null
    Add-Panel $s 520 156 370 270 $(RGB 255 248 240) $(RGB 232 215 197) | Out-Null
    Add-Text $s "도구" 554 188 260 30 22 $(RGB 29 29 28) $true | Out-Null
    Add-Bullets $s @(
        "그릴 또는 두꺼운 팬",
        "집게, 칼, 도마",
        "키친타월",
        "접시 또는 나무 보드",
        "고기 온도계가 있으면 더 좋음"
    ) 554 232 290 140 17 | Out-Null

    # 4
    $s = $pres.Slides.Add(4, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 251 246
    Add-Title $s "만드는 순서 1" "굽기 전 준비"
    Add-StepCard $s 1 "고기 물기 제거" "키친타월로 표면 수분을 닦아야 겉면이 잘 구워진다." 54 154 270 120
    Add-StepCard $s 2 "밑간하기" "소금과 후추를 앞뒤로 뿌리고 오일을 얇게 발라둔다." 345 154 270 120
    Add-StepCard $s 3 "파인애플 손질" "껍질과 심을 제거하고 두툼한 링 또는 큼직한 조각으로 자른다." 636 154 270 120
    Add-StepCard $s 4 "팬 예열" "팬이나 그릴을 충분히 달군 뒤 고기를 올려야 육즙이 덜 빠진다." 54 318 270 120
    Add-StepCard $s 5 "곁들임 준비" "마늘, 허브, 버터를 준비하면 향을 더하기 쉽다." 345 318 270 120
    Add-StepCard $s 6 "안전 확인" "야외 조리 시 화기 주변 정리와 고기용/과일용 도마 구분을 확인한다." 636 318 270 120

    # 5
    $s = $pres.Slides.Add(5, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 255 255
    Add-Title $s "만드는 순서 2" "스테이크와 파인애플 굽기"
    Add-PictureFill $s $thumb 70 158 360 245 | Out-Null
    Add-Bullets $s @(
        "고기를 센 불에서 앞뒤로 구워 겉면을 먼저 익힌다",
        "두꺼운 고기는 중불로 낮춰 속까지 천천히 익힌다",
        "파인애플은 고기 옆에서 살짝 눌어붙게 구워 단맛을 끌어낸다",
        "버터, 마늘, 허브를 넣어 녹인 뒤 고기 위에 끼얹으면 풍미가 좋아짐",
        "구운 고기는 바로 썰지 말고 5분 정도 쉬게 한다"
    ) 506 154 370 250 17 | Out-Null

    # 6
    $s = $pres.Slides.Add(6, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 251 246
    Add-Title $s "맛있게 만드는 포인트" "육즙, 단맛, 향을 살리는 핵심"
    Add-Panel $s 72 160 240 220 $(RGB 255 255 255) $(RGB 229 214 198) | Out-Null
    Add-Text $s "육즙" 104 196 160 28 22 $(RGB 186 82 38) $true $msoAlignCenter | Out-Null
    Add-Text $s "고기는 굽기 전 물기를 닦고, 구운 뒤에는 잠깐 레스팅한다." 104 244 160 74 16 $(RGB 47 43 39) $false $msoAlignCenter | Out-Null
    Add-Panel $s 360 160 240 220 $(RGB 255 255 255) $(RGB 229 214 198) | Out-Null
    Add-Text $s "파인애플" 392 196 160 28 22 $(RGB 186 82 38) $true $msoAlignCenter | Out-Null
    Add-Text $s "과육을 너무 얇게 자르면 쉽게 무르므로 두툼하게 굽는다." 392 244 160 74 16 $(RGB 47 43 39) $false $msoAlignCenter | Out-Null
    Add-Panel $s 648 160 240 220 $(RGB 255 255 255) $(RGB 229 214 198) | Out-Null
    Add-Text $s "불 조절" 680 196 160 28 22 $(RGB 186 82 38) $true $msoAlignCenter | Out-Null
    Add-Text $s "겉은 센 불, 속은 중불. 탄 냄새가 나면 바로 불을 낮춘다." 680 244 160 74 16 $(RGB 47 43 39) $false $msoAlignCenter | Out-Null
    Add-Text $s "덜 익은 고기가 걱정되면 중심 온도를 확인하거나 한 조각을 잘라 익힘 정도를 확인합니다." 104 430 760 30 17 $(RGB 137 91 56) $true | Out-Null

    # 7
    $s = $pres.Slides.Add(7, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 255 255
    Add-Title $s "플레이팅과 발표 팁" "완성 장면을 맛있게 보여주기"
    Add-Bullets $s @(
        "스테이크는 결 반대 방향으로 썰어 부드럽게 먹기",
        "구운 파인애플을 고기 옆이나 위에 올려 색감을 살리기",
        "육즙과 버터 소스를 살짝 뿌려 윤기 더하기",
        "발표할 때는 '파인애플의 단맛과 산미가 고기와 어울린다'고 설명",
        "야외 조리 영상의 분위기를 살려 자연, 불, 고기 굽는 소리를 강조"
    ) 78 160 460 240 18 | Out-Null
    Add-PictureFill $s $thumb 608 148 250 250 | Out-Null

    # 8
    $s = $pres.Slides.Add(8, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 251 246
    Add-Title $s "영상 삽입 / 출처" "PowerPoint에서 온라인 비디오로 넣기"
    Add-Panel $s 78 154 520 292 $(RGB 35 31 28) $(RGB 35 31 28) | Out-Null
    Add-PictureFill $s $thumb 78 154 520 292 | Out-Null
    $dark = $s.Shapes.AddShape($msoShapeRectangle, 78, 154, 520, 292)
    $dark.Fill.ForeColor.RGB = RGB 0 0 0
    $dark.Fill.Transparency = 0.46
    $dark.Line.Visible = $msoFalse
    $play = $s.Shapes.AddShape($msoShapeOval, 288, 246, 96, 96)
    $play.Fill.ForeColor.RGB = RGB 205 65 55
    $play.Line.Visible = $msoFalse
    Add-Text $s "▶" 317 265 44 44 32 $(RGB 255 255 255) $true $msoAlignCenter | Out-Null
    Add-LinkButton $s "YouTube 영상 열기" $videoUrl 662 184 180 44 $(RGB 205 65 55) | Out-Null
    Add-Text $s "삽입 방법:`r삽입 > 비디오 > 온라인 비디오`r아래 주소 붙여넣기" 638 254 240 92 16 $(RGB 47 43 39) | Out-Null
    Add-Text $s $videoUrl 638 362 260 40 12 $(RGB 92 82 73) | Out-Null
    Add-Text $s "출처: YouTube 영상 및 썸네일`r영상 제목: Outdoor Cooking The Juiciest Steaks with Pineapples!" 82 468 760 38 12.5 $(RGB 102 91 81) | Out-Null

    $pres.SaveAs($outPath)
}
finally {
    $pres.Close()
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}

Write-Output $outPath
