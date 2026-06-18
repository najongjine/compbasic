$ErrorActionPreference = "Stop"

$assetDir = Join-Path $PSScriptRoot "assets_dambi"
$outPath = Join-Path $PSScriptRoot "담비_동물_백과사전.pptx"

$videoUrl = "https://www.youtube.com/results?search_query=%EB%8B%B4%EB%B9%84+%EB%8F%99%EB%AC%BC+%EC%98%81%EC%83%81"
$videoDirectUrl = "https://www.youtube.com/watch?v=LNW0MCpg3oU"
$speciesUrl = "https://terms.naver.com/entry.naver?docId=1080300&cid=40942&categoryId=32624"
$encyclopediaUrl = "https://encykorea.aks.ac.kr/Article/E0013917"
$ministryUrl = "https://www.mcee.go.kr/home/web/board/read.do?boardId=1867680&boardMasterId=939&menuId=10598"

$imgForest = Join-Path $assetDir "dambi_forest_stock.png"
$imgPortrait = Join-Path $assetDir "dambi_portrait_stock.png"

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

function RGB($r, $g, $b) {
    return [int]($r + ($g * 256) + ($b * 65536))
}

function Add-Text($slide, $text, $x, $y, $w, $h, $size = 22, $color = $(RGB 28 34 42), $bold = $false, $align = $msoAlignLeft) {
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
    Add-Text $slide $title 48 32 820 54 29 $(RGB 18 27 28) $true | Out-Null
    if ($subtitle) {
        Add-Text $slide $subtitle 52 84 780 28 12.5 $(RGB 84 95 96) $false | Out-Null
    }
    $line = $slide.Shapes.AddShape($msoShapeRectangle, 52, 122, 88, 4)
    $line.Fill.ForeColor.RGB = RGB 87 126 83
    $line.Line.Visible = $msoFalse
}

function Add-Panel($slide, $x, $y, $w, $h, $fill = $(RGB 247 249 246), $line = $(RGB 220 228 218)) {
    $shape = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $x, $y, $w, $h)
    $shape.Fill.ForeColor.RGB = $fill
    $shape.Line.ForeColor.RGB = $line
    $shape.Adjustments.Item(1) = 0.08
    return $shape
}

function Add-Bullets($slide, $items, $x, $y, $w, $h, $size = 18) {
    $shape = Add-Text $slide ([string]::Join("`r", $items)) $x $y $w $h $size $(RGB 38 48 46)
    $shape.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = $msoTrue
    $shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 7
    return $shape
}

function Add-LinkButton($slide, $label, $url, $x, $y, $w, $h, $fill = $(RGB 87 126 83)) {
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

function Add-PictureFit($slide, $path, $x, $y, $w, $h) {
    $pic = $slide.Shapes.AddPicture($path, $msoFalse, $msoTrue, $x, $y, -1, -1)
    $scale = [Math]::Max($w / $pic.Width, $h / $pic.Height)
    $pic.Width = $pic.Width * $scale
    $pic.Height = $pic.Height * $scale
    $pic.Left = $x + (($w - $pic.Width) / 2)
    $pic.Top = $y + (($h - $pic.Height) / 2)
    return $pic
}

function Add-Fact($slide, $label, $value, $x, $y, $w, $h) {
    Add-Panel $slide $x $y $w $h $(RGB 255 255 255) $(RGB 218 226 215) | Out-Null
    Add-Text $slide $label ($x + 16) ($y + 12) ($w - 32) 22 12 $(RGB 87 126 83) $true | Out-Null
    Add-Text $slide $value ($x + 16) ($y + 39) ($w - 32) ($h - 44) 17 $(RGB 31 39 38) $true | Out-Null
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $msoTrue
$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

try {
    # 1
    $s = $pres.Slides.Add(1, $ppLayoutBlank)
    Add-PictureFit $s $imgForest 0 0 960 540 | Out-Null
    $overlay = $s.Shapes.AddShape($msoShapeRectangle, 0, 0, 460, 540)
    $overlay.Fill.ForeColor.RGB = RGB 20 28 25
    $overlay.Fill.Transparency = 0.10
    $overlay.Line.Visible = $msoFalse
    Add-Text $s "담비`r동물 백과사전" 52 116 350 116 40 $(RGB 255 255 255) $true | Out-Null
    Add-Text $s "Yellow-throated marten`rMartes flavigula" 58 252 330 64 19 $(RGB 225 235 222) | Out-Null
    Add-Text $s "PPT 스톡 이미지 느낌의 야생동물 사진 사용" 60 438 340 28 13 $(RGB 221 230 217) | Out-Null

    # 2
    $s = $pres.Slides.Add(2, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 250 251 248
    Add-Title $s "한눈에 보는 담비" "식육목 족제비과에 속하는 야생 포유류"
    Add-PictureFit $s $imgPortrait 612 136 250 250 | Out-Null
    Add-Fact $s "분류" "포유류 / 식육목 / 족제비과" 70 158 235 90
    Add-Fact $s "학명" "Martes flavigula" 326 158 235 90
    Add-Fact $s "몸길이" "약 55~65cm, 꼬리 약 40~45cm" 70 278 235 100
    Add-Fact $s "별명" "목과 가슴이 노란 담비" 326 278 235 100
    Add-Text $s "검은 얼굴과 다리, 노란 목과 가슴, 긴 꼬리가 눈에 띄는 산림성 동물입니다." 76 426 760 34 18 $(RGB 87 126 83) $true | Out-Null

    # 3
    $s = $pres.Slides.Add(3, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 255 255
    Add-Title $s "생김새와 특징" "빠르고 민첩한 숲의 사냥꾼"
    Add-PictureFit $s $imgForest 72 156 372 250 | Out-Null
    Add-Bullets $s @(
        "몸은 길고 날씬하며 다리는 비교적 짧음",
        "나무를 잘 타고 바위와 숲길을 빠르게 이동",
        "노란 목과 가슴, 어두운 얼굴과 꼬리가 대표 특징",
        "시각, 후각, 청각이 발달해 먹이를 찾는 데 유리",
        "단독 또는 작은 무리로 움직이는 모습이 관찰됨"
    ) 510 154 360 250 18 | Out-Null

    # 4
    $s = $pres.Slides.Add(4, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 250 251 248
    Add-Title $s "서식지와 먹이" "산림 생태계에서 중요한 포식자"
    Add-Panel $s 70 152 390 280 $(RGB 255 255 255) $(RGB 218 226 215) | Out-Null
    Add-Text $s "사는 곳" 102 184 300 30 22 $(RGB 18 27 28) $true | Out-Null
    Add-Bullets $s @(
        "한국의 산림, 계곡 주변, 숲이 이어진 지역",
        "동아시아와 동남아시아의 산림 지대",
        "나무와 바위, 덤불이 많은 환경을 선호"
    ) 102 228 300 126 17 | Out-Null
    Add-Panel $s 500 152 390 280 $(RGB 255 255 255) $(RGB 218 226 215) | Out-Null
    Add-Text $s "먹이" 532 184 300 30 22 $(RGB 18 27 28) $true | Out-Null
    Add-Bullets $s @(
        "쥐, 다람쥐, 새, 알, 곤충",
        "개구리와 작은 파충류",
        "계절에 따라 열매와 식물성 먹이도 섭취",
        "멧돼지 폐사체를 먹는 청소동물 역할도 함"
    ) 532 228 300 146 17 | Out-Null

    # 5
    $s = $pres.Slides.Add(5, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 255 255
    Add-Title $s "생태계에서 하는 일" "숲의 균형을 지키는 동물"
    Add-PictureFit $s $imgPortrait 628 146 230 230 | Out-Null
    Add-Bullets $s @(
        "작은 포유류와 조류의 개체 수 조절에 도움",
        "죽은 동물 사체를 먹어 산림 위생 유지에 기여",
        "열매를 먹고 이동하면서 씨앗 확산에 일부 기여",
        "담비가 사는 숲은 먹이망과 서식 환경이 비교적 건강하다는 신호가 될 수 있음"
    ) 82 162 470 220 19 | Out-Null
    Add-Text $s "보전 포인트: 숲이 끊기면 이동로가 줄고, 로드킬과 서식지 단절 위험이 커집니다." 86 426 760 34 18 $(RGB 148 84 44) $true | Out-Null

    # 6
    $s = $pres.Slides.Add(6, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 250 251 248
    Add-Title $s "관찰 영상" "삽입 > 비디오 > 온라인 비디오에 넣을 유튜브 링크"
    Add-Panel $s 82 154 520 292 $(RGB 34 40 38) $(RGB 34 40 38) | Out-Null
    $play = $s.Shapes.AddShape($msoShapeOval, 292, 244, 100, 100)
    $play.Fill.ForeColor.RGB = RGB 205 65 55
    $play.Line.Visible = $msoFalse
    Add-Text $s "▶" 321 264 44 44 34 $(RGB 255 255 255) $true $msoAlignCenter | Out-Null
    Add-Text $s "YouTube 담비 영상" 188 360 310 30 21 $(RGB 255 255 255) $true $msoAlignCenter | Out-Null
    Add-LinkButton $s "유튜브 영상 열기" $videoDirectUrl 652 190 180 44 $(RGB 205 65 55) | Out-Null
    Add-LinkButton $s "검색 결과로 보기" $videoUrl 652 248 180 44 $(RGB 87 126 83) | Out-Null
    Add-Text $s "PowerPoint에서 직접 삽입하려면:`r삽입 > 비디오 > 온라인 비디오 > 유튜브 주소 붙여넣기" 640 326 235 90 15 $(RGB 60 70 68) | Out-Null

    # 7
    $s = $pres.Slides.Add(7, $ppLayoutBlank)
    $s.Background.Fill.ForeColor.RGB = RGB 255 255 255
    Add-Title $s "정리와 출처" "백과사전 발표 마무리"
    Add-Text $s "핵심 정리" 78 156 260 28 22 $(RGB 18 27 28) $true | Out-Null
    Add-Bullets $s @(
        "담비는 노란 목과 긴 꼬리가 특징인 족제비과 동물",
        "산림에서 작은 동물을 사냥하고 사체를 처리하는 역할을 함",
        "서식지 연결성과 산림 보전이 담비 보호에 중요"
    ) 78 200 430 135 17 | Out-Null
    Add-Text $s "자료 링크" 586 156 260 28 22 $(RGB 18 27 28) $true | Out-Null
    Add-LinkButton $s "한국민족문화대백과" $encyclopediaUrl 586 204 210 40 | Out-Null
    Add-LinkButton $s "환경부 자료" $ministryUrl 586 256 210 40 $(RGB 43 92 78) | Out-Null
    Add-LinkButton $s "담비 기본 정보" $speciesUrl 586 308 210 40 $(RGB 91 96 103) | Out-Null
    Add-Text $s "이미지: PowerPoint 스톡 사진 느낌으로 생성한 교육용 이미지`r영상: YouTube 링크 사용" 82 438 760 44 13 $(RGB 103 114 111) | Out-Null

    $pres.SaveAs($outPath)
}
finally {
    $pres.Close()
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}

Write-Output $outPath
