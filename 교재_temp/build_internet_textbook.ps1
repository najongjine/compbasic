$ErrorActionPreference = 'Stop'

$outDir = 'D:\compbasic\교재_temp'
$imgDir = Join-Path $outDir 'images'
$outFile = Join-Path $outDir '컴퓨터기초_인터넷활용.hwpx'
$pdfFile = Join-Path $outDir '컴퓨터기초_인터넷활용_검토용.pdf'

$hwp = New-Object -ComObject HWPFrame.HwpObject
$window = $hwp.XHwpWindows.Item(0)
$window.Visible = $true
$script:hwpHandle = $window.WindowHandle
$hwp.HAction.Run('FileNew') | Out-Null

$C = @{
    Navy = $hwp.RGBColor(24, 54, 93)
    Blue = $hwp.RGBColor(38, 103, 172)
    Sky = $hwp.RGBColor(224, 239, 252)
    Green = $hwp.RGBColor(3, 163, 96)
    Yellow = $hwp.RGBColor(255, 234, 145)
    Orange = $hwp.RGBColor(231, 121, 44)
    Red = $hwp.RGBColor(192, 0, 0)
    Gray = $hwp.RGBColor(90, 90, 90)
    LightGray = $hwp.RGBColor(242, 244, 247)
    White = $hwp.RGBColor(255, 255, 255)
    Black = $hwp.RGBColor(25, 25, 25)
}

function Set-Char([double]$size = 11, [bool]$bold = $false, [int]$color = 0, [int]$shade = -1) {
    $hwp.HAction.GetDefault('CharShape', $hwp.HParameterSet.HCharShape.HSet) | Out-Null
    $cs = $hwp.HParameterSet.HCharShape
    $cs.FaceNameHangul = '맑은 고딕'
    $cs.FaceNameLatin = 'Arial'
    $cs.FaceNameHanja = '맑은 고딕'
    $cs.FaceNameJapanese = '맑은 고딕'
    $cs.FaceNameOther = 'Arial'
    $cs.RatioHangul = 100
    $cs.RatioLatin = 100
    $cs.RatioHanja = 100
    $cs.RatioJapanese = 100
    $cs.RatioOther = 100
    $cs.RatioSymbol = 100
    $cs.RatioUser = 100
    $cs.SizeHangul = 100
    $cs.SizeLatin = 100
    $cs.SizeHanja = 100
    $cs.SizeJapanese = 100
    $cs.SizeOther = 100
    $cs.SizeSymbol = 100
    $cs.SizeUser = 100
    $cs.Height = [int]($size * 100)
    $cs.Bold = $bold
    $cs.Italic = $false
    $cs.TextColor = $color
    if ($shade -ge 0) { $cs.ShadeColor = $shade } else { $cs.ShadeColor = $C.White }
    $hwp.HAction.Execute('CharShape', $cs.HSet) | Out-Null
}

function Set-Para([int]$align = 1, [int]$before = 0, [int]$after = 450, [int]$line = 155) {
    $hwp.HAction.GetDefault('ParaShape', $hwp.HParameterSet.HParaShape.HSet) | Out-Null
    $ps = $hwp.HParameterSet.HParaShape
    $ps.AlignType = $align
    $ps.PrevSpacing = $before
    $ps.NextSpacing = $after
    $ps.LineSpacingType = 0
    $ps.LineSpacing = $line
    $ps.KeepWithNext = $false
    $hwp.HAction.Execute('ParaShape', $ps.HSet) | Out-Null
}

function Put([string]$text) {
    $hwp.HAction.GetDefault('InsertText', $hwp.HParameterSet.HInsertText.HSet) | Out-Null
    $hwp.HParameterSet.HInsertText.Text = $text
    $hwp.HAction.Execute('InsertText', $hwp.HParameterSet.HInsertText.HSet) | Out-Null
}

function Para([string]$text = '', [double]$size = 11, [bool]$bold = $false, [int]$color = 0, [int]$align = 1, [int]$after = 450, [int]$shade = -1) {
    Set-Para $align 0 $after 155
    Set-Char $size $bold $color $shade
    Put $text
    $hwp.HAction.Run('BreakPara') | Out-Null
}

function Body([string]$text) { Para $text 11 $false $C.Black 1 500 }
function Step([string]$n, [string]$text) { Para ("  {0}  {1}  " -f $n, $text) 11 $true $C.Navy 1 360 $C.Sky }
function Tip([string]$text) { Para ("  알아두기  |  {0}  " -f $text) 10.5 $true $C.Black 1 500 $C.Yellow }
function Warn([string]$text) { Para ("  안전 주의  |  {0}  " -f $text) 10.5 $true $C.Red 1 500 $C.Yellow }
function Bullet([string]$text) { Para ("• " + $text) 10.8 $false $C.Black 1 280 }
function Small([string]$text) { Para $text 8.5 $false $C.Gray 1 300 }
function PageBreak { $hwp.HAction.Run('BreakPage') | Out-Null }

function UnitTitle([string]$number, [string]$title, [string]$subtitle) {
    Para ("UNIT {0}" -f $number) 11 $true $C.Blue 1 140
    Para $title 25 $true $C.Navy 1 220
    Para $subtitle 12 $false $C.Gray 1 650
}

function SubTitle([string]$title) {
    Para $title 16 $true $C.Blue 1 320
}

function Practice([string]$title, [string[]]$items) {
    Para ("  직접 해보기  |  {0}  " -f $title) 12 $true $C.White 1 360 $C.Blue
    foreach ($item in $items) { Bullet ("□ " + $item) }
    Para '메모: ________________________________________________________________' 9.5 $false $C.Gray 1 650
}

$script:firstImage = $true
function Add-Image([string]$name, [string]$caption, [string]$source, [double]$width = 150, [double]$height = 84.5) {
    $path = Join-Path $imgDir $name
    if (-not (Test-Path -LiteralPath $path)) {
        Para ("[" + $caption + ' 이미지 첨부]') 10 $true $C.Red 3 400 $C.Yellow
        return
    }
    Set-Para 3 0 200 130
    if ($script:firstImage) {
        $job = Start-Job -ArgumentList $script:hwpHandle -ScriptBlock {
            param($handle)
            Add-Type -AssemblyName UIAutomationClient
            Add-Type -AssemblyName UIAutomationTypes
            $root = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$handle)
            $buttonType = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                [System.Windows.Automation.ControlType]::Button
            )
            for ($i = 0; $i -lt 40; $i++) {
                Start-Sleep -Milliseconds 150
                $buttons = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $buttonType)
                foreach ($button in $buttons) {
                    if ($button.Current.Name -like '*모두 허용*') {
                        $invoke = $button.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                        $invoke.Invoke()
                        return
                    }
                }
            }
        }
        $hwp.InsertPicture($path, $true, 1, $false, $false, 0, $width, $height) | Out-Null
        Wait-Job $job -Timeout 8 | Out-Null
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        $script:firstImage = $false
    } else {
        $hwp.InsertPicture($path, $true, 1, $false, $false, 0, $width, $height) | Out-Null
    }
    $hwp.HAction.Run('BreakPara') | Out-Null
    Para ("▲ " + $caption) 9 $true $C.Gray 3 160
    Small ("화면 출처: {0}  |  캡처: 2026-08-14  |  화면 구성은 서비스 업데이트에 따라 달라질 수 있습니다." -f $source)
}

# 표지
Set-Para 3 0 300 160
Set-Char 12 $true $C.Blue
Put '컴퓨터 기초 실습 교재'
$hwp.HAction.Run('BreakPara') | Out-Null
Para '' 16 $false $C.Black 3 300
Para '인터넷 활용' 34 $true $C.Navy 3 240
Para '웹브라우저부터 검색·회원가입·카카오톡 설치까지' 16 $false $C.Blue 3 900
Para '보고, 따라 하고, 스스로 해보는 초보자용 워크북' 13 $true $C.Black 3 1200 $C.Sky
Para '대상: 컴퓨터와 인터넷을 처음 배우는 학습자' 11 $false $C.Gray 3 240
Para '환경: Windows 10/11 · Chrome 또는 Edge 기준' 11 $false $C.Gray 3 240
Para '제작일: 2026년 8월 14일' 10 $false $C.Gray 3 1100
Warn '교재의 가입 화면에는 예시 개인정보를 입력하지 않습니다. 실제 수업에서는 본인의 정보를 다른 사람에게 보여 주지 마세요.'
PageBreak

# 안내/목차
UnitTitle '00' '이 교재를 사용하는 방법' '한 단계씩 천천히, 화면과 글을 함께 보세요.'
SubTitle '학습 순서'
Step '1' '먼저 실제 화면 그림에서 버튼과 입력칸의 위치를 찾습니다.'
Step '2' '설명 문장을 읽고 마우스 클릭 또는 키보드 입력을 따라 합니다.'
Step '3' '각 단원의 [직접 해보기]를 수행하고 체크 표시를 합니다.'
Step '4' '실수했을 때는 당황하지 말고 뒤로 가기, 취소, 창 닫기를 사용합니다.'
Tip '웹사이트 화면은 수시로 바뀝니다. 버튼 모양보다 버튼에 적힌 글자를 확인하는 습관을 들이세요.'
SubTitle '차례'
Bullet 'UNIT 01 웹브라우저 알아보기'
Bullet 'UNIT 02 인터넷 검색하기'
Bullet 'UNIT 03 즐겨찾기(북마크) 사용하기'
Bullet 'UNIT 04 파일 다운로드하기'
Bullet 'UNIT 05 네이버 회원가입하기'
Bullet 'UNIT 06 구글 회원가입하기'
Bullet 'UNIT 07 PC 카카오톡 설치하기'
Bullet '부록: 인터넷 안전 수칙 · 문제 해결 · 종합 실습'
PageBreak

# UNIT 01
UnitTitle '01' '웹브라우저 알아보기' '인터넷의 문을 여는 프로그램과 기본 버튼 익히기'
SubTitle '1. 웹브라우저란?'
Body '웹브라우저는 인터넷에 있는 글, 사진, 동영상, 지도 같은 정보를 보여 주는 프로그램입니다. 대표적인 웹브라우저로 Google Chrome, Microsoft Edge, Naver Whale, Firefox가 있습니다. 어느 브라우저를 사용해도 주소 입력, 검색, 뒤로 가기 같은 기본 원리는 비슷합니다.'
Bullet '탭: 여러 웹페이지를 한 창에서 나누어 열어 두는 작은 제목표입니다.'
Bullet '주소 표시줄: www.naver.com 같은 인터넷 주소 또는 검색어를 입력하는 곳입니다.'
Bullet '뒤로/앞으로: 전에 보던 페이지 또는 다음 페이지로 이동합니다.'
Bullet '새로고침: 현재 페이지를 다시 불러옵니다. 화면이 멈추었을 때 유용합니다.'
Bullet '더보기 메뉴(⋮): 즐겨찾기, 다운로드, 인쇄, 설정 등의 기능을 엽니다.'
Para '[브라우저의 탭·뒤로·앞으로·새로고침·주소 표시줄·더보기 메뉴 이미지 첨부]' 10 $true $C.Red 3 420 $C.Yellow
Add-Image 'naver_home.png' '네이버 첫 화면: 가운데 검색창과 여러 서비스 바로가기를 확인합니다.' 'https://www.naver.com/'
SubTitle '2. 웹페이지 열기'
Step '①' '바탕 화면이나 작업 표시줄에서 Chrome 또는 Edge 아이콘을 더블클릭합니다.'
Step '②' '위쪽 주소 표시줄을 한 번 클릭합니다. 기존 글자가 파란색으로 선택됩니다.'
Step '③' 'www.naver.com을 입력하고 Enter 키를 누릅니다.'
Step '④' '페이지가 열리면 제목, 검색창, 메뉴, 스크롤 막대의 위치를 살펴봅니다.'
Tip '주소 표시줄을 빠르게 선택하려면 Ctrl+L을 누릅니다. 새 탭은 Ctrl+T, 현재 탭 닫기는 Ctrl+W입니다.'
SubTitle '3. 탭과 창 다루기'
Body '새로운 내용을 보면서 현재 페이지를 남겨 두고 싶다면 링크 위에서 마우스 오른쪽 버튼을 누른 뒤 [새 탭에서 링크 열기]를 선택합니다. 탭이 많아지면 제목을 보고 필요한 탭만 남기고 × 버튼으로 닫습니다.'
Practice '웹브라우저 기본 조작' @(
    '브라우저를 열고 주소 표시줄에 www.naver.com을 입력한다.',
    '새 탭을 열어 www.google.com을 입력한다.',
    '두 탭을 번갈아 눌러 보고 하나를 닫는다.',
    '뒤로 가기와 새로고침 버튼을 한 번씩 사용한다.'
)
PageBreak

# UNIT 02
UnitTitle '02' '인터넷 검색하기' '필요한 정보를 빠르고 정확하게 찾는 방법'
SubTitle '1. 검색어를 구체적으로 쓰기'
Body '검색창에는 질문 전체를 길게 쓰기보다 핵심 낱말을 2~4개 정도 조합해 입력하면 좋습니다. 예를 들어 [날씨]보다 [서울 오늘 날씨], [버스]보다 [서울 7016 버스 노선]처럼 지역·날짜·대상을 함께 적습니다.'
Bullet '좋은 검색어: 서울 오늘 날씨 / 종로구 보건소 운영시간 / KTX 승차권 예매 방법'
Bullet '너무 넓은 검색어: 날씨 / 병원 / 기차'
Bullet '따옴표 검색: 정확히 같은 문장을 찾고 싶을 때 [컴퓨터 기초]처럼 큰따옴표를 사용합니다.'
Bullet '제외 검색: 원치 않는 내용이 많다면 검색어 뒤에 -단어를 붙여 제외할 수 있습니다.'
Add-Image 'naver_search.png' '네이버에서 [서울 날씨]를 검색한 결과 화면' 'https://search.naver.com/'
SubTitle '2. 검색 결과 읽기'
Step '①' '검색 결과의 큰 제목을 읽고 내가 찾는 내용과 맞는지 확인합니다.'
Step '②' '제목 아래의 주소와 짧은 설명을 읽습니다. 기관·학교·기업의 공식 주소인지 살펴봅니다.'
Step '③' '광고 표시가 있는 결과와 일반 검색 결과를 구분합니다.'
Step '④' '중요한 정보는 한 곳만 믿지 말고 작성 날짜와 다른 출처를 함께 확인합니다.'
Warn '[당첨], [긴급], [지금 설치], [계정 정지]처럼 불안을 자극하는 문구를 보고 바로 누르지 마세요. 주소가 이상하면 창을 닫고 공식 사이트를 직접 검색하세요.'
SubTitle '3. 사진과 지도 검색'
Body '검색 결과 위쪽의 이미지, 지도, 뉴스, 동영상 탭을 누르면 종류별 결과를 볼 수 있습니다. 사진은 저작권이 있을 수 있으므로 개인 학습이 아닌 공개 자료에 사용할 때는 이용 조건을 확인합니다.'
Practice '생활 정보 찾기' @(
    '내일 우리 지역 날씨를 검색하고 최고·최저 기온을 적는다.',
    '가까운 주민센터의 공식 주소와 전화번호를 찾는다.',
    '같은 정보를 두 사이트에서 확인하고 서로 일치하는지 비교한다.',
    '광고 결과 하나와 일반 검색 결과 하나를 구분해 본다.'
)
PageBreak

# UNIT 03
UnitTitle '03' '즐겨찾기(북마크) 사용하기' '자주 가는 사이트를 저장하고 다시 찾기'
SubTitle '1. 즐겨찾기와 북마크'
Body 'Chrome에서는 북마크, Edge에서는 즐겨찾기라는 이름을 주로 사용합니다. 이름은 달라도 자주 방문하는 웹페이지 주소를 저장해 두었다가 다시 여는 기능은 같습니다.'
Add-Image 'chrome_bookmarks.png' 'Google Chrome 고객센터의 북마크 안내 화면' 'https://support.google.com/chrome/answer/188842?hl=ko'
SubTitle '2. 현재 페이지 저장하기'
Step '①' '저장할 웹페이지를 엽니다.'
Step '②' '주소 표시줄 오른쪽의 별표(☆)를 누릅니다. 단축키는 Ctrl+D입니다.'
Step '③' '이름을 짧고 알아보기 쉽게 바꿉니다. 예: [우리동네 주민센터].'
Step '④' '저장할 폴더를 선택하고 완료를 누릅니다.'
Tip '북마크바를 표시하거나 숨기려면 Ctrl+Shift+B를 누릅니다. 북마크가 많다면 생활, 건강, 학습, 쇼핑처럼 폴더를 나누세요.'
SubTitle '3. 저장한 페이지 찾고 정리하기'
Body '브라우저의 더보기 메뉴에서 북마크 및 목록 또는 즐겨찾기 메뉴를 엽니다. 필요 없는 항목은 삭제하고, 이름이 모호한 항목은 수정합니다. 주소가 바뀌어 열리지 않으면 검색으로 새 주소를 찾은 뒤 다시 저장합니다.'
Bullet '중요 사이트는 이름 앞에 번호를 붙이면 정렬하기 쉽습니다. 예: 01 은행, 02 병원.'
Bullet '개인정보가 들어 있는 로그인 완료 화면은 즐겨찾기에 저장하지 않는 것이 안전합니다.'
Bullet '공용 컴퓨터에서는 브라우저 동기화 로그인을 피하고, 사용 후 계정에서 로그아웃합니다.'
Practice '즐겨찾기 폴더 만들기' @(
    '[학습] 폴더를 만들고 네이버와 구글을 저장한다.',
    '저장한 이름을 [네이버 검색], [구글 검색]으로 바꾼다.',
    '북마크바를 표시하고 두 사이트를 다시 연다.',
    '연습용 북마크 하나를 삭제한다.'
)
PageBreak

# UNIT 04
UnitTitle '04' '파일 다운로드하기' '인터넷의 파일을 내 컴퓨터에 안전하게 저장하기'
SubTitle '1. 다운로드란?'
Body '다운로드는 인터넷에 있는 문서, 사진, 음악, 설치 파일 등을 내 컴퓨터로 가져와 저장하는 것입니다. 보통 다운로드 폴더에 저장되며, 브라우저 오른쪽 위의 다운로드 아이콘 또는 Ctrl+J로 받은 파일 목록을 확인할 수 있습니다.'
Add-Image 'chrome_downloads.png' 'Google Chrome 고객센터의 파일 다운로드 안내 화면' 'https://support.google.com/chrome/answer/95759?hl=ko'
SubTitle '2. 안전하게 저장하기'
Step '①' '공식 홈페이지인지 주소를 확인합니다. 검색 광고에서 곧바로 설치 파일을 받지 않습니다.'
Step '②' '다운로드 버튼을 누르기 전에 파일 이름과 확장자를 확인합니다.'
Step '③' '저장 위치를 선택할 수 있다면 찾기 쉬운 폴더를 선택하고 저장을 누릅니다.'
Step '④' '다운로드 목록에서 완료 여부를 확인한 뒤 파일을 엽니다.'
Bullet '문서: .pdf, .docx, .hwpx / 사진: .jpg, .png / 압축: .zip / 프로그램 설치: .exe, .msi'
Warn '.exe와 .msi는 프로그램을 설치하는 파일입니다. 출처를 확신할 수 없으면 실행하지 말고 선생님이나 담당자에게 확인하세요.'
SubTitle '3. 받은 파일 찾기'
Body '파일 탐색기를 열고 왼쪽의 다운로드 폴더를 선택합니다. 보기 방식을 [자세히]로 바꾸면 이름, 수정 날짜, 종류, 크기를 확인할 수 있습니다. 최근 받은 파일은 수정 날짜 순으로 정렬하면 빨리 찾을 수 있습니다.'
Para '[브라우저의 다운로드 목록과 Windows 파일 탐색기 다운로드 폴더 이미지 첨부]' 10 $true $C.Red 3 420 $C.Yellow
SubTitle '4. 압축 파일 풀기'
Step '①' '.zip 파일을 마우스 오른쪽 버튼으로 누릅니다.'
Step '②' '모두 압축 풀기를 선택합니다.'
Step '③' '압축을 풀 위치를 확인하고 압축 풀기를 누릅니다.'
Step '④' '새로 만들어진 폴더를 열어 필요한 파일을 확인합니다.'
Practice '다운로드 폴더 정리' @(
    '브라우저에서 Ctrl+J를 눌러 다운로드 목록을 연다.',
    '파일 탐색기에서 다운로드 폴더를 연다.',
    '파일을 종류와 날짜 기준으로 정렬해 본다.',
    '출처가 불분명한 실행 파일은 열지 않고 확인을 요청한다.'
)
PageBreak

# UNIT 05
UnitTitle '05' '네이버 회원가입하기' '약관 확인부터 아이디·비밀번호·휴대전화 인증까지'
SubTitle '1. 가입 전에 준비하기'
Bullet '본인 명의 휴대전화 또는 본인 인증에 사용할 기기'
Bullet '다른 사람이 쉽게 짐작할 수 없는 아이디 후보 2~3개'
Bullet '영문 대·소문자, 숫자, 특수문자를 조합한 긴 비밀번호'
Bullet '복구에 사용할 정보를 정확히 기억하거나 안전한 곳에 보관할 준비'
Add-Image 'naver_signup.png' '네이버 회원가입 약관 동의 첫 화면' 'https://nid.naver.com/user2/join/agree?lang=ko_KR'
SubTitle '2. 약관 동의하기'
Step '①' '네이버 첫 화면 오른쪽의 회원가입을 누릅니다.'
Step '②' '필수와 선택 항목을 구분해 읽습니다. 가입에 꼭 필요하지 않은 선택 동의는 원하지 않으면 체크하지 않아도 됩니다.'
Step '③' '이용약관과 개인정보 수집·이용 내용을 확인한 뒤 다음을 누릅니다.'
Warn '수업 중에는 실제 주민등록번호, 비밀번호, 인증번호를 큰 소리로 말하거나 화면을 촬영하지 않습니다.'
SubTitle '3. 아이디와 비밀번호 만들기'
Body '아이디는 로그인할 때 사용하는 이름입니다. 실명, 생년월일, 휴대전화 번호 전체를 그대로 쓰지 않는 것이 좋습니다. 비밀번호는 사이트마다 다르게 만들고, 최소 12자 이상을 권장합니다.'
Bullet '피해야 할 비밀번호: 12345678, qwerty, 생년월일, 전화번호, 아이디와 같은 문자열'
Bullet '권장 방식: 서로 관련 없는 단어 3개 이상 + 숫자 + 기호를 조합하여 본인만 기억하기'
Step '①' '아이디 입력칸에 원하는 아이디를 입력하고 사용 가능 여부를 확인합니다.'
Step '②' '비밀번호와 비밀번호 확인 칸에 같은 내용을 입력합니다.'
Step '③' '이름과 생년월일 등 요청 정보를 정확히 입력합니다.'
Step '④' '휴대전화 번호를 입력하고 인증번호 받기를 누릅니다.'
Step '⑤' '문자로 온 인증번호를 제한 시간 안에 입력하고 확인합니다.'
Para '[네이버 아이디·비밀번호 입력 화면 이미지 첨부]' 10 $true $C.Red 3 300 $C.Yellow
Para '[네이버 휴대전화 인증 화면 이미지 첨부]' 10 $true $C.Red 3 500 $C.Yellow
SubTitle '4. 가입 후 꼭 할 일'
Bullet '로그인이 되는지 확인하고, 사용하지 않을 때는 로그아웃합니다.'
Bullet '계정 보안 설정에서 2단계 인증을 켜면 계정 도용을 막는 데 도움이 됩니다.'
Bullet '복구 전화번호와 이메일을 최신 정보로 유지합니다.'
Practice '네이버 가입 순서 점검' @(
    '필수 약관과 선택 약관을 구분해 설명할 수 있다.',
    '안전한 비밀번호 조건을 세 가지 이상 말할 수 있다.',
    '인증번호를 다른 사람에게 알려 주면 안 되는 이유를 안다.',
    '공용 컴퓨터에서 로그아웃하는 방법을 안다.'
)
PageBreak

# UNIT 06
UnitTitle '06' '구글 회원가입하기' 'Gmail과 YouTube 등에서 사용하는 Google 계정 만들기'
SubTitle '1. Google 계정의 쓰임'
Body 'Google 계정 하나로 Gmail, YouTube, Google Drive, 지도 등 여러 서비스를 사용할 수 있습니다. 계정 만들기 화면에서 개인용 계정을 선택하고 이름, 생년월일, 사용자 이름, 비밀번호 등을 차례로 입력합니다.'
Add-Image 'google_signup.png' 'Google 계정 만들기 첫 화면: 성과 이름 입력 단계' 'https://accounts.google.com/signup?hl=ko'
SubTitle '2. 가입 순서'
Step '①' 'accounts.google.com을 열고 계정 만들기를 누른 뒤 개인용을 선택합니다.'
Step '②' '성과 이름을 입력합니다. 화면에 보이는 순서대로 입력하고 다음을 누릅니다.'
Step '③' '생년월일과 성별을 입력합니다. 서비스의 연령 기준에 따라 선택 가능한 기능이 달라질 수 있습니다.'
Step '④' '추천 Gmail 주소를 고르거나 직접 사용자 이름을 만듭니다.'
Step '⑤' '안전한 비밀번호를 입력하고 같은 비밀번호를 한 번 더 입력합니다.'
Step '⑥' '필요한 경우 휴대전화 인증을 하고 복구 정보를 설정합니다.'
Step '⑦' '개인정보 보호 및 약관 내용을 읽고 동의 범위를 확인한 뒤 가입을 마칩니다.'
Para '[Google 사용자 이름과 비밀번호 설정 화면 이미지 첨부]' 10 $true $C.Red 3 300 $C.Yellow
Para '[Google 복구 전화번호·약관 확인 화면 이미지 첨부]' 10 $true $C.Red 3 500 $C.Yellow
SubTitle '3. Gmail 주소 이해하기'
Body '사용자 이름이 hong.gildong이라면 Gmail 주소는 hong.gildong@gmail.com이 됩니다. 영문자와 숫자, 마침표 등을 사용할 수 있지만 이미 다른 사람이 사용 중인 주소는 선택할 수 없습니다. 업무용과 개인용 주소를 구분하면 관리가 편리합니다.'
Warn 'Google 계정 비밀번호와 휴대전화로 받은 인증 코드는 누구에게도 보내지 마세요. Google 직원을 사칭해도 알려 주면 안 됩니다.'
SubTitle '4. 계정 전환과 로그아웃'
Body 'Google 화면 오른쪽 위의 프로필 사진 또는 동그란 글자 아이콘을 누르면 현재 로그인한 계정을 확인할 수 있습니다. 여러 계정을 사용한다면 메일을 보내기 전에 현재 계정이 맞는지 확인합니다. 공용 컴퓨터에서는 로그아웃한 뒤 브라우저 창을 모두 닫습니다.'
Practice 'Google 계정 안전 점검' @(
    'Gmail 주소의 사용자 이름 부분과 @gmail.com 부분을 구분한다.',
    '복구 전화번호와 복구 이메일의 역할을 설명한다.',
    '프로필 아이콘에서 현재 로그인 계정을 확인한다.',
    '공용 컴퓨터 사용 후 로그아웃하고 창을 닫는다.'
)
PageBreak

# UNIT 07
UnitTitle '07' 'PC 카카오톡 설치하기' '공식 사이트에서 안전하게 내려받고 로그인하기'
SubTitle '1. 공식 다운로드 페이지 찾기'
Body '검색창에 [카카오톡 공식 다운로드]를 입력하고, 결과의 주소가 kakaocorp.com인지 확인합니다. 광고나 파일 공유 사이트에서 설치 파일을 받지 않습니다.'
Add-Image 'kakao_download.png' '카카오톡 서비스 소개 공식 화면' 'https://www.kakaocorp.com/page/service/service/KakaoTalk?lang=ko'
Add-Image 'kakao_all_services.png' '카카오톡 공식 다운로드 화면: Windows 버튼을 선택합니다.' 'https://www.kakaocorp.com/page/service/all'
SubTitle '2. Windows용 설치 파일 받기'
Step '①' '공식 다운로드 화면에서 Windows를 누릅니다.'
Step '②' '브라우저 다운로드 목록에서 KakaoTalk_Setup.exe와 같은 설치 파일 이름을 확인합니다.'
Step '③' '다운로드가 끝나면 파일 열기를 누르거나 다운로드 폴더에서 설치 파일을 더블클릭합니다.'
Warn 'Windows에서 [이 앱이 디바이스를 변경하도록 허용하시겠어요?]라고 물으면 게시자가 Kakao Corp.인지 확인한 뒤 진행합니다.'
SubTitle '3. 설치 마법사 따라 하기'
Step '①' '언어가 한국어인지 확인하고 다음을 누릅니다.'
Step '②' '설치 폴더는 특별한 이유가 없다면 기본 위치를 사용합니다.'
Step '③' '바로가기 만들기 등 선택 항목을 확인하고 설치를 누릅니다.'
Step '④' '설치가 완료되면 카카오톡 실행을 선택하고 마침을 누릅니다.'
Para '[Windows 카카오톡 설치 마법사의 언어 선택·설치 위치·설치 완료 화면 이미지 첨부]' 10 $true $C.Red 3 500 $C.Yellow
SubTitle '4. 로그인과 인증'
Body 'PC 카카오톡은 카카오계정 이메일 또는 전화번호와 비밀번호로 로그인합니다. 처음 사용하는 컴퓨터라면 보안 인증이 필요할 수 있습니다. 내 컴퓨터가 아니라면 자동 로그인이나 잠금모드 해제 상태를 남기지 않습니다.'
Bullet '내 PC 인증받기: 개인이 계속 사용하는 안전한 컴퓨터에서만 선택합니다.'
Bullet '1회용 인증받기: PC방, 교육장, 공용 컴퓨터처럼 잠시 사용하는 환경에서 선택합니다.'
Bullet '로그인 후 설정에서 잠금모드, 자동 로그인, 알림 표시 범위를 확인합니다.'
Para '[PC 카카오톡 로그인·보안 인증 화면 이미지 첨부]' 10 $true $C.Red 3 450 $C.Yellow
SubTitle '5. 설치 후 기본 사용'
Step '①' '친구 목록과 채팅 목록 탭을 구분합니다.'
Step '②' '검색창에서 친구 이름을 찾고 채팅방을 엽니다.'
Step '③' '파일을 보내기 전 받는 사람과 파일 이름을 다시 확인합니다.'
Step '④' '사용이 끝나면 공용 컴퓨터에서는 반드시 로그아웃합니다.'
Practice '카카오톡 설치 순서 설명하기' @(
    '공식 카카오 사이트 주소를 확인한다.',
    'Windows 설치 파일을 다운로드 폴더에서 찾는다.',
    '개인 PC 인증과 1회용 인증의 차이를 설명한다.',
    '공용 컴퓨터에서 자동 로그인을 사용하지 않는다.'
)
PageBreak

# 부록
UnitTitle '부록' '인터넷 안전 수칙과 문제 해결' '어려운 상황에서도 침착하게 확인하기'
SubTitle '1. 꼭 지켜야 할 인터넷 안전 수칙 10가지'
Bullet '비밀번호는 길고 다르게 만들며 다른 사람에게 알려 주지 않습니다.'
Bullet '문자로 받은 인증번호, 보안코드, OTP는 누구에게도 전달하지 않습니다.'
Bullet '출처가 불분명한 링크와 첨부 파일은 열지 않습니다.'
Bullet '주소 표시줄의 사이트 주소와 자물쇠 표시를 확인합니다.'
Bullet '공용 컴퓨터에서는 로그인 유지와 자동 저장을 선택하지 않습니다.'
Bullet '앱과 프로그램은 공식 홈페이지 또는 공식 앱스토어에서 받습니다.'
Bullet '개인정보를 요구하는 창이 갑자기 나타나면 먼저 창을 닫고 공식 고객센터를 확인합니다.'
Bullet '결제나 송금을 재촉하는 메시지는 가족·기관에 다른 방법으로 다시 확인합니다.'
Bullet '브라우저와 운영체제 보안 업데이트를 미루지 않습니다.'
Bullet '실수로 눌렀더라도 숨기지 말고 즉시 선생님, 가족, 담당자에게 알립니다.'
SubTitle '2. 자주 생기는 문제와 해결 순서'
Step '화면이 멈춤' '잠시 기다린 뒤 새로고침을 누릅니다. 계속 멈추면 탭을 닫고 다시 엽니다.'
Step '글자가 너무 작음' 'Ctrl 키를 누른 채 + 키로 확대하고, Ctrl+0으로 원래 크기로 돌아갑니다.'
Step '파일을 못 찾음' 'Ctrl+J로 다운로드 목록을 열고 [폴더 열기]를 선택합니다.'
Step '로그인이 안 됨' '한글/영문 상태와 Caps Lock을 확인하고, 비밀번호 찾기는 공식 사이트에서 진행합니다.'
Step '이상한 광고가 뜸' '광고 안의 버튼을 누르지 말고 탭을 닫습니다. 설치된 프로그램과 알림 권한을 점검합니다.'
Step '사이트가 진짜인지 의심됨' '창을 닫고 주소를 직접 입력하거나 공식 기관 대표번호로 확인합니다.'
SubTitle '3. 종합 실습'
Practice '나만의 생활 정보 모음 만들기' @(
    '브라우저를 열고 우리 지역 주민센터 공식 사이트를 검색한다.',
    '검색 결과의 주소와 기관명을 확인한다.',
    '해당 페이지를 [생활] 즐겨찾기 폴더에 저장한다.',
    '공개된 안내문 PDF가 있다면 출처와 확장자를 확인해 다운로드한다.',
    '다운로드 폴더에서 파일을 찾고 이름을 알아보기 쉽게 바꾼다.',
    '수업이 끝나면 로그인된 계정에서 로그아웃하고 불필요한 탭을 닫는다.'
)
SubTitle '4. 학습 확인 문제'
Bullet '1) 주소 표시줄을 선택하는 단축키는 무엇인가요? ____________________'
Bullet '2) 새 탭을 여는 단축키는 무엇인가요? _____________________________'
Bullet '3) 즐겨찾기와 북마크는 어떤 기능인가요? __________________________'
Bullet '4) 실행 파일의 대표 확장자 두 가지를 쓰세요. ______________________'
Bullet '5) 선택 약관은 반드시 동의해야 하나요? ____________________________'
Bullet '6) 인증번호를 다른 사람에게 알려 주면 안 되는 이유는 무엇인가요? ____'
Bullet '7) 공용 컴퓨터에서 카카오톡 로그인 시 어떤 인증이 알맞나요? __________'
Bullet '8) 수상한 다운로드 파일을 발견하면 어떻게 해야 하나요? ______________'
Tip '정답 예: 1) Ctrl+L  2) Ctrl+T  3) 웹페이지 주소 저장  4) .exe, .msi  5) 아니요  6) 계정 탈취 위험  7) 1회용 인증  8) 실행하지 않고 확인 요청'
PageBreak

UnitTitle '마무리' '수고하셨습니다!' '인터넷은 천천히 확인하는 사람이 가장 안전하게 잘 사용할 수 있습니다.'
Para '오늘 내가 할 수 있게 된 것' 18 $true $C.Green 1 500
Bullet '□ 웹브라우저의 주소 표시줄과 탭을 사용할 수 있다.'
Bullet '□ 구체적인 검색어로 필요한 정보를 찾을 수 있다.'
Bullet '□ 즐겨찾기를 저장하고 폴더로 정리할 수 있다.'
Bullet '□ 다운로드한 파일의 위치와 확장자를 확인할 수 있다.'
Bullet '□ 네이버와 구글 회원가입 순서를 설명할 수 있다.'
Bullet '□ 카카오톡을 공식 사이트에서 안전하게 설치할 수 있다.'
Bullet '□ 비밀번호와 인증번호를 안전하게 지킬 수 있다.'
Para '' 11 $false $C.Black 1 700
Para '다음에 연습할 내용: _______________________________________________' 11 $false $C.Gray 1 500
Para '도움을 받을 사람/연락처: __________________________________________' 11 $false $C.Gray 1 500
Small '본 교재의 웹 화면은 교육 목적의 예시이며, 각 서비스의 상표와 화면 저작권은 해당 권리자에게 있습니다.'

$hwp.HAction.GetDefault('PageNumPos', $hwp.HParameterSet.HPageNumPos.HSet) | Out-Null
$hwp.HParameterSet.HPageNumPos.DrawPos = 5
$hwp.HAction.Execute('PageNumPos', $hwp.HParameterSet.HPageNumPos.HSet) | Out-Null

$hwp.SetTitleName('컴퓨터 기초 - 인터넷 활용') | Out-Null
$saved = $hwp.SaveAs($outFile, 'HWPX', '')
if (-not $saved) { throw 'HWPX 저장에 실패했습니다.' }

# 한글에서 검토용 PDF도 직접 내보낸다. 실패해도 HWPX 결과는 유지한다.
try { $hwp.SaveAs($pdfFile, 'PDF', '') | Out-Null } catch { }
$hwp.Quit()

Get-Item -LiteralPath $outFile
if (Test-Path -LiteralPath $pdfFile) { Get-Item -LiteralPath $pdfFile }
