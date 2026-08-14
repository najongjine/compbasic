$ErrorActionPreference = 'Stop'

$outDir = $PSScriptRoot
$imgDir = Join-Path $outDir 'keyboard_mouse_images'
$outFile = Join-Path $outDir '컴퓨터기초_키보드와마우스.hwpx'
$pdfFile = Join-Path $outDir '컴퓨터기초_키보드와마우스_검토용.pdf'

$hwp = New-Object -ComObject HWPFrame.HwpObject
$window = $hwp.XHwpWindows.Item(0)
$window.Visible = $true
$script:hwpHandle = $window.WindowHandle
$hwp.HAction.Run('FileNew') | Out-Null

$C = @{
    Navy = $hwp.RGBColor(24, 54, 93)
    Blue = $hwp.RGBColor(38, 103, 172)
    Sky = $hwp.RGBColor(224, 239, 252)
    Green = $hwp.RGBColor(23, 142, 104)
    Mint = $hwp.RGBColor(223, 245, 238)
    Yellow = $hwp.RGBColor(255, 239, 174)
    Orange = $hwp.RGBColor(230, 126, 34)
    Peach = $hwp.RGBColor(252, 232, 215)
    Red = $hwp.RGBColor(190, 45, 45)
    Gray = $hwp.RGBColor(90, 96, 105)
    LightGray = $hwp.RGBColor(241, 244, 247)
    White = $hwp.RGBColor(255, 255, 255)
    Black = $hwp.RGBColor(28, 31, 35)
}

function Set-Char([double]$size = 11, [bool]$bold = $false, [int]$color = 0, [int]$shade = -1) {
    $hwp.HAction.GetDefault('CharShape', $hwp.HParameterSet.HCharShape.HSet) | Out-Null
    $cs = $hwp.HParameterSet.HCharShape
    $cs.FaceNameHangul = '맑은 고딕'
    $cs.FaceNameLatin = 'Arial'
    $cs.FaceNameHanja = '맑은 고딕'
    $cs.FaceNameJapanese = '맑은 고딕'
    $cs.FaceNameOther = 'Arial'
    $cs.RatioHangul = 100; $cs.RatioLatin = 100; $cs.RatioHanja = 100
    $cs.RatioJapanese = 100; $cs.RatioOther = 100; $cs.RatioSymbol = 100; $cs.RatioUser = 100
    $cs.SizeHangul = 100; $cs.SizeLatin = 100; $cs.SizeHanja = 100
    $cs.SizeJapanese = 100; $cs.SizeOther = 100; $cs.SizeSymbol = 100; $cs.SizeUser = 100
    $cs.Height = [int]($size * 100)
    $cs.Bold = $bold
    $cs.Italic = $false
    $cs.TextColor = $color
    if ($shade -ge 0) { $cs.ShadeColor = $shade } else { $cs.ShadeColor = $C.White }
    $hwp.HAction.Execute('CharShape', $cs.HSet) | Out-Null
}

function Set-Para([int]$align = 1, [int]$before = 0, [int]$after = 420, [int]$line = 150) {
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

function Para([string]$text = '', [double]$size = 11, [bool]$bold = $false, [int]$color = 0, [int]$align = 1, [int]$after = 420, [int]$shade = -1) {
    Set-Para $align 0 $after 150
    Set-Char $size $bold $color $shade
    Put $text
    $hwp.HAction.Run('BreakPara') | Out-Null
}

function Body([string]$text) { Para $text 11 $false $C.Black 1 480 }
function Bullet([string]$text) { Para ('● ' + $text) 10.7 $false $C.Black 1 250 }
function Check([string]$text) { Para ('□ ' + $text) 10.7 $false $C.Black 1 260 }
function Step([string]$n, [string]$text) { Para (('  {0}단계  {1}  ' -f $n, $text)) 11 $true $C.Navy 1 330 $C.Sky }
function Tip([string]$text) { Para ('  알아두기  |  ' + $text + '  ') 10.4 $true $C.Black 1 430 $C.Yellow }
function Caution([string]$text) { Para ('  주의  |  ' + $text + '  ') 10.4 $true $C.Red 1 430 $C.Peach }
function Key([string]$keys, [string]$meaning) { Para (('  {0,-20}  {1}  ' -f $keys, $meaning)) 10.5 $true $C.Navy 1 210 $C.LightGray }
function Blank([string]$label) { Para ($label + '  ____________________________________________________________') 10.5 $false $C.Gray 1 340 }
function Small([string]$text) { Para $text 8.5 $false $C.Gray 1 250 }
function PageBreak { $hwp.HAction.Run('BreakPage') | Out-Null }

function UnitTitle([string]$number, [string]$title, [string]$subtitle) {
    Para ('UNIT ' + $number) 11 $true $C.Blue 1 100
    Para $title 25 $true $C.Navy 1 190
    Para $subtitle 11.5 $false $C.Gray 1 500
}

function SubTitle([string]$title) { Para $title 16 $true $C.Blue 1 300 }

function Goal([string[]]$items) {
    Para '  이 단원의 학습 목표  ' 11.5 $true $C.White 1 300 $C.Blue
    foreach ($item in $items) { Check $item }
}

function Practice([string]$title, [string[]]$items) {
    Para ('  직접 해보기  |  ' + $title + '  ') 12 $true $C.White 1 320 $C.Green
    foreach ($item in $items) { Check $item }
}

$script:firstImage = $true
function Add-Image([string]$name, [string]$caption, [double]$width, [double]$height) {
    $path = Join-Path $imgDir $name
    if (-not (Test-Path -LiteralPath $path)) {
        Para ('[' + $caption + ' 이미지 첨부]') 10.5 $true $C.Red 3 400 $C.Yellow
        return
    }
    Set-Para 3 0 180 125
    if ($script:firstImage) {
        $job = Start-Job -ArgumentList $script:hwpHandle -ScriptBlock {
            param($handle)
            Add-Type -AssemblyName UIAutomationClient
            Add-Type -AssemblyName UIAutomationTypes
            $root = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$handle)
            $condition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                [System.Windows.Automation.ControlType]::Button
            )
            for ($i = 0; $i -lt 60; $i++) {
                Start-Sleep -Milliseconds 150
                $buttons = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
                foreach ($button in $buttons) {
                    if ($button.Current.Name -like '*모두 허용*') {
                        $button.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
                        return
                    }
                }
            }
        }
        $hwp.InsertPicture($path, $true, 1, $false, $false, 0, $width, $height) | Out-Null
        Wait-Job $job -Timeout 10 | Out-Null
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        $script:firstImage = $false
    } else {
        $hwp.InsertPicture($path, $true, 1, $false, $false, 0, $width, $height) | Out-Null
    }
    $hwp.HAction.Run('BreakPara') | Out-Null
    Para ('그림 | ' + $caption) 9 $true $C.Gray 3 150
    Small '실제 Windows 화면을 직접 캡처한 예시입니다. Windows 버전과 설정에 따라 모양은 조금 다를 수 있습니다.'
}

# 표지
Set-Para 3 0 260 160
Set-Char 11 $true $C.Blue
Put '컴퓨터 기초 실습 교재'
$hwp.HAction.Run('BreakPara') | Out-Null
Para '' 12 $false $C.Black 3 220
Para '키보드와 마우스' 32 $true $C.Navy 3 220
Para '한글·영문 입력부터 특수문자, 단축키, 마우스 조작까지' 15 $false $C.Blue 3 750
Para '보고 → 따라 하고 → 스스로 확인하는 초보자용 워크북' 12.5 $true $C.Black 3 720 $C.Sky
Para '학습자 이름  ____________________________' 11 $false $C.Gray 3 300
Para '수업 날짜     ____________________________' 11 $false $C.Gray 3 650
Para '기준 환경: Windows 10/11 · 한컴 한글' 9.5 $false $C.Gray 3 180
Para '제작일: 2026년 8월 14일' 9.5 $false $C.Gray 3 180
PageBreak

# 안내와 차례
UnitTitle '00' '이 교재를 사용하는 방법' '그림을 먼저 보고, 한 단계씩 직접 조작해 보세요.'
SubTitle '학습 순서'
Step '1' '실제 화면에서 버튼과 입력 위치를 찾습니다.'
Step '2' '설명을 읽고 같은 동작을 천천히 따라 합니다.'
Step '3' '직접 해보기의 체크 상자에 완료 표시를 합니다.'
Step '4' '틀렸다면 Esc 또는 Ctrl+Z로 되돌리고 다시 시도합니다.'
Tip '키보드와 마우스는 빠르게 하는 것보다 정확하게 하는 것이 먼저입니다.'
SubTitle '차례'
Key 'UNIT 01' '키보드와 친해지기 · 한글/영문 입력'
Key 'UNIT 02' '특수문자 입력하기'
Key 'UNIT 03' '자주 쓰는 단축키'
Key 'UNIT 04' '마우스 조작하기'
Key '마무리' '종합 실습 · 문제 해결 · 확인 문제'
PageBreak

# 1단원 개요
UnitTitle '01' '키보드와 친해지기' '키의 위치와 역할을 알고 한글과 영어를 번갈아 입력합니다.'
Goal @('문자 키와 기능 키의 차이를 말할 수 있다.','한/영 키로 입력 언어를 바꿀 수 있다.','Backspace와 Delete로 잘못 입력한 글자를 고칠 수 있다.')
SubTitle '키보드의 다섯 구역'
Key '문자·숫자 키' '글자, 숫자, 기호를 입력합니다.'
Key '기능 키 F1~F12' '프로그램마다 정해진 기능을 실행합니다.'
Key '편집 키' 'Insert, Delete, Home, End로 문서를 편집합니다.'
Key '방향 키' '커서를 위·아래·왼쪽·오른쪽으로 움직입니다.'
Key '숫자 키패드' '숫자를 빠르게 입력합니다. Num Lock을 확인하세요.'
Caution '키를 세게 누를 필요는 없습니다. 한 번 눌렀다가 바로 손을 떼세요.'
PageBreak

# 한영 전환 설명
UnitTitle '01-1' '한글과 영어 입력 전환' '현재 입력 언어를 확인한 뒤 글자를 입력합니다.'
SubTitle '한/영 키로 바꾸기'
Step '1' '문서의 글자를 입력할 곳을 한 번 클릭합니다.'
Step '2' '키보드 아래쪽의 한/영 키를 한 번 누릅니다.'
Step '3' '화면 오른쪽 아래의 입력 표시가 가/한 또는 A로 바뀌는지 확인합니다.'
Step '4' '짧은 단어를 입력해 실제 입력 언어를 확인합니다.'
Key '한글 상태' 'ㄱ + ㅏ → 가 / ㅎ + ㅏ + ㄴ → 한'
Key '영문 상태' 'H + e + l + l + o → Hello'
Tip '영문 대문자는 Shift를 누른 채 문자 키를 누르거나 Caps Lock을 사용합니다.'
SubTitle '잘못 입력했을 때'
Key 'Backspace' '커서 왼쪽 글자를 지웁니다.'
Key 'Delete' '커서 오른쪽 글자를 지웁니다.'
Key 'Enter' '문단을 끝내고 다음 줄로 이동합니다.'
Key 'Spacebar' '글자 사이에 한 칸을 띄웁니다.'
PageBreak

# 실제 메모장 입력 화면
UnitTitle '01-2' '실제 화면으로 입력 확인' '메모장에서 한글과 영어가 함께 입력된 모습을 살펴봅니다.'
Add-Image '01_input_notepad.png' '메모장에 입력한 한글·영문·숫자·특수문자 예시' 165 88
SubTitle '화면에서 확인할 점'
Bullet '한글 문장과 영문 문장이 각각 올바른 글자 모양으로 보입니다.'
Bullet '문장 사이에는 Enter를 눌러 빈 줄을 만들었습니다.'
Bullet '영문 문장의 첫 글자 H와 I는 Shift를 이용한 대문자입니다.'
Bullet '날짜의 하이픈(-)과 문장 끝 마침표(.)도 키보드로 입력했습니다.'
PageBreak

# 입력 실습
UnitTitle '01-3' '한글·영문 입력 실습' '한컴 한글의 새 문서에서 아래 내용을 그대로 입력해 보세요.'
Practice '기본 문장 입력' @(
    '한글 상태에서 “안녕하세요. 반갑습니다.”를 입력했다.',
    '한/영 키를 누르고 “Nice to meet you.”를 입력했다.',
    'Enter를 눌러 다음 줄로 이동했다.',
    'Backspace로 일부 글자를 지운 뒤 다시 입력했다.',
    'Shift를 이용해 영문 대문자 A, B, C를 입력했다.'
)
Para '  따라 입력할 문장  ' 11.5 $true $C.White 1 260 $C.Blue
Key '1행' '오늘은 키보드 입력을 연습합니다.'
Key '2행' 'I can type Korean and English.'
Key '3행' '전화번호: 010-1234-5678'
Blank '입력하면서 어려웠던 점'
Blank '다시 연습할 키'
PageBreak

# 특수문자 개요
UnitTitle '02' '특수문자 입력하기' '문장에 필요한 기호를 키보드와 문자표에서 찾습니다.'
Goal @('Shift와 숫자 키를 조합해 기호를 입력할 수 있다.','문자표에서 키보드에 없는 기호를 찾을 수 있다.','기호의 쓰임에 맞게 띄어쓰기와 위치를 정할 수 있다.')
SubTitle '숫자 키 위의 기호'
Key 'Shift + 1' '!  느낌표'
Key 'Shift + 2' '@  골뱅이'
Key 'Shift + 3' '#  샵·번호 표시'
Key 'Shift + 4' '$  달러 표시'
Key 'Shift + 5' '%  퍼센트'
Key 'Shift + 7' '&  앤드 기호'
Key 'Shift + 8' '*  별표'
Key 'Shift + 9 / 0' '( )  괄호'
Tip '노트북이나 키보드 종류에 따라 키에 인쇄된 기호 위치가 다를 수 있습니다. 키의 윗부분을 확인하세요.'
PageBreak

# 문자표 화면
UnitTitle '02-1' '문자표에서 기호 찾기' '키보드에 없는 ©, ㎡, → 같은 기호는 문자표를 이용합니다.'
Add-Image '02_special_charmap.png' 'Windows 문자표에서 여러 기호를 살펴보는 실제 화면' 165 88
Step '1' 'Windows 검색에서 “문자표”를 찾아 실행합니다.'
Step '2' '원하는 글꼴을 고르고 기호를 한 번 클릭합니다.'
Step '3' '선택 버튼을 누른 뒤 복사 버튼을 누릅니다.'
Step '4' '한글 문서로 돌아와 Ctrl+V로 붙여 넣습니다.'
Caution '글꼴에 따라 사용할 수 없는 기호가 있습니다. 붙여 넣은 뒤 네모(□)로 보이면 다른 글꼴을 선택하세요.'
PageBreak

# 특수문자 실습
UnitTitle '02-2' '특수문자 실습' '기호의 이름과 쓰임을 생각하면서 입력합니다.'
Practice '기호가 포함된 문장 만들기' @(
    '“완료율 80%”를 입력했다.',
    '“문의: help@example.com”을 입력했다.',
    '“준비물(필수): 키보드, 마우스”를 입력했다.',
    '“주의! 전원을 먼저 확인하세요.”를 입력했다.',
    '문자표에서 © 또는 →를 골라 붙여 넣었다.'
)
SubTitle '헷갈리기 쉬운 기호'
Key '-' '하이픈: 단어 또는 번호를 연결합니다.'
Key '_' '밑줄: 전자우편 주소나 파일 이름에 쓰입니다.'
Key '/' '슬래시: 날짜나 선택 항목을 나눌 때 씁니다.'
Key '\\' '역슬래시: Windows 경로에서 볼 수 있습니다.'
Key ':' '콜론: 제목과 내용, 시와 분을 나눕니다.'
Blank '내가 찾은 새로운 특수문자'
PageBreak

# 단축키 개요
UnitTitle '03' '자주 쓰는 단축키' '메뉴를 찾지 않고 두세 개의 키를 함께 눌러 작업합니다.'
Goal @('Ctrl 조합 단축키를 바른 순서로 누를 수 있다.','복사·붙여넣기와 저장·실행 취소를 사용할 수 있다.','프로그램 전환과 창 닫기 단축키를 구분할 수 있다.')
SubTitle '단축키를 누르는 방법'
Step '1' '먼저 Ctrl, Alt 또는 Shift 같은 보조 키를 누른 채 유지합니다.'
Step '2' '두 번째 키를 한 번 누릅니다.'
Step '3' '두 키에서 손을 모두 뗍니다.'
Key 'Ctrl + C' '선택한 내용을 복사합니다.'
Key 'Ctrl + X' '선택한 내용을 잘라냅니다.'
Key 'Ctrl + V' '복사하거나 잘라낸 내용을 붙여 넣습니다.'
Key 'Ctrl + Z' '바로 전 작업을 취소합니다.'
Key 'Ctrl + S' '현재 문서를 저장합니다.'
Key 'Ctrl + A' '문서의 내용을 모두 선택합니다.'
PageBreak

# 화상 키보드
UnitTitle '03-1' '키 위치를 화면에서 익히기' '화상 키보드에서 Ctrl, Shift, Alt, 한/영 키의 위치를 확인합니다.'
Add-Image '03_shortcuts_osk.png' 'Windows 화상 키보드의 실제 한글 배열' 165 48
SubTitle '그림에서 찾아 표시해 보기'
Check '키보드 왼쪽 아래의 Ctrl 키를 찾았다.'
Check 'Ctrl 오른쪽 근처의 Windows 키와 Alt 키를 찾았다.'
Check '양쪽의 Shift 키를 찾았다.'
Check '스페이스바 오른쪽의 한/영 키를 찾았다.'
Check '오른쪽 위쪽의 Delete 키를 찾았다.'
Tip '두 키를 동시에 누르기 어려우면 Ctrl을 먼저 누른 채 유지하고, 다른 손으로 문자 키를 누르세요.'
SubTitle 'Windows에서 자주 쓰는 조합'
Key 'Alt + Tab' '실행 중인 프로그램을 전환합니다.'
Key 'Alt + F4' '현재 창을 닫습니다.'
Key 'Win + E' '파일 탐색기를 엽니다.'
Key 'Win + Shift + S' '화면의 일부를 캡처합니다.'
PageBreak

# 단축키 실습
UnitTitle '03-2' '복사·붙여넣기 실습' '짧은 문장을 선택하고 단축키로 복제해 봅니다.'
Para '연습 문장:  컴퓨터는 정확한 명령을 따라 움직입니다.' 12 $true $C.Navy 1 420 $C.Sky
Step '1' '연습 문장을 마우스로 끌어 파란색으로 선택합니다.'
Step '2' 'Ctrl+C를 눌러 선택한 문장을 복사합니다.'
Step '3' '문장 끝을 클릭하고 Enter를 눌러 다음 줄로 갑니다.'
Step '4' 'Ctrl+V를 두 번 눌러 같은 문장을 두 줄 붙여 넣습니다.'
Step '5' 'Ctrl+Z를 한 번 눌러 마지막 붙여넣기를 취소합니다.'
Step '6' 'Ctrl+S를 눌러 파일 이름을 “단축키연습”으로 저장합니다.'
Practice '완료 확인' @('복사와 붙여넣기의 차이를 설명할 수 있다.','실행 취소 후 문장이 한 줄 줄어든 것을 확인했다.','저장 위치와 파일 이름을 직접 확인했다.')
PageBreak

# 마우스 개요
UnitTitle '04' '마우스 조작하기' '포인터를 정확하게 이동하고 클릭·더블클릭·드래그를 구분합니다.'
Goal @('왼쪽 버튼, 오른쪽 버튼, 휠의 역할을 말할 수 있다.','클릭과 더블클릭을 상황에 맞게 사용할 수 있다.','드래그로 범위를 선택하거나 대상을 이동할 수 있다.')
Para '[마우스 왼쪽 버튼·오른쪽 버튼·휠의 위치를 표시한 실물 사진 첨부]' 11 $true $C.Red 3 420 $C.Yellow
SubTitle '마우스를 잡는 기본 자세'
Bullet '검지는 왼쪽 버튼, 중지는 오른쪽 버튼 위에 가볍게 둡니다.'
Bullet '손목에 힘을 빼고 팔 전체를 조금씩 움직입니다.'
Bullet '포인터를 목표 위에 올린 다음 손을 멈추고 클릭합니다.'
Bullet '마우스 패드나 평평한 책상 위에서 사용합니다.'
Caution '마우스를 들어 올린 상태로 클릭하면 포인터가 흔들립니다. 바닥에 내려놓고 클릭하세요.'
PageBreak

# 마우스 동작 종류
UnitTitle '04-1' '클릭의 종류와 쓰임' '버튼을 누르는 횟수와 속도에 따라 결과가 달라집니다.'
Key '가리키기' '버튼을 누르지 않고 포인터만 대상 위로 옮깁니다.'
Key '한 번 클릭' '대상을 선택하거나 버튼을 누릅니다.'
Key '더블클릭' '왼쪽 버튼을 빠르게 두 번 눌러 파일·폴더를 엽니다.'
Key '오른쪽 클릭' '대상과 관련된 바로 가기 메뉴를 엽니다.'
Key '드래그' '왼쪽 버튼을 누른 채 움직여 선택하거나 이동합니다.'
Key '휠 굴리기' '긴 문서나 웹 페이지를 위아래로 이동합니다.'
SubTitle '더블클릭이 잘 안 될 때'
Bullet '첫 번째 클릭 뒤에 오래 쉬지 말고 바로 두 번째 클릭을 합니다.'
Bullet '두 번 누르는 동안 마우스를 옆으로 움직이지 않습니다.'
Bullet '파일을 열 때 더블클릭 대신 선택 후 Enter를 눌러도 됩니다.'
Tip '웹 페이지의 링크와 프로그램의 버튼은 대부분 한 번만 클릭합니다.'
PageBreak

# 탐색기 실제 화면
UnitTitle '04-2' '파일 탐색기에서 마우스 사용' '실제 파일 목록에서 선택과 실행 위치를 확인합니다.'
Add-Image '04_mouse_explorer.png' '파일과 폴더가 보이는 Windows 파일 탐색기 실제 화면' 165 88
SubTitle '이 화면에서 할 수 있는 조작'
Bullet '폴더를 한 번 클릭하면 선택되고, 더블클릭하면 폴더 안으로 들어갑니다.'
Bullet '파일 이름을 오른쪽 클릭하면 열기, 복사, 삭제, 속성 등의 메뉴가 나타납니다.'
Bullet '오른쪽 스크롤 막대나 마우스 휠로 긴 목록을 이동합니다.'
Bullet '열의 경계를 드래그하면 이름·날짜·유형 열의 너비를 바꿀 수 있습니다.'
PageBreak

# 드래그 실습
UnitTitle '04-3' '선택과 드래그 실습' '실수로 파일을 이동하지 않도록 안전한 순서로 연습합니다.'
Practice '파일 탐색기 조작' @(
    '바탕 화면의 빈 곳을 한 번 클릭했다.',
    '파일 하나를 한 번 클릭해 선택 표시를 확인했다.',
    'Ctrl을 누른 채 다른 파일을 클릭해 여러 개를 선택했다.',
    '빈 곳을 클릭해 선택을 해제했다.',
    '폴더를 더블클릭해 열고, 뒤로 버튼으로 돌아왔다.',
    '파일을 오른쪽 클릭한 뒤 Esc를 눌러 메뉴를 닫았다.'
)
Caution '중요한 파일을 드래그하면 원하지 않는 폴더로 이동될 수 있습니다. 처음에는 연습용 파일로만 실습하세요.'
SubTitle '문서에서 범위 선택하기'
Step '1' '선택할 글자의 맨 앞에 포인터를 둡니다.'
Step '2' '왼쪽 버튼을 누른 채 마지막 글자까지 천천히 끕니다.'
Step '3' '선택 부분이 파란색으로 바뀌면 버튼에서 손을 뗍니다.'
Step '4' '선택을 취소하려면 문서의 빈 곳을 한 번 클릭합니다.'
PageBreak

# 종합 실습
UnitTitle '마무리' '키보드와 마우스 종합 실습' '새 문서를 만들고 입력·수정·선택·저장을 한 번에 연습합니다.'
Para '  완성할 안내문  ' 12 $true $C.White 1 260 $C.Blue
Para '컴퓨터 기초 수업 안내' 15 $true $C.Navy 3 240
Para '일시: 2026년 8월 14일(금) 오전 10:00' 11 $false $C.Black 1 190
Para '장소: 컴퓨터 교육실' 11 $false $C.Black 1 190
Para '준비물: 필기도구, USB 메모리' 11 $false $C.Black 1 190
Para '문의: class@example.com' 11 $false $C.Black 1 380
Step '1' '한글을 실행하고 새 문서를 엽니다.'
Step '2' '위 안내문을 한글·영문·숫자·특수문자까지 그대로 입력합니다.'
Step '3' '제목을 마우스로 드래그해 선택하고 굵게 표시합니다.'
Step '4' '준비물 줄을 Ctrl+C, Ctrl+V로 한 번 복제합니다.'
Step '5' '복제된 줄을 Ctrl+Z로 취소합니다.'
Step '6' 'Ctrl+S를 눌러 “컴퓨터기초_종합실습”으로 저장합니다.'
PageBreak

# 문제 해결
UnitTitle '마무리-1' '자주 생기는 문제 해결' '당황하지 말고 현재 상태부터 확인합니다.'
Key '영어만 입력돼요' '한/영 키를 한 번 누르고 화면의 A/가 표시를 확인합니다.'
Key '숫자 키패드가 안 돼요' 'Num Lock 키를 한 번 눌러 상태를 바꿉니다.'
Key '글자가 계속 지워져요' 'Backspace나 Delete 키가 눌려 있지 않은지 확인합니다.'
Key '창이 갑자기 닫혔어요' 'Alt+F4를 눌렀는지 확인하고, 저장된 파일을 다시 엽니다.'
Key '더블클릭이 안 돼요' '간격을 줄이고 마우스가 움직이지 않게 두 번 누릅니다.'
Key '화면이 너무 빨리 움직여요' '휠을 조금씩 굴리고 스크롤 막대를 천천히 드래그합니다.'
Key '붙여넣기가 안 돼요' '먼저 내용을 선택하고 Ctrl+C가 실행됐는지 확인합니다.'
Caution '오류 메시지가 나오면 바로 닫지 말고 문구를 읽거나 화면을 캡처해 도움을 요청하세요.'
SubTitle '도움을 요청할 때 말할 내용'
Check '어떤 프로그램에서 문제가 생겼는가?'
Check '어떤 키나 버튼을 눌렀는가?'
Check '화면에 어떤 글자나 메시지가 보이는가?'
Check '같은 동작을 다시 하면 문제가 반복되는가?'
PageBreak

# 확인 문제
UnitTitle '마무리-2' '확인 문제와 자기 점검' '답을 가리지 말고 먼저 스스로 생각해 보세요.'
Para '1. 한글과 영어 입력 상태를 바꾸는 키는 무엇인가요?' 11 $true $C.Navy 1 200
Blank '답'
Para '2. 커서 왼쪽 글자를 지우는 키는 무엇인가요?' 11 $true $C.Navy 1 200
Blank '답'
Para '3. 선택한 내용을 복사하고 붙여 넣는 단축키를 각각 쓰세요.' 11 $true $C.Navy 1 200
Blank '답'
Para '4. 파일을 열 때 주로 사용하는 마우스 동작은 무엇인가요?' 11 $true $C.Navy 1 200
Blank '답'
Para '5. 마우스 오른쪽 버튼을 누르면 무엇이 나타나나요?' 11 $true $C.Navy 1 200
Blank '답'
SubTitle '자기 점검'
Check '한글과 영어를 바꿔 입력할 수 있다.'
Check '특수문자를 키보드와 문자표에서 찾을 수 있다.'
Check '복사·붙여넣기·저장 단축키를 사용할 수 있다.'
Check '클릭·더블클릭·오른쪽 클릭·드래그를 구분할 수 있다.'
PageBreak

# 정답 및 마무리
UnitTitle '정답' '확인 문제 정답' '틀린 문제는 해당 단원으로 돌아가 한 번 더 직접 해보세요.'
Key '1번' '한/영 키'
Key '2번' 'Backspace 키'
Key '3번' '복사 Ctrl+C, 붙여넣기 Ctrl+V'
Key '4번' '왼쪽 버튼 더블클릭'
Key '5번' '바로 가기 메뉴(상황에 맞는 명령 목록)'
Para '' 10 $false $C.Black 1 500
Para '수고하셨습니다!' 21 $true $C.Green 3 360
Para '정확하게 입력하고, 천천히 포인터를 움직이는 습관이 컴퓨터 활용의 기초입니다.' 12 $false $C.Navy 3 600 $C.Mint
Blank '오늘 가장 자신 있게 할 수 있게 된 것'
Blank '다음 시간에 더 연습하고 싶은 것'
Small '화면 이미지: Windows 메모장, 문자표, 화상 키보드, 파일 탐색기를 2026-08-14에 직접 캡처했습니다. 상표와 프로그램 화면의 권리는 각 권리자에게 있습니다.'

# 쪽 번호와 저장
$hwp.HAction.GetDefault('PageNumPos', $hwp.HParameterSet.HPageNumPos.HSet) | Out-Null
$hwp.HParameterSet.HPageNumPos.DrawPos = 5
$hwp.HAction.Execute('PageNumPos', $hwp.HParameterSet.HPageNumPos.HSet) | Out-Null

$hwp.SetTitleName('컴퓨터 기초 - 키보드와 마우스') | Out-Null
$saved = $hwp.SaveAs($outFile, 'HWPX', '')
if (-not $saved) { throw 'HWPX 저장에 실패했습니다.' }
try { $hwp.SaveAs($pdfFile, 'PDF', '') | Out-Null } catch { }
$hwp.Quit()

Get-Item -LiteralPath $outFile
if (Test-Path -LiteralPath $pdfFile) { Get-Item -LiteralPath $pdfFile }
