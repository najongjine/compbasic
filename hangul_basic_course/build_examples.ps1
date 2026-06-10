$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Skeleton = "C:\Users\najon\AppData\Local\Programs\Python\Python310\Lib\site-packages\hwpx\data\Skeleton.hwpx"
$TempRoot = Join-Path $Root "_hwpx_work"
$OutRoot = Join-Path $Root "examples"

function XmlEscape([string]$Text) {
    return [System.Security.SecurityElement]::Escape($Text)
}

function New-Para([string]$Text, [int]$CharPr = 0) {
    $safe = XmlEscape $Text
    return @"
  <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
    <hp:run charPrIDRef="$CharPr">
      <hp:t>$safe</hp:t>
    </hp:run>
    <hp:linesegarray>
      <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="42520" flags="393216"/>
    </hp:linesegarray>
  </hp:p>
"@
}

function New-Section([string[]]$Lines) {
    $ns = 'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
    $paras = @()
    $paras += @"
  <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
    <hp:run charPrIDRef="0">
      <hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0">
        <hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>
        <hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>
        <hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>
        <hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>
        <hp:pagePr landscape="WIDELY" width="59528" height="84186" gutterType="LEFT_ONLY">
          <hp:margin header="4252" footer="4252" gutter="0" left="8504" right="8504" top="5668" bottom="4252"/>
        </hp:pagePr>
        <hp:footNotePr>
          <hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>
          <hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/>
          <hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/>
          <hp:numbering type="CONTINUOUS" newNum="1"/>
          <hp:placement place="EACH_COLUMN" beneathText="0"/>
        </hp:footNotePr>
        <hp:endNotePr>
          <hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>
          <hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/>
          <hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/>
          <hp:numbering type="CONTINUOUS" newNum="1"/>
          <hp:placement place="END_OF_DOCUMENT" beneathText="0"/>
        </hp:endNotePr>
        <hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">
          <hp:offset left="1417" right="1417" top="1417" bottom="1417"/>
        </hp:pageBorderFill>
      </hp:secPr>
      <hp:ctrl>
        <hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/>
      </hp:ctrl>
    </hp:run>
    <hp:linesegarray>
      <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="42520" flags="393216"/>
    </hp:linesegarray>
  </hp:p>
"@
    foreach ($line in $Lines) {
        if ($line.Trim().Length -eq 0) {
            $paras += New-Para " "
        } elseif ($line.StartsWith("# ")) {
            $paras += New-Para ($line.Substring(2)) 6
        } elseif ($line.StartsWith("## ")) {
            $paras += New-Para ($line.Substring(3)) 5
        } else {
            $paras += New-Para $line 0
        }
    }
    return "<?xml version='1.0' encoding='UTF-8'?>`n<hs:sec $ns>`n$($paras -join "`n")`n</hs:sec>`n"
}

$docs = @(
    @{
        Name = "01_메뉴_저장_연습.hwpx"
        Preview = "나의 첫 한글 문서"
        Lines = @(
            "# 나의 첫 한글 문서",
            "이름: 홍길동",
            "작성일: 2026년 6월 10일",
            "",
            "## 오늘 배운 메뉴",
            "파일: 새 문서, 저장, 불러오기, 인쇄",
            "편집: 복사, 붙이기, 찾기, 바꾸기",
            "입력: 표, 그림, 특수문자",
            "서식: 글자 모양, 문단 모양, 개요 번호",
            "도구: 글자판과 환경 설정",
            "",
            "연습: 이 문서를 내 이름으로 저장한 뒤 다시 불러옵니다."
        )
    },
    @{
        Name = "02_글자판_입력_연습.hwpx"
        Preview = "기본 입력 연습"
        Lines = @(
            "# 기본 입력 연습",
            "이름: 홍길동",
            "사는 동네: 중앙동",
            "좋아하는 음식: 김치찌개, 비빔밥, 잔치국수",
            "",
            "## 한글과 영어 섞어 쓰기",
            "오늘은 Hancom Office 한글 프로그램을 연습합니다.",
            "한/영 키를 눌러 한글과 영어를 바꿔 봅니다.",
            "",
            "확인: 원하지 않는 자동 변경이 생기면 도구 > 글자판 > 글자판 자동 변경을 끕니다."
        )
    },
    @{
        Name = "03_특수문자_한자_안내문.hwpx"
        Preview = "마을 모임 안내"
        Lines = @(
            "# 마을 모임 안내",
            "① 날짜: 2026년 6월 20일 토요일",
            "② 장소: 마을회관 2층",
            "③ 준비물: 개인 컵, 필기도구",
            "☎ 문의: 010-1234-5678",
            "※ 참석이 어려우면 미리 연락 바랍니다.",
            "",
            "한자 연습: 한국 -> 韓國",
            "",
            "연습: Ctrl + F10 문자표에서 동그라미 숫자와 전화기 표시를 넣어 봅니다."
        )
    },
    @{
        Name = "04_글자모양_행사안내문.hwpx"
        Preview = "건강 걷기 모임 안내"
        Lines = @(
            "# 건강 걷기 모임 안내",
            "일시: 2026년 6월 25일 오전 10시",
            "장소: 중앙공원 정문",
            "",
            "## 안내",
            "가벼운 운동화와 물을 준비해 주세요.",
            "비가 오면 일정이 다음 주로 미뤄집니다.",
            "문의: 010-1234-5678",
            "",
            "모양 목표: 제목은 크게/굵게/가운데, 준비물 문장은 밑줄, 문의 전화는 파란색으로 바꿉니다."
        )
    },
    @{
        Name = "05_개요번호_회의순서.hwpx"
        Preview = "동호회 회의 순서"
        Lines = @(
            "# 동호회 회의 순서",
            "1. 인사",
            "2. 지난 활동 확인",
            "   2.1) 사진 정리",
            "   2.2) 회비 사용 내역 확인",
            "3. 다음 일정 정하기",
            "   3.1) 날짜",
            "   3.2) 장소",
            "   3.3) 준비물",
            "      3.3.1) 개인 준비물",
            "      3.3.2) 공통 준비물",
            "",
            "연습: Tab과 Shift+Tab으로 번호 단계를 바꿔 봅니다."
        )
    },
    @{
        Name = "06_이미지배치_초대장.hwpx"
        Preview = "우리 동네 작은 음악회"
        Lines = @(
            "# 우리 동네 작은 음악회",
            "초대합니다.",
            "일시: 2026년 7월 3일 오후 3시",
            "장소: 주민센터 강당",
            "",
            "[그림 자리 1] 글자처럼 취급으로 넣기",
            "[그림 자리 2] 어울림으로 넣기",
            "[그림 자리 3] 글 뒤로 넣기",
            "",
            "연습: 입력 > 그림으로 같은 사진을 넣고 배치 방식을 각각 바꿉니다."
        )
    },
    @{
        Name = "07_표기본_연락망.hwpx"
        Preview = "우리 반 연락망"
        Lines = @(
            "# 우리 반 연락망",
            "번호 | 이름 | 전화번호 | 비고",
            "1 | 홍길동 | 010-1111-2222 | 반장",
            "2 | 김영희 | 010-3333-4444 | ",
            "3 | 이민수 | 010-5555-6666 | ",
            "",
            "모양 목표: 4줄 4칸 표로 만들고 첫 줄은 배경색을 넣습니다.",
            "연습: 셀 너비와 높이를 같게 맞추고, L 키로 테두리를 바꿉니다."
        )
    },
    @{
        Name = "08_표응용_일정표.hwpx"
        Preview = "6월 교육 일정표"
        Lines = @(
            "# 6월 교육 일정표",
            "용지 목표: F7을 눌러 가로 용지로 설정합니다.",
            "",
            "교육명 | 월 | 화 | 수 | 목",
            "컴퓨터 기초 | 마우스 | 키보드 | 한글 | 복습",
            "한글 문서 | 저장 | 글자 모양 | 표 만들기 | 그림 넣기",
            "마무리 | 찾기 | 바꾸기 | 인쇄 | 발표",
            "",
            "모양 목표: 첫 줄 셀을 합쳐 큰 제목을 만들고 요일 칸에 색을 넣습니다.",
            "가능하면 표 안에 작은 그림이나 도장 이미지를 넣습니다."
        )
    },
    @{
        Name = "09_찾기바꾸기_점검문서.hwpx"
        Preview = "찾기와 바꾸기 점검 문서"
        Lines = @(
            "# 수업 안내문",
            "홍길동 님께 안내드립니다.",
            "홍길동 님의 다음 수업은 6월 20일입니다.",
            "장소는 마을회관입니다.",
            "마을회관에 오실 때 필기도구를 준비해 주세요.",
            "",
            "바꾸기 연습:",
            "홍길동 -> 내 이름",
            "6월 20일 -> 실제 수업 날짜",
            "마을회관 -> 원하는 장소",
            "",
            "확인: Ctrl+F로 찾고 Ctrl+H로 바꿉니다."
        )
    }
)

if (Test-Path -LiteralPath $TempRoot) {
    Remove-Item -LiteralPath $TempRoot -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $TempRoot, $OutRoot | Out-Null

foreach ($doc in $docs) {
    $work = Join-Path $TempRoot ([IO.Path]::GetFileNameWithoutExtension($doc.Name))
    hwpx-unpack.exe --force --pretty-xml $Skeleton $work | Out-Null
    $section = New-Section $doc.Lines
    Set-Content -LiteralPath (Join-Path $work "Contents\section0.xml") -Value $section -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $work "Preview\PrvText.txt") -Value $doc.Preview -Encoding UTF8
    $out = Join-Path $OutRoot $doc.Name
    if (Test-Path -LiteralPath $out) {
        Remove-Item -LiteralPath $out -Force
    }
    hwpx-pack.exe --force $work $out | Out-Null
}

Remove-Item -LiteralPath $TempRoot -Recurse -Force
Write-Host "Generated HWPX examples in $OutRoot"

$hwpComBuilder = Join-Path $Root "rebuild_06_08_with_hwp.ps1"
if (Test-Path -LiteralPath $hwpComBuilder) {
    & $hwpComBuilder
}



