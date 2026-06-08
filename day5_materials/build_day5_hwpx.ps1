$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Skeleton = "C:\Users\najon\AppData\Local\Programs\Python\Python310\Lib\site-packages\hwpx\data\Skeleton.hwpx"
$TempRoot = Join-Path $Root "_hwpx_work"
$OutRoot = $Root

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
        <hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">
          <hp:offset left="1417" right="1417" top="1417" bottom="1417"/>
        </hp:pageBorderFill>
        <hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">
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
        Name = "5일차_결과물1_주간업무보고서.hwpx"
        Preview = "주간 업무 보고서"
        Lines = @(
            "# 주간 업무 보고서",
            "작성자: 홍길동    부서: 운영팀    작성일: 2026.06.12",
            "보고 기간: 2026.06.08 - 2026.06.12",
            "",
            "## 1. 이번 주 주요 업무",
            "- 신규 고객 문의 24건 확인 및 처리",
            "- 매장 재고 현황 점검",
            "- 직원 교육 일정 안내문 작성",
            "",
            "## 2. 진행 현황",
            "업무명 | 담당자 | 상태 | 마감일",
            "고객 문의 정리 | 홍길동 | 완료 | 6/10",
            "재고표 업데이트 | 김영희 | 진행 중 | 6/13",
            "교육 안내문 배포 | 이민수 | 완료 | 6/12",
            "",
            "## 3. 문제점 및 요청 사항",
            "- 재고 수량 입력 기준이 담당자마다 달라 통일이 필요함",
            "- 고객 문의 유형을 엑셀에서 분류하면 다음 보고가 쉬워짐",
            "",
            "## 4. 다음 주 계획",
            "- 고객 문의 유형별 통계 정리",
            "- 재고 부족 상품 목록 작성",
            "- 교육 참석자 명단 확정"
        )
    },
    @{
        Name = "5일차_결과물2_교육참석보고서.hwpx"
        Preview = "교육 참석 보고서"
        Lines = @(
            "# 교육 참석 보고서",
            "작성자: 홍길동    부서: 총무팀    작성일: 2026.06.12",
            "교육명: AI 활용 기초 교육",
            "",
            "## 1. 교육 개요",
            "일시: 2026.06.11 14:00-17:00",
            "장소: 2층 교육실",
            "참석 대상: 사무직 직원 12명",
            "",
            "## 2. 교육 내용",
            "- GPT로 문장 초안 작성하기",
            "- 이미지 생성 요청문 작성하기",
            "- 한글 보고서 양식 만들기",
            "- 엑셀 자료 정리와 기본 함수 이해하기",
            "",
            "## 3. 배운 점",
            "- 문서를 처음부터 쓰기보다 GPT로 초안을 만들면 시간이 줄어듦",
            "- 보고서는 제목, 요약, 본문, 요청 사항 순서로 쓰면 읽기 쉬움",
            "- 엑셀 자료는 정렬과 필터만 알아도 확인 속도가 빨라짐",
            "",
            "## 4. 업무 적용 계획",
            "- 주간 보고서 초안 작성에 GPT 활용",
            "- 교육 참석 명단을 엑셀로 정리",
            "- 반복 문서는 한글 양식으로 저장해 재사용"
        )
    },
    @{
        Name = "5일차_결과물3_매장운영보고서.hwpx"
        Preview = "매장 운영 보고서"
        Lines = @(
            "# 매장 운영 보고서",
            "작성자: 홍길동    매장명: 중앙점    작성일: 2026.06.12",
            "보고 기간: 2026.06.01 - 2026.06.07",
            "",
            "## 1. 운영 요약",
            "- 방문 고객 수가 전주보다 증가함",
            "- 주말 매출 비중이 높았음",
            "- 인기 상품 일부의 재고가 부족함",
            "",
            "## 2. 주요 지표",
            "항목 | 이번 주 | 지난주 | 메모",
            "방문 고객 수 | 342명 | 301명 | 증가",
            "총 매출 | 4,250,000원 | 3,880,000원 | 증가",
            "반품 건수 | 6건 | 8건 | 감소",
            "",
            "## 3. 개선이 필요한 부분",
            "- 재고 부족 상품을 매주 금요일 오전에 확인 필요",
            "- 고객 문의 내용을 유형별로 기록하면 응대 개선에 도움 됨",
            "",
            "## 4. 다음 주 실행 계획",
            "- 인기 상품 3종 추가 발주",
            "- 고객 문의 목록 엑셀 정리",
            "- 매장 안내문 문구 수정"
        )
    }
)

if (Test-Path -LiteralPath $TempRoot) {
    Remove-Item -LiteralPath $TempRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $TempRoot | Out-Null

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

Write-Host "Generated HWPX files in $OutRoot"

