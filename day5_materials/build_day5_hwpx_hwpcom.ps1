$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path

function New-Hwp {
    if ($null -eq $script:SharedHwp) {
        $script:SharedHwp = New-Object -ComObject HWPFrame.HwpObject
        try { $script:SharedHwp.XHwpWindows.Item(0).Visible = $false } catch {}
    }
    $hwp = $script:SharedHwp
    try { $hwp.Run("FileNew") | Out-Null } catch {}
    return $hwp
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

function Try-Run($hwp, [string]$command) {
    try { $hwp.Run($command) | Out-Null } catch {}
}

function Put-Title($hwp, [string]$text) {
    Try-Run $hwp "ParagraphShapeAlignCenter"
    Try-Run $hwp "CharShapeBold"
    Put-Line $hwp $text
    Try-Run $hwp "CharShapeBold"
    Try-Run $hwp "ParagraphShapeAlignLeft"
    New-Line $hwp
}

function Put-Heading($hwp, [string]$text) {
    Try-Run $hwp "CharShapeBold"
    Put-Line $hwp $text
    Try-Run $hwp "CharShapeBold"
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

function Save-Hwpx($hwp, [string]$fileName) {
    $out = Join-Path $Root $fileName
    if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
    $hwp.SaveAs($out, "HWPX", "") | Out-Null
}

function Build-WeeklyReport {
    $hwp = New-Hwp
    Put-Title $hwp "주간 업무 보고서"
    Put-Line $hwp "작성자: 홍길동    부서: 운영팀    작성일: 2026.06.12"
    Put-Line $hwp "보고 기간: 2026.06.08 - 2026.06.12"
    New-Line $hwp
    Put-Heading $hwp "1. 이번 주 주요 업무"
    Put-Line $hwp "- 신규 고객 문의 24건 확인 및 처리"
    Put-Line $hwp "- 매장 재고 현황 점검"
    Put-Line $hwp "- 직원 교육 일정 안내문 작성"
    New-Line $hwp
    Put-Heading $hwp "2. 진행 현황"
    Put-Table $hwp @(
        @("업무명", "담당자", "상태", "마감일"),
        @("고객 문의 정리", "홍길동", "완료", "6/10"),
        @("재고표 업데이트", "김영희", "진행 중", "6/13"),
        @("교육 안내문 배포", "이민수", "완료", "6/12")
    )
    Put-Heading $hwp "3. 문제점 및 요청 사항"
    Put-Line $hwp "- 재고 수량 입력 기준이 담당자마다 달라 통일이 필요함"
    Put-Line $hwp "- 고객 문의 유형을 엑셀에서 분류하면 다음 보고가 쉬워짐"
    New-Line $hwp
    Put-Heading $hwp "4. 다음 주 계획"
    Put-Line $hwp "- 고객 문의 유형별 통계 정리"
    Put-Line $hwp "- 재고 부족 상품 목록 작성"
    Put-Line $hwp "- 교육 참석자 명단 확정"
    Save-Hwpx $hwp "5일차_결과물1_주간업무보고서.hwpx"
}

function Build-EducationReport {
    $hwp = New-Hwp
    Put-Title $hwp "교육 참석 보고서"
    Put-Line $hwp "작성자: 홍길동    부서: 총무팀    작성일: 2026.06.12"
    Put-Line $hwp "교육명: AI 활용 기초 교육"
    New-Line $hwp
    Put-Heading $hwp "1. 교육 개요"
    Put-Table $hwp @(
        @("항목", "내용"),
        @("일시", "2026.06.11 14:00-17:00"),
        @("장소", "2층 교육실"),
        @("참석 대상", "사무직 직원 12명")
    )
    Put-Heading $hwp "2. 교육 내용"
    Put-Line $hwp "- GPT로 문장 초안 작성하기"
    Put-Line $hwp "- 이미지 생성 요청문 작성하기"
    Put-Line $hwp "- 한글 보고서 양식 만들기"
    Put-Line $hwp "- 엑셀 자료 정리와 기본 함수 이해하기"
    New-Line $hwp
    Put-Heading $hwp "3. 업무 적용 계획"
    Put-Line $hwp "- 주간 보고서 초안 작성에 GPT 활용"
    Put-Line $hwp "- 교육 참석 명단을 엑셀로 정리"
    Put-Line $hwp "- 반복 문서는 한글 양식으로 저장해 재사용"
    Save-Hwpx $hwp "5일차_결과물2_교육참석보고서.hwpx"
}

function Build-StoreReport {
    $hwp = New-Hwp
    Put-Title $hwp "매장 운영 보고서"
    Put-Line $hwp "작성자: 홍길동    매장명: 중앙점    작성일: 2026.06.12"
    Put-Line $hwp "보고 기간: 2026.06.01 - 2026.06.07"
    New-Line $hwp
    Put-Heading $hwp "1. 운영 요약"
    Put-Line $hwp "- 방문 고객 수가 전주보다 증가함"
    Put-Line $hwp "- 주말 매출 비중이 높았음"
    Put-Line $hwp "- 인기 상품 일부의 재고가 부족함"
    New-Line $hwp
    Put-Heading $hwp "2. 주요 지표"
    Put-Table $hwp @(
        @("항목", "이번 주", "지난주", "메모"),
        @("방문 고객 수", "342명", "301명", "증가"),
        @("총 매출", "4,250,000원", "3,880,000원", "증가"),
        @("반품 건수", "6건", "8건", "감소")
    )
    Put-Heading $hwp "3. 개선이 필요한 부분"
    Put-Line $hwp "- 재고 부족 상품을 매주 금요일 오전에 확인 필요"
    Put-Line $hwp "- 고객 문의 내용을 유형별로 기록하면 응대 개선에 도움 됨"
    New-Line $hwp
    Put-Heading $hwp "4. 다음 주 실행 계획"
    Put-Line $hwp "- 인기 상품 3종 추가 발주"
    Put-Line $hwp "- 고객 문의 목록 엑셀 정리"
    Put-Line $hwp "- 매장 안내문 문구 수정"
    Save-Hwpx $hwp "5일차_결과물3_매장운영보고서.hwpx"
}

$script:SharedHwp = $null
Build-WeeklyReport
Build-EducationReport
Build-StoreReport
if ($null -ne $script:SharedHwp) {
    $script:SharedHwp.Quit() | Out-Null
}

Write-Host "Generated HWPX files with Hanword COM."

