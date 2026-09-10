$ErrorActionPreference='Stop'
$root='D:/compbasic/한글기초'
$asset=Join-Path $root 'assets/토끼고기요리'
$out=Join-Path $root '예제/토끼고기_간장조림_만들기.hwpx'
$h=New-Object -ComObject HWPFrame.HwpObject
$h.XHwpWindows.Item(0).Visible=$false
$h.Clear(1)|Out-Null
function Font([double]$size=11,[bool]$bold=$false){
 $h.HAction.GetDefault('CharShape',$h.HParameterSet.HCharShape.HSet)|Out-Null
 $c=$h.HParameterSet.HCharShape
 $c.Height=[int]($size*100); $c.Bold=[int]$bold
 $c.FaceNameHangul='함초롬돋움'; $c.FaceNameLatin='함초롬돋움'; $c.FontTypeHangul=1; $c.FontTypeLatin=1; $c.TextColor=0
 $h.HAction.Execute('CharShape',$c.HSet)|Out-Null
}
function Para([int]$spacing=145,[int]$after=400){
 $h.HAction.GetDefault('ParagraphShape',$h.HParameterSet.HParaShape.HSet)|Out-Null
 $p=$h.HParameterSet.HParaShape; $p.LineSpacing=$spacing; $p.PrevSpacing=0; $p.NextSpacing=$after
 $h.HAction.Execute('ParagraphShape',$p.HSet)|Out-Null
}
function Text([string]$t){
 $h.HAction.GetDefault('InsertText',$h.HParameterSet.HInsertText.HSet)|Out-Null
 $h.HParameterSet.HInsertText.Text=$t
 $h.HAction.Execute('InsertText',$h.HParameterSet.HInsertText.HSet)|Out-Null
}
function Line([string]$t,[double]$size=11,[bool]$bold=$false){ Font $size $bold; Text $t; $h.Run('BreakPara')|Out-Null }
function Heading([string]$t){ Para 135 450; Line $t 14 $true; Para 145 350 }
function Pic([string]$path,[double]$w,[double]$hh){
 $ctrl=$h.InsertPicture($path,$true,1,$false,$false,0,$w,$hh)
 if($null -eq $ctrl){throw 'Picture insertion failed'}
 $set=$ctrl.Properties; $set.SetItem('TreatAsChar',1); $ctrl.Properties=$set
 $h.Run('MoveLineEnd')|Out-Null; $h.Run('BreakPara')|Out-Null
}
function Table($rows){
 $a=$h.CreateAction('TableCreate'); $s=$a.CreateSet(); $a.GetDefault($s)|Out-Null
 $s.SetItem('Rows',$rows.Count); $s.SetItem('Cols',3); $s.SetItem('WidthType',2); $s.SetItem('HeightType',1)
 $s.SetItem('WidthValue',$h.MiliToHwpUnit(170)); $s.SetItem('HeightValue',$h.MiliToHwpUnit(65))
 $cw=$s.CreateItemArray('ColWidth',3)
  $cw.SetItem(0,$h.MiliToHwpUnit(44));  $cw.SetItem(1,$h.MiliToHwpUnit(37));  $cw.SetItem(2,$h.MiliToHwpUnit(89))
 $a.Execute($s)|Out-Null
 for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt 3;$c++){
  Para 120 0; Font 10 ($r -eq 0); Text $rows[$r][$c]
  if(!($r -eq $rows.Count-1 -and $c -eq 2)){$h.Run('TableRightCell')|Out-Null}
 }}
 $h.Run('CloseEx')|Out-Null; $h.Run('MoveDocEnd')|Out-Null; $h.Run('BreakPara')|Out-Null; Para 145 350
}
try{
 $h.HAction.GetDefault('PageSetup',$h.HParameterSet.HSecDef.HSet)|Out-Null
 $pg=$h.HParameterSet.HSecDef.PageDef
 $pg.PaperWidth=$h.MiliToHwpUnit(210); $pg.PaperHeight=$h.MiliToHwpUnit(297)
 $pg.LeftMargin=$h.MiliToHwpUnit(20); $pg.RightMargin=$h.MiliToHwpUnit(20)
 $pg.TopMargin=$h.MiliToHwpUnit(16); $pg.BottomMargin=$h.MiliToHwpUnit(16)
 $pg.HeaderLen=$h.MiliToHwpUnit(4); $pg.FooterLen=$h.MiliToHwpUnit(4)
 $h.HAction.Execute('PageSetup',$h.HParameterSet.HSecDef.HSet)|Out-Null
 Para 130 450
 Line '토끼고기 간장조림 만들기' 24 $true
 Line '감자와 당근을 넣고 약한 불에서 부드럽게 익히는 요리' 11
 Line '3~4인분  |  약 90~110분  |  냉동 고기 해동 시간 별도' 10
 Pic (Join-Path $asset '01_준비와굽기.png') 170 56.7
 Para 120 450; Line '재료 준비 → 겉면 굽기 → 채소와 함께 조리기' 9
 Heading '재료와 분량'
 Table @(
 @('재료','분량','준비 방법'),
 @('토끼고기','뼈 포함 1 kg','식용으로 판매되는 토막 고기'),
 @('감자 · 당근','감자 2개 · 당근 1개','감자 약 300 g, 당근 약 150 g'),
 @('양파 · 대파','양파 1개 · 대파 1대','양파는 큼직하게, 대파는 어슷하게'),
 @('진간장 · 설탕','4큰술 · 1큰술','1큰술은 15 mL 계량스푼 기준'),
 @('다진 마늘 · 생강','1큰술 · 얇은 생강 3쪽','생강은 마지막에 건져 내기'),
 @('물','600 mL + 보충용','고기가 절반 정도 잠기도록 조절'),
 @('식용유 · 후추','1큰술 · 약간','겉면 굽기와 마무리에 사용')
 )
 Heading '1  고기와 채소 준비하기'
 Line '냉동 고기는 냉장실에서 완전히 해동한다. 고기의 물기를 키친타월로 닦고, 감자와 당근은 3~4 cm 크기로 썬다. 생고기에 쓴 도마와 칼은 씻은 뒤 채소에 사용한다.' 10.5
 Heading '2  양념 섞고 고기 굽기'
 Line '간장·설탕·마늘을 섞는다. 두꺼운 냄비에 식용유를 두르고 중강불로 달군 뒤, 고기를 겹치지 않게 나누어 넣는다. 앞뒤로 각각 2~3분씩 갈색이 나도록 굽는다. 이 단계에서는 속까지 익히지 않는다.' 10.5
 $h.Run('BreakPage')|Out-Null
 Para 130 450
 Line '천천히 익혀 완성하기' 22 $true
 Line '고기는 먼저 충분히 익히고, 감자와 당근은 나중에 넣습니다.' 11
 Pic (Join-Path $asset '02_조림과완성.png') 170 85
 Para 120 450; Line '왼쪽 고기 먼저 익히기  |  오른쪽 채소를 넣고 완성한 상태' 9
 Heading '3  국물을 넣고 약불로 익히기'
 Line '구운 고기를 냄비에 모두 담고 양념·생강·물 600 mL를 넣는다. 중불에서 끓기 시작하면 약불로 낮추고 뚜껑을 덮어 40분간 익힌다. 15분마다 살펴 바닥이 마르면 뜨거운 물을 50~100 mL씩 더한다.' 10.5
 Heading '4  채소 넣고 더 익히기'
 Line '감자·당근·양파를 넣고 뚜껑을 덮어 약불에서 20~30분 더 익힌다. 감자에 젓가락이 쉽게 들어가고 고기가 포크로 부드럽게 갈라지는지 확인한다. 질기면 물을 조금 더 넣고 10분씩 연장한다.' 10.5
 Heading '5  익힘 확인하고 국물 졸이기'
 Line '뼈에 닿지 않게 가장 두꺼운 살의 중심 온도를 재어 71.1 ℃ 이상인지 확인한다. 대파와 후추를 넣고 뚜껑을 연 채 중약불에서 5~10분 졸인다. 짜면 물을 조금 넣고, 싱거우면 간장을 1작은술씩 추가한다.' 10.5
 Heading '6  담아내고 보관하기'
 Line '생강을 건져 내고 고기와 채소를 그릇에 담아 국물을 끼얹는다. 먹을 때 작은 뼛조각에 주의한다. 남은 음식은 조리 후 2시간 안에 냉장하고 3~4일 안에 먹는다.' 10.5
 Para 120 200
 Line '조리 시간은 고기의 크기와 질김에 따라 달라집니다. 시간이나 색만으로 익힘을 판단하지 마세요.' 9
 Line '그림은 조리 과정을 설명하기 위한 AI 생성 이미지입니다. 분량은 본문 재료표를 기준으로 합니다.' 8
 Line '안전 기준 참고  USDA FSIS  Rabbit From Farm to Table' 8
 Line 'https://www.fsis.usda.gov/food-safety/safe-food-handling-and-preparation/meat-fish/rabbit-farm-table' 7
 if(!$h.SaveAs($out,'HWPX','')){throw 'HWPX save failed'}
 $pdf=Join-Path $asset '_검수용.pdf'
 if(!$h.SaveAs($pdf,'PDF','')){throw 'PDF export failed'}
 Write-Output "Saved: $out"
 Write-Output "Pages: $($h.PageCount)"
} finally { $h.Quit()|Out-Null }
