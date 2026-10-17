param([string]$WorkbookPath,[switch]$SkipBoundaries,[switch]$SkipTables)
# Native Excel refresh. No macros, no external data, no new financing plug.
$ErrorActionPreference='Stop'
if(-not $WorkbookPath){$WorkbookPath=Join-Path $PSScriptRoot '产业园投资测算_V3_决策测算.xlsx'}
$manifestPath=Join-Path $PSScriptRoot '_过程文件/manifest.json'
$manifest=Get-Content -LiteralPath $manifestPath -Raw -Encoding utf8 | ConvertFrom-Json
$excel=$null;$book=$null
$saved=@{}
function Recalc {
 # Value2/Formula changes recalculate synchronously in AutomaticExceptTables.
 # Application.Calculate would explicitly recalculate every sensitivity table
 # for every root-search point; wait only if ordinary calculation is active.
 while($script:excel.CalculationState -eq 1){Start-Sleep -Milliseconds 10}
}
function Get-Result([string]$Address,[double]$Value,[bool]$AllConditions){
 $script:testSheet.Range($Address).Value2=$Value
 Recalc
 if($AllConditions){
  $ok=($script:metrics.Range('B50').Value2 -eq 1)
  if($script:inputs.Range('B82').Value2 -eq 1){$ok=$ok -and ($script:metrics.Range('E50').Value2 -eq 1)}
  return [bool]$ok
 }
 $v=$script:metrics.Range('B8').Value2
 return [bool](($v -is [double]) -and $v -ge 0)
}
function Find-Boundary([string]$Address,[double]$Upper,[bool]$Maximum,[bool]$AllConditions,[bool]$Integer=$false){
 $script:testSheet=if($Address -eq 'B41'){$script:inputs}else{$script:sens}
 $oldFormula=$script:testSheet.Range($Address).Formula
 try{
  if($Upper -le 0){return @{value='搜索上限无效';note='请填写正的搜索上限'}}
  $points=[System.Collections.Generic.List[object]]::new()
  $steps=if($Integer){[int]$Upper}else{16}
  for($i=0;$i -le $steps;$i++){
   $x=if($Integer){[double]$i}else{$Upper*$i/$steps}
   if($Address -eq 'B41'){
    $lower=[Math]::Max($script:inputs.Range('B38').Value2,[Math]::Max($script:inputs.Range('B39').Value2,$script:inputs.Range('B40').Value2))
    $x=$lower+($Upper-$lower)*$i/$steps
   }
   $ok=Get-Result $Address $x $AllConditions
   $points.Add(@{x=$x;ok=$ok})
  }
  $good=@($points | Where-Object {$_.ok})
  if($good.Count -eq 0){return @{value='测试范围内未找到可行值';note='其他条件固定；检查未通过指标及搜索范围。扫描不证明范围外或网格间不存在解'}}
  $point=if($Maximum){$good[-1]}else{$good[0]}
  $idx=$points.IndexOf($point)
  if($Integer){return @{value=$point.x;note=if($Maximum -and $idx -eq $points.Count-1){'测试上限仍满足，需扩大范围'}else{'按整数季度逐一检验'}}}
  if(($Maximum -and $idx -eq $points.Count-1) -or ((-not $Maximum) -and $idx -eq 0)){
   return @{value=$point.x;note=if($Maximum){'搜索上限仍满足；实际最高值可能更高'}else{'搜索下限已满足'}}
  }
  $lo=if($Maximum){$point.x}else{$points[$idx-1].x}
  $hi=if($Maximum){$points[$idx+1].x}else{$point.x}
  for($i=0;$i -lt 25;$i++){
   $mid=($lo+$hi)/2
   $ok=Get-Result $Address $mid $AllConditions
   if($Maximum){if($ok){$lo=$mid}else{$hi=$mid}}
   else{if($ok){$hi=$mid}else{$lo=$mid}}
  }
  return @{value=if($Maximum){$lo}else{$hi};note='扫描定位可行区间后25次二分；其他条件固定，非多变量全局最优'}
 } finally {
  $restore=$script:testSheet.Range($Address)
  if($null -eq $oldFormula -or ($oldFormula -is [string] -and $oldFormula -eq '')){$restore.ClearContents()|Out-Null}
  elseif($oldFormula -is [string]){$restore.Formula=[string]$oldFormula}
  else{$restore.Value2=[double]$oldFormula}
  Recalc
 }
}
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
 $book=$excel.Workbooks.Open($WorkbookPath,0,$false)
 if($book.ReadOnly){throw '文件被占用或只读，请先在Excel关闭该文件后刷新。'}
 foreach($sheet in $book.Worksheets){$sheet.Unprotect()}
 $inputs=$book.Worksheets.Item('核心参数输入');$metrics=$book.Worksheets.Item('指标与情景');$sens=$book.Worksheets.Item('敏感性与边界')
 $excel.Calculation=2 # Automatic except data tables, for repeated boundary solves.
 Recalc
 if(-not $SkipBoundaries){
  if($metrics.Range('B49').Value2 -ne 1 -or ($inputs.Range('B82').Value2 -eq 1 -and $metrics.Range('E49').Value2 -ne 1)){throw '输入、计划或融资条件无效，请先修正后再计算边界。'}
  $solve=@(
   @{row=73;address='B6';upper=[double]$inputs.Range('B132').Value2;max=$true;all=$false},
   @{row=74;address='B6';upper=[double]$inputs.Range('B132').Value2;max=$true;all=$true},
   @{row=75;address='B7';upper=[double]$inputs.Range('B133').Value2;max=$false;all=$false},
   @{row=76;address='B7';upper=[double]$inputs.Range('B133').Value2;max=$false;all=$true},
   @{row=77;address='B8';upper=[double]$inputs.Range('B134').Value2;max=$false;all=$false},
   @{row=78;address='B8';upper=[double]$inputs.Range('B134').Value2;max=$false;all=$true},
   @{row=79;address='B41';upper=[double]$inputs.Range('B137').Value2;max=$false;all=$true},
   @{row=80;address='B9';upper=[double]$inputs.Range('B135').Value2;max=$true;all=$true},
   @{row=81;address='B10';upper=[double]$inputs.Range('B136').Value2;max=$true;all=$true;integer=$true}
  )
  foreach($s in $solve){
   $r=Find-Boundary $s.address $s.upper $s.max $s.all ([bool]$s.integer)
   $target=$sens.Range('B'+$s.row)
   if($r.value -is [string]){$target.Value2=[string]$r.value}else{$target.Value2=[double]$r.value}
   $sens.Range('D'+$s.row).Value2=[string]$r.note
   Write-Output ('边界 '+$s.row+'：'+$r.value)
  }
  $sens.Range('B79').NumberFormat='0.00%'
  Recalc
  $sens.Range('B69').Value2=$sens.Range('B68').Value2
  $sens.Range('B67').Value2=[DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
 }
 if(-not $SkipTables){
  foreach($table in $manifest.dataTables){
   $sens.Range($table.range).Table($sens.Range($table.rowInput),$sens.Range($table.columnInput)) | Out-Null
  }
 }
 $excel.Calculation=-4105
 $excel.CalculateFullRebuild()
 # Keep boundary precision readable and explanatory notes within the table.
 $sens.Range('B73:B81').NumberFormat='#,##0.00;(#,##0.00);"–"'
 $sens.Range('B79').NumberFormat='0.00%'
 $sens.Range('B81').NumberFormat='0'
 foreach($row in (73..81)+84){
  $note=$sens.Range('D'+$row+':G'+$row)
  if(-not $note.MergeCells){$note.Merge()|Out-Null}
  $note.WrapText=$true;$note.VerticalAlignment=-4160
  $sens.Rows.Item($row).RowHeight=34
 }
 # Dynamic named sources exclude all reserved years after N from Excel charts.
 $chartColumns=@{Year='A';CFADS='B';Service='C';Cash='D';Gap='E';Equity='F';Distribution='G'}
 foreach($key in $chartColumns.Keys){
  $name='V3Chart'+$key;$c=$chartColumns[$key]
  try{$book.Names.Item($name).Delete()}catch{}
  $formula="='决策看板'!`$$c`$81:INDEX('决策看板'!`$$c`$81:`$$c`$111,'核心参数输入'!`$B`$9+1)"
  $book.Names.Add($name,$formula)|Out-Null
 }
 $bookName=$book.Name.Replace("'","''")
 $dashboard=$book.Worksheets.Item('决策看板')
 $chartSeries=@(@('CFADS','Service'),@('Cash','Gap'),@('Equity','Distribution'))
 $chartLabels=@(@('经营CFADS','合同本息'),@('年末留存','未落实缺口'),@('所需股东出资','股东分配'))
 for($ci=1;$ci -le 3;$ci++){
  $chart=$dashboard.ChartObjects().Item($ci).Chart
  for($si=1;$si -le 2;$si++){
   $key=$chartSeries[$ci-1][$si-1];$label=$chartLabels[$ci-1][$si-1]
   $chart.SeriesCollection($si).Formula=('=SERIES("'+$label+'",'+"'"+$bookName+"'!V3ChartYear,"+"'"+$bookName+"'!V3Chart"+$key+','+$si+')')
  }
 }
 $errors=[System.Collections.Generic.List[object]]::new()
 foreach($sheet in $book.Worksheets){
  try{$bad=$sheet.Cells.SpecialCells(-4123,16);foreach($cell in $bad.Cells){$errors.Add(@{sheet=$sheet.Name;cell=$cell.Address($false,$false);text=$cell.Text})}}catch{}
 }
 if($errors.Count){$errors | ConvertTo-Json -Depth 3 | Write-Output;throw '发现公式错误，未保存。'}
 # Protection is applied natively because the bundled authoring API exposes no worksheet protection surface.
 foreach($sheet in $book.Worksheets){$sheet.Cells.Locked=$true}
 foreach($range in $manifest.controls){$inputs.Range($range).Locked=$false}
 foreach($range in $manifest.quarterInputs){$book.Worksheets.Item('季度计划').Range($range).Locked=$false}
 if($manifest.monthlyInputs){foreach($range in $manifest.monthlyInputs){$book.Worksheets.Item('关键月份核对').Range($range).Locked=$false}}
 $sens.Range('B6:B10').Locked=$false
 $book.Worksheets.Item('年度汇总').Rows.Item(42).Hidden=$true;$book.Worksheets.Item('年度汇总').Rows.Item(102).Hidden=$true
 $book.Worksheets.Item('指标与情景').Rows.Item(90).Hidden=$true;$book.Worksheets.Item('指标与情景').Rows.Item(200).Hidden=$true
 $sens.Rows.Item('105:115').Hidden=$true
 foreach($sheet in $book.Worksheets){
  $sheet.Protect('',$true,$true,$true,$true)
  $sheet.EnableSelection=0
 }
 $book.Worksheets.Item('决策看板').Activate()
 $book.Save()
 Write-Output ('已重算并保存：'+$WorkbookPath)
} finally {
 if($null -ne $book){$book.Close($false);[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
 if($null -ne $excel){$excel.Quit();[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
}
