param([string]$Path)
$ErrorActionPreference='Stop'
if(-not $Path){$Path=Join-Path $PSScriptRoot '候选_V3.xlsx'}
$m=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'manifest.json') -Raw -Encoding utf8 | ConvertFrom-Json
$cases=@(
 @{name='base';edits=@{}},
 @{name='zero_stress';edits=@{B78=0;B79=0;B80=0;B81=0}},
 @{name='no_loan';edits=@{B57=0}},
 @{name='equity_10000';edits=@{B61=10000}},
 @{name='pure_sale';edits=@{B95=2}},
 @{name='pure_hold';edits=@{B95=3}},
 @{name='delay_4';edits=@{B107=4}},
 @{name='horizon_10';edits=@{B9=10}},
 @{name='horizon_30';edits=@{B9=30}},
 @{name='disposal';edits=@{B64=1}},
 @{name='VAT_separate';edits=@{B96=2;B118=0.09;B119=0.09;B120=0.05;B121=0.02;B122=0.01}},
 @{name='tax_lag_2';edits=@{B111=2;B112=2}},
 @{name='draw_lag_4';edits=@{B106=4}},
 @{name='finance_fee';edits=@{B113=0.01}},
 @{name='major_update';edits=@{B117=1000}},
 @{name='blank_rent';edits=@{B36=$null}},
 @{name='negative_cost';edits=@{B24=-1}},
 @{name='explicit_zero_override';edits=@{};overrideZero=$true},
 @{name='shifted_terminal';edits=@{};shiftTerminal=$true}
)
$excel=$null;$results=[System.Collections.Generic.List[object]]::new()
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
 foreach($case in $cases){
  $book=$null
  try{
   $book=$excel.Workbooks.Open($Path,0,$true)
   foreach($s in $book.Worksheets){$s.Unprotect()}
   $excel.Calculation=2
   $i=$book.Worksheets.Item('核心参数输入');$metric=$book.Worksheets.Item('指标与情景');$audit=$book.Worksheets.Item('模型校验');$cf=$book.Worksheets.Item('季度现金流')
   foreach($key in $case.edits.Keys){if($null -eq $case.edits[$key]){$i.Range($key).ClearContents()|Out-Null}else{$i.Range($key).Value2=[double]$case.edits[$key]}}
   if($case.overrideZero){$book.Worksheets.Item('季度计划').Range('E20').Value2=0.0}
   if($case.shiftTerminal){$cf.Range('CE17').Value2=500.0;$cf.Range('CF17').Value2=0.0}
   $excel.Calculate()
   $errors=[System.Collections.Generic.List[object]]::new()
   foreach($s in $book.Worksheets){try{$bad=$s.Cells.SpecialCells(-4123,16);foreach($cell in $bad.Cells){$errors.Add(@{sheet=$s.Name;cell=$cell.Address($false,$false);text=$cell.Text})}}catch{}}
   $values=@{}
   foreach($c in @('B6','B7','B8','B9','B10','B11','B12','B13','B14','B15','B16','B21','B34','B38','B39','B40','B41','B42','B43','B49','B50','B51','E6','E7','E8','E13','E14','E49')){$values[$c]=$metric.Range($c).Value2}
   $failed=[System.Collections.Generic.List[object]]::new()
   foreach($ch in $m.checks){if($audit.Range('D'+$ch.row).Value2 -ne '通过'){$failed.Add($ch.name)}}
   $results.Add(@{case=$case.name;errors=$errors.ToArray();metrics=$values;failedChecks=$failed.ToArray();inputValid=$i.Range('B150').Value2;planValid=$book.Worksheets.Item('季度计划').Range('B30').Value2;first=$book.Worksheets.Item('季度投融资').Range('B28').Value2;maturity=$book.Worksheets.Item('季度投融资').Range('B29').Value2;capitalTotal=$excel.WorksheetFunction.Sum($cf.Range('D55:DT55'));rentCash=$excel.WorksheetFunction.Sum($book.Worksheets.Item('季度经营税费').Range('D26:DT26'));saleCash=$excel.WorksheetFunction.Sum($book.Worksheets.Item('季度经营税费').Range('D12:DT12'));opCost=$excel.WorksheetFunction.Sum($book.Worksheets.Item('季度经营税费').Range('D37:DT39'));firstFormation=$book.Worksheets.Item('季度计划').Range('E11').Value2;boundaryStatus=$book.Worksheets.Item('敏感性与边界').Range('B66').Value2})
   Write-Output ($case.name+' errors='+$errors.Count+' failedChecks='+$failed.Count)
  }finally{if($null -ne $book){$book.Close($false);[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}}
 }
 $results.ToArray() | ConvertTo-Json -Depth 8 -Compress | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'scenario_results.json') -Encoding utf8
}finally{if($null -ne $excel){$excel.Quit();[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}}
