param([string]$Path)
$ErrorActionPreference='Stop'
if(-not $Path){$Path=Join-Path $PSScriptRoot '候选_V3.xlsx'}
$excel=$null;$book=$null;$tests=[System.Collections.Generic.List[object]]::new()
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
 $book=$excel.Workbooks.Open($Path,0,$true)
 $sens=$book.Worksheets.Item('敏感性与边界');$metric=$book.Worksheets.Item('指标与情景');$inp=$book.Worksheets.Item('核心参数输入');$month=$book.Worksheets.Item('关键月份核对')
 foreach($s in $book.Worksheets){$s.Unprotect()}
 $excel.Calculation=2;$excel.Calculate()
 foreach($t in @(@{cell='G18';metric='B8'},@{cell='G28';metric='B50'},@{cell='G38';metric='B8'},@{cell='G48';metric='B13'},@{cell='F56';metric='B8'})){
  $tests.Add(@{test='base_'+$t.cell;table=$sens.Range($t.cell).Value2;model=$metric.Range($t.metric).Value2;formula=$sens.Range($t.cell).Formula})
 }
 $expected=$sens.Range('E16').Value2;$sens.Range('B7').Value2=[double]($inp.Range('B31').Value2*0.8);$sens.Range('B8').Value2=[double]($inp.Range('B36').Value2*0.8);$excel.Calculate()
 $tests.Add(@{test='price_rent_80pct';table=$expected;model=$metric.Range('B8').Value2})
 $sens.Range('B7:B8').ClearContents()|Out-Null;$excel.Calculate()
 $expected=$sens.Range('E58').Value2;$sens.Range('B7').Value2=[double]($inp.Range('B31').Value2*0.9);$sens.Range('B10').Value2=2.0;$excel.Calculate()
 $tests.Add(@{test='delay_2_price_90pct';table=$expected;model=$metric.Range('B8').Value2})
 $sens.Range('B7').ClearContents()|Out-Null;$sens.Range('B10').ClearContents()|Out-Null;$excel.Calculate()
 $baseGap=$month.Range('B33').Value2;$baseDiff=$month.Range('G28').Value2
 $month.Range('B13').Value2=0.0;$month.Range('C13').Value2=0.0;$month.Range('D13').Value2=1.0;$excel.Calculate()
 $tests.Add(@{test='monthly_late_draw';baseGap=$baseGap;lateGap=$month.Range('B33').Value2;quarterReconciliation=$month.Range('G28').Value2;baseReconciliation=$baseDiff;valid=$month.Range('B34').Value2})
 $tests.ToArray() | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'native_table_checks.json') -Encoding utf8
 $tests.ToArray() | ConvertTo-Json -Depth 4
}finally{
 if($null -ne $book){$book.Close($false);[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
 if($null -ne $excel){$excel.Quit();[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
}
