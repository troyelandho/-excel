param([string]$WorkDir)
$ErrorActionPreference='Stop'
$m=Get-Content -LiteralPath (Join-Path $WorkDir 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$excel=$null;$book=$null;$reports=New-Object System.Collections.Generic.List[object]
function Reset-Inputs {
 foreach($property in $m.defaults.PSObject.Properties) {
  $row=$m.input_rows.($property.Name)
  if($property.Value -is [string]) {
   $book.Worksheets.Item(2).Range('B'+$row).Value2=[string]$property.Value
  } else {
   $book.Worksheets.Item(2).Range('B'+$row).Value2=[double]$property.Value
  }
 }
 for($y=0;$y -le 30;$y++) {
  $book.Worksheets.Item(2).Cells.Item($m.build_row,4+$y).Value2=[double]$m.build[$y]
  $book.Worksheets.Item(2).Cells.Item($m.sale_row,4+$y).Value2=[double]$m.sales[$y]
 }
}
function Snapshot([string]$name) {
 $r=[ordered]@{name=$name;scenarios=@();checks=@();decision=$book.Worksheets.Item(10).Range('A5').Value2}
 foreach($off in @(0,80)) {
  $s=[ordered]@{}
  foreach($pr in $m.grids.PSObject.Properties) {
   $sheet=$book.Worksheets.Item(1+[int]$pr.Value[0]);$rows=[ordered]@{}
   foreach($row in $pr.Value[1]) {
    $vals=$sheet.Range($sheet.Cells.Item([int]$row+$off,4),$sheet.Cells.Item([int]$row+$off,34)).Value2
    $a=@();for($j=1;$j -le 31;$j++) { $a+=,$vals.GetValue(1,$j) };$rows.([string]$row)=$a
   };$s.($pr.Name)=$rows
  }
  foreach($pr in $m.scalars.PSObject.Properties) {
   $sheet=$book.Worksheets.Item(1+[int]$pr.Value[0]);$rows=[ordered]@{}
   foreach($row in $pr.Value[1]) { $rows.([string]$row)=$sheet.Cells.Item([int]$row+$(if($pr.Name -eq 'area'){0}else{$off}),2).Value2 };$s.($pr.Name)=$rows
  }
  $col=$(if($off -eq 0){2}else{4});$statusCol=$(if($off -eq 0){4}else{6})
  $s.metrics=[ordered]@{};for($row=6;$row -le 28;$row++) { $s.metrics.([string]$row)=$book.Worksheets.Item(9).Cells.Item($row,$col).Value2 }
  $s.status=@();for($row=34;$row -le 43;$row++) {$s.status+=,$book.Worksheets.Item(9).Cells.Item($row,$statusCol).Value2}
  $s.overall=$book.Worksheets.Item(9).Cells.Item(46,$(if($off -eq 0){2}else{5})).Value2
  $s.reasons=$book.Worksheets.Item(9).Cells.Item($(if($off -eq 0){48}else{50}),2).Value2
  $r.scenarios+=,$s
 }
 for($row=35;$row -le 64;$row++) {$r.checks+=,$book.Worksheets.Item(10).Cells.Item($row,4).Value2}
 $r.chartPoints=@();foreach($chart in $book.Worksheets.Item(10).ChartObjects()) { $r.chartPoints+=,$chart.Chart.SeriesCollection(1).Points().Count }
 return $r
}
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
 $book=$excel.Workbooks.Open((Join-Path $WorkDir 'Excel实际测试.xlsx'),0,$false)
 $excel.Calculation=-4105
 foreach($case in $m.cases) {
  Reset-Inputs
  foreach($pr in $case.overrides.PSObject.Properties) {
   if($pr.Name -eq 'sale_plan') {for($y=0;$y -le 30;$y++) {$book.Worksheets.Item(2).Cells.Item($m.sale_row,4+$y).Value2=[double]$pr.Value[$y]}}
   else { $book.Worksheets.Item(2).Cells.Item([int]$m.input_rows.($pr.Name),2).Value2=[double]$pr.Value }
  }
  $excel.CalculateFullRebuild()
  if($excel.CalculationState -ne 0) {throw 'Excel recalculation incomplete'}
  $r=Snapshot $case.name;$reports.Add($r)
  Write-Output ('Completed: '+$case.name+' / '+$r.decision)
 }
 Reset-Inputs;$excel.CalculateFullRebuild();$book.Save()
 # Export native Excel charts and the opening view of every worksheet for visual QA.
 $qa=Join-Path $WorkDir 'visual_qa';New-Item -ItemType Directory -Path $qa -Force | Out-Null
 for($i=1;$i -le 10;$i++) {
  $sheet=$book.Worksheets.Item($i);$oldArea=$sheet.PageSetup.PrintArea
  $sheet.PageSetup.PrintArea=$(if($i -eq 1){'A1:G36'}elseif($i -eq 2){'A1:K36'}elseif($i -eq 9){'A1:G29'}elseif($i -eq 10){'A1:G32'}else{'A1:J36'})
  $sheet.PageSetup.Zoom=$false;$sheet.PageSetup.FitToPagesWide=1;$sheet.PageSetup.FitToPagesTall=1
  $sheet.ExportAsFixedFormat(0,(Join-Path $qa ('sheet_'+$i+'.pdf')))
  $sheet.PageSetup.PrintArea=$oldArea;$sheet.PageSetup.FitToPagesTall=$false
 }
 $sheet=$book.Worksheets.Item(10);$oldArea=$sheet.PageSetup.PrintArea
 $sheet.PageSetup.PrintArea='A67:H117';$sheet.PageSetup.FitToPagesWide=1;$sheet.PageSetup.FitToPagesTall=$false
 $sheet.ExportAsFixedFormat(0,(Join-Path $qa 'charts.pdf'))
 $sheet.PageSetup.PrintArea=$oldArea
 $r=Snapshot '恢复默认';$reports.Add($r)
 $book.Save()
 [ordered]@{excel_version=$excel.Version;cases=$reports} | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $WorkDir 'actual_results.json') -Encoding UTF8
} finally {
 if($null -ne $book) {$book.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
 if($null -ne $excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
