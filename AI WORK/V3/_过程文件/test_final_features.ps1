$ErrorActionPreference='Stop'
$scriptPath=Join-Path (Split-Path $PSScriptRoot -Parent) '刷新_V3.ps1'
$ast=[System.Management.Automation.Language.Parser]::ParseFile($scriptPath,[ref]$null,[ref]$null)
foreach($n in @('Recalc','Get-Result','Find-Boundary')){
 $fn=$ast.Find({param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n},$true)
 Invoke-Expression $fn.Extent.Text
}
$excel=$null;$book=$null
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
 $book=$excel.Workbooks.Open((Join-Path (Split-Path $PSScriptRoot -Parent) '产业园投资测算_V3_决策测算.xlsx'),0,$true);$excel.Calculation=2
 foreach($s in $book.Worksheets){$s.Unprotect()}
 $inputs=$book.Worksheets.Item('核心参数输入');$sens=$book.Worksheets.Item('敏感性与边界');$metrics=$book.Worksheets.Item('指标与情景');$testSheet=$sens
 $chart=$book.Worksheets.Item('决策看板').ChartObjects().Item(1).Chart
 $chart.Refresh();$initialCount=$chart.SeriesCollection(1).Points().Count;$initialStatus=$sens.Range('B66').Value2
 $inputs.Range('B9').Value2=30.0;$chart.Refresh();$count30=$chart.SeriesCollection(1).Points().Count;$stale=$sens.Range('B66').Value2
 $inputs.Range('B9').Value2=20.0;$chart.Refresh();$count20=$chart.SeriesCollection(1).Points().Count
 $sens.Range('B6').Value2=0.0
 $r=Find-Boundary 'B6' 3200 $true $false
 $zero=$sens.Range('B6').Value2;$zeroCount=$excel.WorksheetFunction.Count($sens.Range('B6'))
 $sens.Range('B6').ClearContents()|Out-Null
 $price=[double]$sens.Range('B75').Value2
 $sens.Range('B7').Value2=$price-0.01;$below=$metrics.Range('B8').Value2
 $sens.Range('B7').Value2=$price+0.01;$above=$metrics.Range('B8').Value2
 $sens.Range('B7').ClearContents()|Out-Null
 $rent=[double]$sens.Range('B78').Value2
 $rentBelow=Get-Result 'B8' ($rent-0.001) $true;$rentAbove=Get-Result 'B8' ($rent+0.001) $true
 $sens.Range('B8').ClearContents()|Out-Null
 $result=@{chartPointsInitial=$initialCount;chartPoints30=$count30;chartPointsRestored=$count20;initialSignatureStatus=$initialStatus;changedSignatureStatus=$stale;restoredSignatureStatus=$sens.Range('B66').Value2;zeroOverrideAfterRefresh=$zero;zeroOverrideNumericCount=$zeroCount;priceNPVBelow=$below;priceNPVAbove=$above;rentAllConditionsBelow=$rentBelow;rentAllConditionsAbove=$rentAbove}
 $result | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'final_feature_validation.json') -Encoding utf8
 $result | ConvertTo-Json -Depth 4
}finally{
 if($null -ne $book){$book.Close($false);[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
 if($null -ne $excel){$excel.Quit();[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
}
