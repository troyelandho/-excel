param([string]$WorkDirectory)
$ErrorActionPreference='Stop'
$work=(Resolve-Path -LiteralPath $WorkDirectory).Path
$manifest=Get-Content -LiteralPath (Join-Path $work 'case_manifest.json') -Raw -Encoding utf8 | ConvertFrom-Json
$temp=Join-Path $work 'Excel重算副本.xlsx'
if(Test-Path -LiteralPath $temp){
 if((Get-FileHash -LiteralPath $temp).Hash -ne (Get-FileHash -LiteralPath (Join-Path $work '参数已填入_待重算.xlsx')).Hash){throw '重算文件已存在且已更改，停止覆盖'}
}else{Copy-Item -LiteralPath (Join-Path $work '参数已填入_待重算.xlsx') -Destination $temp}
$app=$null;$wb=$null;$dummy=$null
function CellValue($cell){ return $cell.Value2 }
function Snapshot($wb,$name){
 $met=$wb.Worksheets.Item('指标与情景测算');$dash=$wb.Worksheets.Item('决策看板')
 $base=@{};$stress=@{}
 foreach($p in $manifest.maps.metrics.PSObject.Properties){
  $base[$p.Name]=CellValue $met.Cells.Item([int]$p.Value,2)
  $stress[$p.Name]=CellValue $met.Cells.Item([int]$p.Value,4)
 }
 $statuses=@();for($r=19;$r -le 28;$r++){$statuses+=@{row=$r;base=$dash.Cells.Item($r,4).Value2;stress=$dash.Cells.Item($r,6).Value2}}
 $checks=@();for($r=35;$r -le 64;$r++){if($dash.Cells.Item($r,4).HasFormula){$checks+=@{row=$r;label=$dash.Cells.Item($r,1).Value2;value=$dash.Cells.Item($r,4).Value2}}}
 $grids=@{}
 $spec=@{plan='建筑面积与开发计划';investment_annual='投资与资金安排';income='销售与出租收入';costs='运营成本与税费';financing='融资与还本付息';cashflow='利润与现金流'}
 foreach($group in $spec.Keys){
  $s=$wb.Worksheets.Item($spec[$group]);$b=@{};$t=@{}
  foreach($r in $manifest.maps.$group.PSObject.Properties){
   if([int]$r.Value -lt 10 -and $group -eq 'financing'){continue}
   $arr=@();$arr2=@();$a=$s.Range(('D'+$r.Value+':AH'+$r.Value)).Value2;$b2=$s.Range(('D'+([int]$r.Value+80)+':AH'+([int]$r.Value+80))).Value2
   for($c=1;$c -le 31;$c++){$arr+=$a[1,$c];$arr2+=$b2[1,$c]}
   $b[$r.Name]=$arr;$t[$r.Name]=$arr2
  }
  $grids[$group]=@{base=$b;stress=$t}
 }
 return @{name=$name;base=$base;stress=$stress;statuses=$statuses;checks=$checks;decision=$dash.Range('A5').Value2;reasons=$dash.Range('A15').Value2;grids=$grids}
}
try{
 $app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 $dummy=$app.Workbooks.Add();$app.Calculation=-4135
 $wb=$app.Workbooks.Open($temp,0,$false)
 $dummy.Close($false);$dummy=$null
 $app.Calculation=-4105;$app.CalculateFullRebuild()
 if($app.CalculationState -ne 0){throw 'Excel计算未完成'}
 $wb.Save()
 Write-Output '已实际重算并保存方案；开始参数联动测试。'
 $results=@();$input=$wb.Worksheets.Item('核心参数输入')
 foreach($case in $manifest.cases){
  foreach($p in $manifest.params.PSObject.Properties){if($manifest.input_rows.$($p.Name)){if($p.Value -is [string]){$input.Cells.Item([int]$manifest.input_rows.$($p.Name),2).Value2=[string]$p.Value}else{$input.Cells.Item([int]$manifest.input_rows.$($p.Name),2).Value2=[double]$p.Value}}}
  foreach($p in $case.overrides.PSObject.Properties){$input.Cells.Item([int]$manifest.input_rows.$($p.Name),2).Value2=[double]$p.Value}
  $app.CalculateFullRebuild()
  $results+=Snapshot $wb $case.name
  Write-Output ('测试完成：'+$case.name)
 }
 @{ExcelVersion=$app.Version;Calculation='CalculateFullRebuild';Cases=$results} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $work 'Excel实际重算及情景结果.json') -Encoding utf8
 # 临时测试不保存。已保存的副本仍是交付方案参数。
}finally{
 if($null -ne $wb){$wb.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)}
 if($null -ne $dummy){$dummy.Close($false)}
 if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
