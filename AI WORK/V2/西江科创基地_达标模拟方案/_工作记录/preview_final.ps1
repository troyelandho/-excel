param([string]$WorkDirectory,[switch]$InputOnly)
$ErrorActionPreference='Stop'
$work=(Resolve-Path -LiteralPath $WorkDirectory).Path
$root=Split-Path $work -Parent
$view=Join-Path $work $(if($InputOnly){'行高修正检查'}else{'视觉检查'})
New-Item -ItemType Directory -Path $view -Force | Out-Null
$copy=Join-Path $view '预览临时副本.xlsx'
if(Test-Path -LiteralPath $copy){throw '预览副本已存在'}
Copy-Item -LiteralPath (Join-Path $root '西江科创基地_投资测算_达标模拟.xlsx') -Destination $copy
$app=$null;$book=$null;$dummy=$null
try{
 $app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
 $dummy=$app.Workbooks.Add();$app.Calculation=-4135
 $book=$app.Workbooks.Open($copy,0,$true);$dummy.Close($false);$dummy=$null
 $ranges=@{1=@('A1:G36','A65:G83');2=@('A1:K44','A46:K82','A94:K107');9=@('A1:F29','A31:G50');10=@('A1:G32','A33:G64','A66:H118')}
 $indices=if($InputOnly){@(2)}else{@(1,2,9,10)}
 if($InputOnly){$ranges[2]=@('A1:K44')}
 foreach($i in $indices){
  $sheet=$book.Worksheets.Item($i);$j=0
  foreach($range in $ranges[$i]){
   $j++;$sheet.PageSetup.PrintArea=$range;$sheet.PageSetup.Zoom=$false
   $sheet.PageSetup.FitToPagesWide=1;$sheet.PageSetup.FitToPagesTall=1
   $sheet.ExportAsFixedFormat(0,(Join-Path $view ('sheet_'+$i+'_'+$j+'.pdf')))
  }
 }
 $hashes=@();$critical=@()
 foreach($i in 2,9,10){
  $sheet=$book.Worksheets.Item($i)
  $range=if($i -eq 2){'A1:K82'}elseif($i -eq 9){'A1:G50'}else{'A1:H118'}
  foreach($c in $sheet.Range($range).Cells){
   if($null -ne $c.Value2){
    if([string]$c.Text -match '^#+$'){$hashes+=($sheet.Name+'!'+$c.Address($false,$false))}
    if([string]$c.Text -match '^#(REF!|DIV/0!|VALUE!|NAME\?|NUM!|N/A|NULL!)$'){$critical+=($sheet.Name+'!'+$c.Address($false,$false))}
   }
  }
 }
 @{ExcelVersion=$app.Version;HashDisplays=$hashes;FormulaErrors=$critical;ChartCount=$book.Worksheets.Item(10).ChartObjects().Count;Decision=$book.Worksheets.Item(10).Range('A5').Text;Saved=$false} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $view '视觉检查结果.json') -Encoding utf8
 Write-Output '10个重点区域已导出，三张原图表已保留。'
}finally{
 if($null -ne $book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
 if($null -ne $dummy){$dummy.Close($false)}
 if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
