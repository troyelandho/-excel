$ErrorActionPreference='Stop'
$auditDir=$PSScriptRoot
$sourcePath=Join-Path (Split-Path $auditDir -Parent) 'V2/产业园投资测算_V2_样式调整.xlsx'
$copyPath=Join-Path $auditDir 'review_temporary.xlsx'
Copy-Item -LiteralPath $sourcePath -Destination $copyPath -Force
$cases=@(
 @{name='base'; edits=@{}},
 @{name='pressure_zero'; edits=@{B78=0; B79=0; B80=0; B81=0}},
 @{name='no_loan'; edits=@{B57=0}},
 @{name='equity_10000'; edits=@{B61=10000}},
 @{name='disposal'; edits=@{B64=1}},
 @{name='horizon_10_disposal'; edits=@{B9=10; B64=1}},
 @{name='horizon_30'; edits=@{B9=30}},
 @{name='horizon_2_invalid'; edits=@{B9=2}},
 @{name='construction_4'; edits=@{B16=4; B17=5; B18=5; B60=5}; construction=@(0,0.25,0.25,0.25,0.25); sales=@(0,0,0,0,0,0.3,0.4,0.3)},
 @{name='first_draw_3'; edits=@{}; construction=@(0,0,0,1)},
 @{name='pure_hold'; edits=@{B14=0}},
 @{name='pure_sale'; edits=@{B14=1}},
 @{name='blank_rent'; edits=@{B36=$null}},
 @{name='zero_rent'; edits=@{B36=0}},
 @{name='negative_cost'; edits=@{B24=-4500}},
 @{name='zero_collection'; edits=@{B43=0}},
 @{name='high_sale_price'; edits=@{B31=12000}},
 @{name='zero_sale_price'; edits=@{B31=0}},
 @{name='pressure_100'; edits=@{B78=1; B79=1; B80=100; B82=1}}
)
$excel=$null
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false
 $excel.DisplayAlerts=$false
 $excel.AskToUpdateLinks=$false
 foreach($case in $cases){
   $book=$null
   try {
     $book=$excel.Workbooks.Open($copyPath,0,$true)
     $inputs=$book.Worksheets.Item('核心参数输入')
     foreach($address in $case.edits.Keys){
       if($null -eq $case.edits[$address]){$inputs.Range($address).ClearContents() | Out-Null}
       else {$inputs.Range($address).Value2=[double]$case.edits[$address]}
     }
     foreach($p in @(@{key='construction';row=88},@{key='sales';row=89})){
       if($case.ContainsKey($p.key)){
         for($i=0;$i -lt 31;$i++){
           $v=0.0
           if($i -lt $case[$p.key].Count){$v=[double]$case[$p.key][$i]}
           $inputs.Cells.Item($p.row,$i+4).Value2=$v
         }
       }
     }
     $excel.CalculateFullRebuild()
     $sheets=@{}
     $errors=@()
     foreach($sheet in $book.Worksheets){
       $range=$sheet.UsedRange
       $values=$range.Value2
       $rows=@()
       for($r=1;$r -le $range.Rows.Count;$r++){
         $rr=@()
         for($c=1;$c -le $range.Columns.Count;$c++){$rr+=,$values[$r,$c]}
         $rows+=,@($rr)
       }
       $sheets[$sheet.Name]=$rows
       try {
         $bad=$sheet.Cells.SpecialCells(-4123,16)
         foreach($cell in $bad.Cells){$errors+=@{sheet=$sheet.Name;address=$cell.Address($false,$false);text=$cell.Text}}
       } catch {}
     }
     $summary=@{case=$case.name;excel=$excel.Version;errors=$errors;sheets=$sheets}
     $outPath=Join-Path $auditDir ('native_'+$case.name+'.json')
     $summary | ConvertTo-Json -Depth 8 -Compress | Set-Content -LiteralPath $outPath -Encoding utf8
     Write-Output ($case.name+' errors='+$errors.Count)
   } finally {if($null -ne $book){$book.Close($false);[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}}
 }
} finally {
 if($null -ne $excel){$excel.Quit();[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
}
