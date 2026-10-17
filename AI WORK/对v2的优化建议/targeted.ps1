$ErrorActionPreference='Stop'
$copyPath=Join-Path $PSScriptRoot 'review_temporary.xlsx'
$excel=$null
$result=@()
try{
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false
 $excel.DisplayAlerts=$false
 foreach($name in @('high_sale_16000','wc_shifted','disposal_shifted')){
  $book=$null
  try{
   $book=$excel.Workbooks.Open($copyPath,0,$true)
   $inputs=$book.Worksheets.Item('核心参数输入')
   if($name -eq 'high_sale_16000'){$inputs.Range('B31').Value2=16000.0}
   if($name -eq 'wc_shifted'){
    $i=$book.Worksheets.Item('投资与资金安排')
    $i.Range('W32').Value2=500.0; $i.Range('X32').Value2=0.0
    $i.Range('W112').Value2=500.0; $i.Range('X112').Value2=0.0
   }
   if($name -eq 'disposal_shifted'){
    $inputs.Range('B64').Value2=1.0
    $i=$book.Worksheets.Item('销售与出租收入')
    $i.Range('W24').Value2=30000.0; $i.Range('X24').Value2=0.0
    $i.Range('W104').Value2=30000.0; $i.Range('X104').Value2=0.0
   }
   $excel.CalculateFullRebuild()
   $r=@{case=$name;check62=$book.Worksheets.Item('决策看板').Range('D62').Value2;payback=$book.Worksheets.Item('指标与情景测算').Range('B10').Value2;discounted_payback=$book.Worksheets.Item('指标与情景测算').Range('B11').Value2}
   if($name -eq 'high_sale_16000'){
    $flows=@();$disc=@();$c=$book.Worksheets.Item('利润与现金流');$v=$c.Range('D26:AH26').Value2;$w=$c.Range('D50:AH50').Value2
    for($t=1;$t -le 31;$t++){$flows+=,$v[1,$t];$disc+=,$w[1,$t]}
    $r['flows']=$flows;$r['discounted_flows']=$disc
   }
   $result+=,$r
   Write-Output ($name+' completed')
  }finally{if($null -ne $book){$book.Close($false);[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}}
 }
 $result | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'targeted_results.json') -Encoding utf8
}finally{if($null -ne $excel){$excel.Quit();[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}}
