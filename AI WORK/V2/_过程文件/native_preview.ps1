param([string]$InputPath,[string]$OutputDirectory,[switch]$AllSheets,[switch]$Reference,[int]$StartSheet=1)
$ErrorActionPreference='Stop'
$inputResolved=(Resolve-Path -LiteralPath $InputPath).Path
$previewDirectory=[System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $previewDirectory -Force | Out-Null
$temporaryCopy=Join-Path $previewDirectory 'preview_copy.xlsx'
Copy-Item -LiteralPath $inputResolved -Destination $temporaryCopy
$excel=$null;$book=$null;$dummy=$null
try {
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false
  $excel.AutomationSecurity=3
  $dummy=$excel.Workbooks.Add()
  $excel.Calculation=-4135
  $book=$excel.Workbooks.Open($temporaryCopy,0,$true)
  $dummy.Close($false);$dummy=$null
  $rows=@()
  if ($Reference) {
    $sheet=$book.Worksheets.Item('9决策看板')
    $sheet.PageSetup.PrintArea='A1:E23'
    $sheet.PageSetup.Zoom=$false;$sheet.PageSetup.FitToPagesWide=1;$sheet.PageSetup.FitToPagesTall=1
    $sheet.ExportAsFixedFormat(0,(Join-Path $previewDirectory 'reference_dashboard.pdf'))
  } else {
    $indices=if ($AllSheets) {$StartSheet..10} else {2,10}
    foreach ($i in $indices) {
      $sheet=$book.Worksheets.Item($i)
      $ranges=if ($i -eq 1) {@('A1:G36','A37:G83')} elseif ($i -eq 2) {@('A1:K44','A46:K80','A82:AF92')} elseif ($i -eq 9) {@('A1:G29','A31:G50')} elseif ($i -eq 10) {@('A1:G32','A33:G64','A66:H118')} else {@('A1:J38',('A81:J'+$(if($i -eq 8){146}else{116})))}
      for ($j=0;$j -lt $ranges.Count;$j++) {
        $sheet.PageSetup.PrintArea=$ranges[$j]
        $sheet.PageSetup.Zoom=$false;$sheet.PageSetup.FitToPagesWide=1;$sheet.PageSetup.FitToPagesTall=1
        $sheet.ExportAsFixedFormat(0,(Join-Path $previewDirectory ('sheet_'+$i+'_'+($j+1)+'.pdf')))
      }
    }
    $hashes=@()
    $errors=@()
    $textCells=0
    $fontNames=@{}
    foreach ($sheet in $book.Worksheets) {
      foreach ($cell in $sheet.UsedRange.Cells) {
        if ($null -ne $cell.Value2) {
          $textCells++
          $display=[string]$cell.Text
          if ($display -match '^#+$') {$hashes+=($sheet.Name+'!'+$cell.Address($false,$false))}
          if ($display -match '^#(REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|NULL!|SPILL!|CALC!)$') {$errors+=($sheet.Name+'!'+$cell.Address($false,$false))}
          $fontNames[[string]$cell.Font.Name]=$true
        }
      }
    }
    [pscustomobject]@{ExcelVersion=$excel.Version;SheetCount=$book.Worksheets.Count;ChartCount=$book.Worksheets.Item(10).ChartObjects().Count;NonblankDisplayCells=$textCells;HashDisplayCells=$hashes;FormulaErrorDisplayCells=$errors;CellFontNames=@($fontNames.Keys);CalculationMode='Manual';Saved=$false} | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $previewDirectory 'native_check.json') -Encoding utf8
  }
  Write-Output ('预览完成：'+$previewDirectory)
} finally {
  if ($null -ne $book) {$book.Close($false);[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
  if ($null -ne $dummy) {$dummy.Close($false)}
  if ($null -ne $excel) {$excel.Quit();[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
