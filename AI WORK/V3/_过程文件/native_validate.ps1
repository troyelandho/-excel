param([string]$Path,[string]$OutputName='native_candidate.json',[switch]$Save)
$ErrorActionPreference='Stop'
if(-not $Path){$Path=Join-Path $PSScriptRoot '候选_V3.xlsx'}
$manifest=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'manifest.json') -Raw -Encoding utf8 | ConvertFrom-Json
$excel=$null;$book=$null
try {
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
 $book=$excel.Workbooks.Open($Path,0,$false)
 $excel.CalculateFullRebuild()
 $errors=[System.Collections.Generic.List[object]]::new()
 $sheets=@{}
 foreach($sheet in $book.Worksheets){
  $range=$sheet.UsedRange;$vals=$range.Value2;$nrows=$range.Rows.Count;$ncols=$range.Columns.Count
  $rows=[System.Collections.Generic.List[object]]::new()
  for($r=1;$r -le $nrows;$r++){
   $row=[System.Collections.Generic.List[object]]::new()
   for($c=1;$c -le $ncols;$c++){$row.Add($vals[$r,$c])}
   $rows.Add($row.ToArray())
  }
  $sheets[$sheet.Name]=@{row=$range.Row;col=$range.Column;values=$rows.ToArray()}
  try{$bad=$sheet.Cells.SpecialCells(-4123,16);foreach($cell in $bad.Cells){$errors.Add(@{sheet=$sheet.Name;address=$cell.Address($false,$false);value=$cell.Text;formula=$cell.Formula})}}catch{}
 }
 $result=@{excel=$excel.Version;errors=$errors.ToArray();sheets=$sheets;file=$Path}
 $result | ConvertTo-Json -Depth 9 -Compress | Set-Content -LiteralPath (Join-Path $PSScriptRoot $OutputName) -Encoding utf8
 Write-Output ('Native formula errors: '+$errors.Count)
 if($errors.Count -gt 0){$errors | Select-Object -First 12 | ConvertTo-Json -Depth 4}
 if($Save){$book.Save()}
} finally {
 if($null -ne $book){$book.Close($false);[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
 if($null -ne $excel){$excel.Quit();[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
}
