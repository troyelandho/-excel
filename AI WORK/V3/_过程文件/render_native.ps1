param([string]$Path,[switch]$FinalOnly)
$ErrorActionPreference='Stop'
if(-not $Path){$Path=Join-Path $PSScriptRoot '候选_V3.xlsx'}
$out=Join-Path $PSScriptRoot '视觉检查';New-Item -ItemType Directory -Path $out -Force | Out-Null
$excel=$null;$book=$null;$temp=$null
try{
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
 $book=$excel.Workbooks.Open($Path,0,$true);$excel.Calculation=2;$excel.Calculate()
 $tasks=@(
  @('决策看板','A1:G42','01_看板'),@('决策看板','I7:Q40','02_图表'),
  @('核心参数输入','A5:L28','03_原参数'),@('核心参数输入','A93:L117','04_新增参数'),
  @('指标与情景','A4:G54','05_指标'),@('年度汇总','A4:O42','06_年报'),
  @('季度投融资','A6:F33','07_投资'),@('年度所得税','A4:N65','10_所得税'),
  @('季度计划','A4:S30','11_计划'),@('敏感性与边界','A13:J61','12_敏感性'),
  @('敏感性与边界','A64:G84','13_边界'),@('模型校验','A5:E38','14_勾稽'),
  @('使用说明','A2:L51','15_说明'),@('关键月份核对','A2:G40','16_月度')
 )
 if($FinalOnly){$tasks=@(@('决策看板','A1:G42','01_看板'),@('决策看板','I7:Q40','02_图表'),@('敏感性与边界','A64:G84','13_边界'))}
 foreach($t in $tasks){
  $s=$book.Worksheets.Item($t[0]);$s.Unprotect();$s.PageSetup.PrintArea=$s.Range($t[1]).Address()
  $s.PageSetup.Orientation=2;$s.PageSetup.PaperSize=8;$s.PageSetup.Zoom=$false;$s.PageSetup.FitToPagesWide=1;$s.PageSetup.FitToPagesTall=1
  $s.PageSetup.LeftMargin=15;$s.PageSetup.RightMargin=15;$s.PageSetup.TopMargin=18;$s.PageSetup.BottomMargin=18
  $s.ExportAsFixedFormat(0,(Join-Path $out ($t[2]+'.pdf')))
 }
 if(-not $FinalOnly){
 $temp=$excel.Workbooks.Add();$dest=$temp.Worksheets.Item(1)
 foreach($t in @(@('季度经营税费',76,'08_季度经营'),@('季度现金流',79,'09_季度现金'))){
  $dest.Cells.Clear()|Out-Null;$src=$book.Worksheets.Item($t[0]);$src.Unprotect()
  $left=$src.Range('A1:C'+$t[1]);$left.Copy($dest.Range('A1'));$dest.Range('A1:C'+$t[1]).Value2=$left.Value2
  $right=$src.Range('Q4:AC'+$t[1]);$right.Copy($dest.Range('D4'));$dest.Range('D4:P'+$t[1]).Value2=$right.Value2
  $dest.Columns.Item('A').ColumnWidth=43;$dest.Columns.Item('B').ColumnWidth=18;$dest.Columns.Item('C').ColumnWidth=3;$dest.Columns.Item('D:P').ColumnWidth=16
  $dest.UsedRange.RowHeight=21;$dest.Rows.Item(2).RowHeight=30
  $dest.PageSetup.PrintArea=$dest.Range('A1:P'+$t[1]).Address();$dest.PageSetup.Orientation=2;$dest.PageSetup.PaperSize=8;$dest.PageSetup.Zoom=$false;$dest.PageSetup.FitToPagesWide=1;$dest.PageSetup.FitToPagesTall=1
  $dest.ExportAsFixedFormat(0,(Join-Path $out ($t[2]+'.pdf')))
 }
 }
 Write-Output 'Native Excel previews exported.'
}finally{
 if($null -ne $temp){$temp.Close($false);[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($temp)}
 if($null -ne $book){$book.Close($false);[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
 if($null -ne $excel){$excel.Quit();[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
}
