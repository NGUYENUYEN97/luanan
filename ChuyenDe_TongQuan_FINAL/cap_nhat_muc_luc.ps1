# Mở file chuyên đề bằng Word, cập nhật mục lục và danh mục bảng/hình, lưu lại và xuất bản PDF để kiểm tra.
param(
  [string]$Docx = "G:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\ChuyenDe_TongQuan_FINAL.docx",
  [string]$Pdf = ""
)
$w = New-Object -ComObject Word.Application
$w.Visible = $false
try {
  $d = $w.Documents.Open($Docx, $false, $false)
  $d.Repaginate()
  foreach ($toc in $d.TablesOfContents) { $toc.Update() }
  $d.Fields.Update() | Out-Null
  foreach ($toc in $d.TablesOfContents) { $toc.Update() }
  $d.Save()
  if ($Pdf -ne "") { $d.SaveAs2($Pdf, 17) }
  "Pages: " + $d.ComputeStatistics(2)
  $d.Close($false)
} finally { $w.Quit() }
