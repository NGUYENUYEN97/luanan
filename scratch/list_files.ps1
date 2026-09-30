$targetDir = "g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\01_NotebookLM_Inputs\05_BaiBao_QuocTe"
$files = Get-ChildItem -Path $targetDir -File
Write-Host "Total files: $($files.Count)"
foreach ($f in $files) {
    Write-Host "$($f.Name)"
}
