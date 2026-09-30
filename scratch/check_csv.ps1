$csvPath = "g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\01_NotebookLM_Inputs\05_BaiBao_QuocTe\scopus_export_Aug 22-2026_352f2f0d-45fb-4ce5-8034-b53ac9330e50.csv"
$csv = Import-Csv -Path $csvPath
Write-Host "Total rows in CSV: $($csv.Count)"
for ($i = 0; $i -lt [Math]::Min($csv.Count, 15); $i++) {
    Write-Host "--- Row $i ---"
    Write-Host "Title: $($csv[$i].Title)"
    Write-Host "Authors: $($csv[$i].Authors)"
    Write-Host "Year: $($csv[$i].Year)"
    Write-Host "Source: $($csv[$i].'Source title')"
    Write-Host "DOI: $($csv[$i].DOI)"
}
