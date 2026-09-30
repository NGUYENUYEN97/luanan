$content = Get-Content -Raw -Encoding UTF8 -Path "g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\01_NotebookLM_Inputs\05_BaiBao_QuocTe\_all_abstracts.txt"
$sections = $content -split "(?m)^--- FILE: "
Write-Host "Total sections: $($sections.Count)"
for ($i = 1; $i -lt [Math]::Min($sections.Count, 15); $i++) {
    $sec = $sections[$i].Trim()
    $firstLine = ($sec -split "`r?`n")[0]
    Write-Host "[$i] $firstLine"
}
