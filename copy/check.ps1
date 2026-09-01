$p = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
if (Test-Path $p) {
    Get-ChildItem $p | ForEach-Object {
        $name = $_.PSChildName
        $url = $_.GetValue("Url")
        Write-Host "$name -> $url"
    }
} else {
    Write-Host "TrustedCatalogs not found"
}
