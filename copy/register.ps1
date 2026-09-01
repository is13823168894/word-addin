$copyDir = "E:\A001 重要软件\OFFICE加载项\word-addin\copy"
$manifestFile = Join-Path $copyDir "manifest-copy.xml"

if (!(Test-Path $manifestFile)) {
    Write-Host "未找到 manifest-copy.xml" -ForegroundColor Red
    exit 1
}

$trustedCatalogsPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
if (!(Test-Path $trustedCatalogsPath)) {
    New-Item -Path $trustedCatalogsPath -Force | Out-Null
}

$hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash(
    [System.Text.Encoding]::UTF8.GetBytes($copyDir)
)
$guidBytes = New-Object byte[] 16
[Array]::Copy($hash, $guidBytes, 16)
$catalogGuid = (New-Object Guid $guidBytes).ToString()
$catalogPath = "$trustedCatalogsPath\$catalogGuid"

New-Item -Path $catalogPath -Force | Out-Null
New-ItemProperty -Path $catalogPath -Name "Id" -Value $catalogGuid -PropertyType String -Force | Out-Null
New-ItemProperty -Path $catalogPath -Name "Url" -Value $copyDir -PropertyType String -Force | Out-Null
New-ItemProperty -Path $catalogPath -Name "Flags" -Value 1 -PropertyType DWord -Force | Out-Null

Write-Host ""
Write-Host "已注册共享文件夹加载项：" -ForegroundColor Green
Write-Host "  目录: $copyDir" -ForegroundColor Green
Write-Host "  GUID: $catalogGuid" -ForegroundColor Green
Write-Host ""
Write-Host "现在打开 Word -> 插入 -> 加载项 -> 我的加载项 -> 共享文件夹" -ForegroundColor Yellow
Write-Host "应该能看到「段落复制」加载项。" -ForegroundColor Yellow
