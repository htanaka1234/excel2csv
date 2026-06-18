[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

if ($IsLinux -or $IsMacOS) {
    throw "SendTo shortcuts can only be removed from Windows."
}

$SendTo = [Environment]::GetFolderPath("SendTo")
if (-not $SendTo) {
    $SendTo = Join-Path $env:APPDATA "Microsoft\Windows\SendTo"
}

$ShortcutNames = @(
    "excel2csv CSV.lnk",
    "excel2csv CSV.GZ.lnk"
)

foreach ($ShortcutName in $ShortcutNames) {
    $ShortcutPath = Join-Path $SendTo $ShortcutName
    if (Test-Path -LiteralPath $ShortcutPath) {
        Remove-Item -LiteralPath $ShortcutPath
        Write-Host "Removed: $ShortcutPath"
    }
    else {
        Write-Host "Not found: $ShortcutPath"
    }
}
