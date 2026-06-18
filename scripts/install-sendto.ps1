[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"

if ($IsLinux -or $IsMacOS) {
    throw "SendTo shortcuts can only be installed from Windows."
}

$ProjectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).ProviderPath
$SendTo = [Environment]::GetFolderPath("SendTo")
if (-not $SendTo) {
    $SendTo = Join-Path $env:APPDATA "Microsoft\Windows\SendTo"
}

New-Item -ItemType Directory -Force -Path $SendTo | Out-Null

$Shortcuts = @(
    @{
        Name = "excel2csv CSV.lnk"
        Target = Join-Path $ProjectRoot "excel2csv.cmd"
        Description = "Merge Excel workbooks into UTF-8 BOM CSV"
    },
    @{
        Name = "excel2csv CSV.GZ.lnk"
        Target = Join-Path $ProjectRoot "excel2csv-gzip.cmd"
        Description = "Merge Excel workbooks into gzip UTF-8 CSV"
    }
)

$Shell = New-Object -ComObject WScript.Shell

foreach ($ShortcutSpec in $Shortcuts) {
    $ShortcutPath = Join-Path $SendTo $ShortcutSpec.Name
    if ((Test-Path -LiteralPath $ShortcutPath) -and -not $Force) {
        Write-Host "Already exists: $ShortcutPath"
        continue
    }

    $Shortcut = $Shell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $ShortcutSpec.Target
    $Shortcut.WorkingDirectory = $ProjectRoot
    $Shortcut.Description = $ShortcutSpec.Description
    $Shortcut.Save()

    Write-Host "Installed: $ShortcutPath"
}

Write-Host "Right-click Excel files, then choose Send to -> excel2csv CSV or excel2csv CSV.GZ."
