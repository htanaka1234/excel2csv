[CmdletBinding()]
param(
    [switch]$Gzip,
    [string]$Output,
    [string]$Password,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$InputPaths
)

$ErrorActionPreference = "Stop"
$Image = if ($env:EXCEL2CSV_IMAGE) { $env:EXCEL2CSV_IMAGE } else { "excel2csv:local" }
$ProjectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).ProviderPath

if (-not $InputPaths -or $InputPaths.Count -eq 0) {
    throw "Drop one or more Excel files onto excel2csv.cmd, or pass input paths."
}

$ResolvedInputs = @()
foreach ($InputPath in $InputPaths) {
    $ResolvedInputs += (Resolve-Path -LiteralPath $InputPath).ProviderPath
}

if (-not $Output) {
    $FirstInput = Get-Item -LiteralPath $ResolvedInputs[0]
    $OutputDir = if ($FirstInput.PSIsContainer) { $FirstInput.FullName } else { $FirstInput.DirectoryName }
    $Suffix = if ($Gzip) { "csv.gz" } else { "csv" }
    $Stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $Output = Join-Path $OutputDir "merged_$Stamp.$Suffix"
}

$OutputFullPath = [System.IO.Path]::GetFullPath($Output)
$OutputDirPath = Split-Path -Parent $OutputFullPath
$OutputName = Split-Path -Leaf $OutputFullPath
New-Item -ItemType Directory -Force -Path $OutputDirPath | Out-Null

if (-not $Password -and -not $env:EXCEL2CSV_PASSWORD) {
    $SecurePassword = Read-Host "Password for encrypted workbooks (blank for none)" -AsSecureString
    if ($SecurePassword.Length -gt 0) {
        $Bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
        try {
            $Password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($Bstr)
        }
        finally {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($Bstr)
        }
    }
}

function Invoke-NativeQuiet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [string[]]$ArgumentList = @()
    )

    $PreviousErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        & $FilePath @ArgumentList *> $null
        return $LASTEXITCODE
    }
    catch {
        return 1
    }
    finally {
        $ErrorActionPreference = $PreviousErrorActionPreference
    }
}

function Test-DockerCommand {
    param([switch]$UseWsl)

    if ($UseWsl) {
        $ExitCode = Invoke-NativeQuiet -FilePath "wsl" -ArgumentList @("docker", "version")
    }
    else {
        $ExitCode = Invoke-NativeQuiet -FilePath "docker" -ArgumentList @("version")
    }

    return $ExitCode -eq 0
}

$UseWslDocker = $false
if (Test-DockerCommand) {
    Write-Host "Using Docker from this shell."
}
elseif (Get-Command wsl -ErrorAction SilentlyContinue) {
    if (Test-DockerCommand -UseWsl) {
        $UseWslDocker = $true
        Write-Host "Using Docker through WSL."
    }
    else {
        throw "Docker is not available. Start Docker Desktop or install/start Docker inside WSL."
    }
}
else {
    throw "Docker is not available. Start Docker Desktop or install/start Docker inside WSL."
}

function Convert-ToDockerPath {
    param([string]$Path)

    if ($UseWslDocker) {
        $Converted = & wsl wslpath -a $Path
        if ($LASTEXITCODE -ne 0) {
            throw "Could not convert path for WSL Docker: $Path"
        }
        return $Converted.Trim()
    }
    return $Path
}

function Invoke-DockerChecked {
    param([string[]]$DockerArgs)

    if ($UseWslDocker) {
        & wsl docker @DockerArgs
    }
    else {
        & docker @DockerArgs
    }

    if ($LASTEXITCODE -ne 0) {
        throw "docker $($DockerArgs -join ' ') failed with exit code $LASTEXITCODE"
    }
}

function Test-DockerImage {
    param([string]$ImageName)

    if ($UseWslDocker) {
        $ExitCode = Invoke-NativeQuiet -FilePath "wsl" -ArgumentList @("docker", "image", "inspect", $ImageName)
    }
    else {
        $ExitCode = Invoke-NativeQuiet -FilePath "docker" -ArgumentList @("image", "inspect", $ImageName)
    }

    return $ExitCode -eq 0
}

if (-not (Test-DockerImage -ImageName $Image)) {
    $BuildContext = Convert-ToDockerPath -Path $ProjectRoot
    Invoke-DockerChecked -DockerArgs @("build", "-t", $Image, $BuildContext)
}

$OutputMount = Convert-ToDockerPath -Path $OutputDirPath
$DockerArgs = @("run", "--rm", "-v", "${OutputMount}:/output")
$ContainerInputs = @()

for ($Index = 0; $Index -lt $ResolvedInputs.Count; $Index++) {
    $InputItem = Get-Item -LiteralPath $ResolvedInputs[$Index]
    if ($InputItem.PSIsContainer) {
        $InputMount = Convert-ToDockerPath -Path $InputItem.FullName
        $DockerArgs += @("-v", "${InputMount}:/input${Index}:ro")
        $ContainerInputs += "/input${Index}"
    }
    else {
        $InputMount = Convert-ToDockerPath -Path $InputItem.DirectoryName
        $DockerArgs += @("-v", "${InputMount}:/input${Index}:ro")
        $ContainerInputs += "/input${Index}/$($InputItem.Name)"
    }
}

if ($Password) {
    $DockerArgs += @("-e", "EXCEL2CSV_PASSWORD=$Password")
}

$DockerArgs += $Image
$DockerArgs += $ContainerInputs
$DockerArgs += @("-o", "/output/$OutputName")

if ($Gzip) {
    $DockerArgs += "--gzip"
}

Invoke-DockerChecked -DockerArgs $DockerArgs
Write-Host "Wrote $OutputFullPath"
