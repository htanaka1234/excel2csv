[CmdletBinding()]
param(
    [switch]$Gzip,
    [switch]$AskPassword,
    [string]$Output,
    [string]$Password,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$InputPaths
)

$ErrorActionPreference = "Stop"
$Image = if ($env:EXCEL2CSV_IMAGE) { $env:EXCEL2CSV_IMAGE } else { "excel2csv:local" }
$ProjectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).ProviderPath
$FingerprintLabel = "org.opencontainers.image.source-fingerprint"

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

$ShouldAskPassword = $AskPassword -or (-not $Password -and -not $env:EXCEL2CSV_PASSWORD)

function Invoke-NativeCapture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [string[]]$ArgumentList = @()
    )

    $PreviousErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $Output = & $FilePath @ArgumentList 2>$null
        return [pscustomobject]@{
            ExitCode = $LASTEXITCODE
            Output = ($Output -join "`n")
        }
    }
    catch {
        return [pscustomobject]@{
            ExitCode = 1
            Output = ""
        }
    }
    finally {
        $ErrorActionPreference = $PreviousErrorActionPreference
    }
}

function Invoke-NativeQuiet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [string[]]$ArgumentList = @()
    )

    $Result = Invoke-NativeCapture -FilePath $FilePath -ArgumentList $ArgumentList
    return $Result.ExitCode
}

function Test-DockerCommand {
    param([switch]$UseWsl)

    if ($UseWsl) {
        $ExitCode = Invoke-NativeQuiet -FilePath "wsl.exe" -ArgumentList @("--exec", "docker", "version")
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
elseif (Get-Command wsl.exe -ErrorAction SilentlyContinue) {
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
        $Result = Invoke-NativeCapture -FilePath "wsl.exe" -ArgumentList @("--exec", "wslpath", "-u", "-a", $Path)
        $Converted = $Result.Output.Trim()
        if ($Result.ExitCode -ne 0 -or -not $Converted) {
            throw "Could not convert path for WSL Docker: $Path"
        }
        return $Converted
    }
    return $Path
}

function Invoke-DockerChecked {
    param([string[]]$DockerArgs)

    if ($UseWslDocker) {
        & wsl.exe --exec docker @DockerArgs
    }
    else {
        & docker @DockerArgs
    }

    if ($LASTEXITCODE -ne 0) {
        throw "docker $($DockerArgs -join ' ') failed with exit code $LASTEXITCODE"
    }
}

function Get-ProjectRelativePath {
    param([string]$Path)

    $FullPath = (Resolve-Path -LiteralPath $Path).ProviderPath
    $PathSeparators = [char[]]@([char]92, [char]47)
    $Root = $ProjectRoot.TrimEnd($PathSeparators)
    if (-not $FullPath.StartsWith($Root, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Path is outside project root: $Path"
    }

    return $FullPath.Substring($Root.Length).TrimStart($PathSeparators).Replace("\", "/")
}

function Test-FingerprintFile {
    param([string]$RelativePath)

    return $RelativePath -notmatch "(^|/)__pycache__/" -and $RelativePath -notmatch "\.pyc$"
}

function Get-SourceFingerprint {
    $RelativePaths = New-Object System.Collections.Generic.List[string]

    foreach ($RelativePath in @("Dockerfile", "pyproject.toml", "README.md")) {
        $FullPath = Join-Path $ProjectRoot $RelativePath
        if (Test-Path -LiteralPath $FullPath -PathType Leaf) {
            $RelativePaths.Add($RelativePath)
        }
    }

    foreach ($Directory in @("src", "tests")) {
        $FullDirectory = Join-Path $ProjectRoot $Directory
        if (Test-Path -LiteralPath $FullDirectory -PathType Container) {
            foreach ($File in Get-ChildItem -LiteralPath $FullDirectory -File -Recurse) {
                $RelativePath = Get-ProjectRelativePath -Path $File.FullName
                if (Test-FingerprintFile -RelativePath $RelativePath) {
                    $RelativePaths.Add($RelativePath)
                }
            }
        }
    }

    $SortedPaths = $RelativePaths.ToArray()
    [System.Array]::Sort($SortedPaths, [System.StringComparer]::Ordinal)

    $Lines = New-Object System.Collections.Generic.List[string]
    foreach ($RelativePath in $SortedPaths) {
        $FullPath = Join-Path $ProjectRoot $RelativePath
        $Hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $FullPath).Hash.ToLowerInvariant()
        $Lines.Add("$RelativePath`t$Hash")
    }

    $Payload = ([string]::Join("`n", $Lines.ToArray())) + "`n"
    $Bytes = [System.Text.Encoding]::UTF8.GetBytes($Payload)
    $Sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $HashBytes = $Sha256.ComputeHash($Bytes)
        return ([System.BitConverter]::ToString($HashBytes)).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $Sha256.Dispose()
    }
}

function Get-DockerImageLabel {
    param(
        [string]$ImageName,
        [string]$LabelName
    )

    $Format = '{{ index .Config.Labels "' + $LabelName + '" }}'

    if ($UseWslDocker) {
        $Result = Invoke-NativeCapture -FilePath "wsl.exe" -ArgumentList @("--exec", "docker", "image", "inspect", "--format", $Format, $ImageName)
    }
    else {
        $Result = Invoke-NativeCapture -FilePath "docker" -ArgumentList @("image", "inspect", "--format", $Format, $ImageName)
    }

    if ($Result.ExitCode -ne 0) {
        return $null
    }

    $Value = $Result.Output.Trim()
    if (-not $Value -or $Value -eq "<no value>") {
        return $null
    }

    return $Value
}

function Test-DockerImageCurrent {
    param(
        [string]$ImageName,
        [string]$SourceFingerprint
    )

    $ImageFingerprint = Get-DockerImageLabel -ImageName $ImageName -LabelName $FingerprintLabel
    return $ImageFingerprint -eq $SourceFingerprint
}

$SourceFingerprint = Get-SourceFingerprint
if (-not (Test-DockerImageCurrent -ImageName $Image -SourceFingerprint $SourceFingerprint)) {
    $BuildContext = Convert-ToDockerPath -Path $ProjectRoot
    Write-Host "Building Docker image $Image for source fingerprint $SourceFingerprint."
    Invoke-DockerChecked -DockerArgs @("build", "--build-arg", "EXCEL2CSV_IMAGE_FINGERPRINT=$SourceFingerprint", "-t", $Image, $BuildContext)
}

$OutputMount = Convert-ToDockerPath -Path $OutputDirPath
$DockerArgs = @("run", "--rm")
if ($ShouldAskPassword) {
    $DockerArgs += "-i"
    if (-not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected) {
        $DockerArgs += "-t"
    }
}
$DockerArgs += @("-v", "${OutputMount}:/output")
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

if ($ShouldAskPassword -and -not $Password -and -not $env:EXCEL2CSV_PASSWORD) {
    $DockerArgs += "--ask-password"
}

Invoke-DockerChecked -DockerArgs $DockerArgs
Write-Host "Wrote $OutputFullPath"
