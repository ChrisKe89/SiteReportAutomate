param(
    [string]$ExamplePath = ".env.example",
    [string]$EnvPath = ".env",
    [switch]$ForceCopy,     # Overwrite .env entirely with .env.example
    [switch]$DryRun         # Show what would change without writing
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Read-EnvFile {
    param([string]$Path)
    $map = [ordered]@{}
    $lines = @()
    if (-not (Test-Path -LiteralPath $Path)) {
        return @{ Map = $map; Lines = $lines }
    }

    # IMPORTANT: -split is an operator, so use (...) -split 'regex'
    # Use a regex that handles both LF and CRLF
    $content = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $lines = ($content) -split '\r?\n'

    $regex = '^\s*([A-Za-z_][A-Za-z0-9_\.]*)\s*=\s*(.*)$'  # KEY=VALUE
    foreach ($line in $lines) {
        if ($line -match '^\s*#' -or $line -match '^\s*$') { continue }
        if ($line -match $regex) {
            $key = $Matches[1]
            if (-not $map.Contains($key)) {
                # Preserve the original full line to keep quotes/spacing
                $map[$key] = $line
            }
        }
    }
    return @{ Map = $map; Lines = $lines }
}

if (-not (Test-Path -LiteralPath $ExamplePath)) {
    throw "Example env file not found at '$ExamplePath'."
}

# If .env doesn't exist and not forcing, copy example → .env
if (-not (Test-Path -LiteralPath $EnvPath) -and -not $ForceCopy) {
    if ($DryRun) {
        Write-Host "[DRY RUN] Would copy '$ExamplePath' -> '$EnvPath'"
        exit 0
    }
    Copy-Item -LiteralPath $ExamplePath -Destination $EnvPath -Force
    Write-Host "Created '$EnvPath' from '$ExamplePath'."
    exit 0
}

# Force overwrite mode
if ($ForceCopy) {
    $backup = "$EnvPath.bak.$((Get-Date).ToString('yyyyMMdd-HHmmss'))"
    if (Test-Path -LiteralPath $EnvPath) {
        if ($DryRun) {
            Write-Host "[DRY RUN] Would back up '$EnvPath' -> '$backup' and overwrite with example."
            exit 0
        }
        Copy-Item -LiteralPath $EnvPath -Destination $backup -Force
        Write-Host "Backed up current env to '$backup'."
    }
    Copy-Item -LiteralPath $ExamplePath -Destination $EnvPath -Force
    Write-Host "Overwrote '$EnvPath' with '$ExamplePath'."
    exit 0
}

# Merge: keep existing .env values, append any missing keys from example
$envData = Read-EnvFile -Path $EnvPath
$exData = Read-EnvFile -Path $ExamplePath

$envKeys = [System.Collections.Generic.HashSet[string]]::new($envData.Map.Keys, [System.StringComparer]::Ordinal)
$missing = @()
foreach ($k in $exData.Map.Keys) {
    if (-not $envKeys.Contains($k)) { $missing += $k }
}

if ($missing.Count -eq 0) {
    Write-Host "No changes needed. '$EnvPath' already contains all keys from '$ExamplePath'."
    exit 0
}

Write-Host "Missing keys to append: $($missing -join ', ')"

if ($DryRun) {
    Write-Host "[DRY RUN] Would append the above keys to '$EnvPath'."
    exit 0
}

$appendBlock = @()
$appendBlock += ""
$appendBlock += "# --- Added from $ExamplePath on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ---"
foreach ($k in $missing) {
    $appendBlock += $exData.Map[$k]
}

Add-Content -LiteralPath $EnvPath -Value ($appendBlock -join [Environment]::NewLine) -Encoding UTF8
Write-Host "Appended $($missing.Count) key(s) to '$EnvPath'. Done."
# End of script