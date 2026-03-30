param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PacArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not $PSBoundParameters.ContainsKey("PacArgs")) {
    $PacArgs = @()
}

$customPacPath = $env:PPCLI_PATH
if ($customPacPath -and (Test-Path $customPacPath)) {
    $pacExe = $customPacPath
}
else {
    $searchRoot = Join-Path $env:LOCALAPPDATA "Microsoft\PowerAppsCLI"
    $pacCandidates = @()
    if (Test-Path $searchRoot) {
        $pacCandidates = @(Get-ChildItem -Path $searchRoot -Recurse -Filter pac.exe -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending)
    }

    if ($pacCandidates.Count -eq 0) {
        throw "Power Platform CLI not found. Install it from https://aka.ms/PowerAppsCLI"
    }

    $pacExe = $pacCandidates[0].FullName
}

Write-Host "Using pac: $pacExe"
& $pacExe @PacArgs
exit $LASTEXITCODE
