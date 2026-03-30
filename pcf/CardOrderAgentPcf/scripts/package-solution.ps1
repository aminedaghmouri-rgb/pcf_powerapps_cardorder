Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectDir = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $projectDir)

$bundlePath = Join-Path $projectDir "out\controls\CardOrderAgentPcf\bundle.js"
$controlManifestPath = Join-Path $projectDir "out\controls\CardOrderAgentPcf\ControlManifest.xml"
$customizationsXmlPath = Join-Path $repoRoot "solution\CardOrderAgentSolution\src\Other\Customizations.xml"
$solutionXmlPath = Join-Path $repoRoot "solution\CardOrderAgentSolution\src\Other\Solution.xml"

[xml]$controlManifestXml = Get-Content -Path $controlManifestPath
$controlNamespace = $controlManifestXml.manifest.control.namespace
$controlConstructor = $controlManifestXml.manifest.control.constructor
if ([string]::IsNullOrWhiteSpace($controlNamespace) -or [string]::IsNullOrWhiteSpace($controlConstructor)) {
    throw "Unable to resolve control namespace/constructor from ControlManifest.xml"
}
$controlZipFolder = "cth_{0}.{1}" -f $controlNamespace, $controlConstructor

$zipPaths = @(
    (Join-Path $repoRoot "solution\CardOrderAgentSolution\bin\Debug\CardOrderAgentSolution.zip"),
    (Join-Path $repoRoot "solution\CardOrderAgentSolution\bin\Debug\CardOrderAgentSolution_managed.zip")
)

$requiredFiles = @($bundlePath, $controlManifestPath, $customizationsXmlPath, $solutionXmlPath)
foreach ($required in $requiredFiles) {
    if (-not (Test-Path $required)) {
        throw "Required file not found: $required"
    }
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Update-ZipEntry {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip,
        [Parameter(Mandatory = $true)]
        [string]$EntryPath,
        [Parameter(Mandatory = $true)]
        [string]$SourceFile
    )

    $existing = $Zip.GetEntry($EntryPath)
    if ($existing) {
        $existing.Delete()
    }

    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Zip, $SourceFile, $EntryPath) | Out-Null
}

function Remove-ZipEntriesByPrefix {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip,
        [Parameter(Mandatory = $true)]
        [string]$EntryPrefix
    )

    $toDelete = @($Zip.Entries | Where-Object { $_.FullName.StartsWith($EntryPrefix, [System.StringComparison]::OrdinalIgnoreCase) })
    foreach ($entry in $toDelete) {
        $entry.Delete()
    }
}

foreach ($zipPath in $zipPaths) {
    if (-not (Test-Path $zipPath)) {
        Write-Warning "Skipping missing zip: $zipPath"
        continue
    }

    $zip = [System.IO.Compression.ZipFile]::Open($zipPath, "Update")
    try {
        # Keep only the current control payload in the zip to avoid stale Last/LastX entries.
        Remove-ZipEntriesByPrefix -Zip $zip -EntryPrefix "Controls/cth_Hermes.Controls.CardOrderAgentPcf"

        Update-ZipEntry -Zip $zip -EntryPath "Controls/$controlZipFolder/bundle.js" -SourceFile $bundlePath
        Update-ZipEntry -Zip $zip -EntryPath "Controls/$controlZipFolder/ControlManifest.xml" -SourceFile $controlManifestPath
        Update-ZipEntry -Zip $zip -EntryPath "customizations.xml" -SourceFile $customizationsXmlPath
        Update-ZipEntry -Zip $zip -EntryPath "solution.xml" -SourceFile $solutionXmlPath

        $bundleSize = (Get-Item $bundlePath).Length
        $manifestSize = (Get-Item $controlManifestPath).Length
        $customizationsSize = (Get-Item $customizationsXmlPath).Length
        $solutionSize = (Get-Item $solutionXmlPath).Length

        Write-Host "Updated: $zipPath"
        Write-Host "  bundle.js: $bundleSize bytes"
        Write-Host "  control manifest: $manifestSize bytes"
        Write-Host "  customizations.xml: $customizationsSize bytes"
        Write-Host "  solution.xml: $solutionSize bytes"
    }
    finally {
        $zip.Dispose()
    }
}

Write-Host "Packaging complete."

