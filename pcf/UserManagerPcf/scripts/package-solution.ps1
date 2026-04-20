Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectDir = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $projectDir)

# Discover all built controls
$controlsDir = Join-Path $projectDir "out\controls"
$controlDirs = Get-ChildItem -Path $controlsDir -Directory

$controls = @()
foreach ($controlDir in $controlDirs) {
    $bundlePath = Join-Path $controlDir.FullName "bundle.js"
    $manifestPath = Join-Path $controlDir.FullName "ControlManifest.xml"
    
    if ((Test-Path $bundlePath) -and (Test-Path $manifestPath)) {
        [xml]$manifestXml = Get-Content -Path $manifestPath
        $namespace = $manifestXml.manifest.control.namespace
        $constructor = $manifestXml.manifest.control.constructor
        
        if (-not [string]::IsNullOrWhiteSpace($namespace) -and -not [string]::IsNullOrWhiteSpace($constructor)) {
            $zipFolder = "sa_{0}.{1}" -f $namespace, $constructor
            $controls += @{
                Name = $controlDir.Name
                BundlePath = $bundlePath
                ManifestPath = $manifestPath
                ZipFolder = $zipFolder
            }
            Write-Host "Found control: $($controlDir.Name) -> $zipFolder"
        }
    }
}

if ($controls.Count -eq 0) {
    throw "No valid controls found in $controlsDir"
}

$customizationsXmlPath = Join-Path $repoRoot "solution\UserManagerSolution\src\Other\Customizations.xml"
$solutionXmlPath = Join-Path $repoRoot "solution\UserManagerSolution\src\Other\Solution.xml"

$zipPaths = @(
    (Join-Path $repoRoot "solution\UserManagerSolution\bin\Debug\UserManagerSolution.zip"),
    (Join-Path $repoRoot "solution\UserManagerSolution\bin\Debug\UserManagerSolution_managed.zip")
)

# Generate [Content_Types].xml dynamically with all controls
$overrides = ($controls | ForEach-Object { '<Override PartName="/Controls/{0}/ControlManifest.xml" ContentType="application/octet-stream" />' -f $_.ZipFolder }) -join ''
$contentTypesXml = @"
<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="text/xml" /><Default Extension="js" ContentType="application/octet-stream" />$overrides</Types>
"@
$contentTypesTmp = [System.IO.Path]::GetTempFileName()
[System.IO.File]::WriteAllText($contentTypesTmp, $contentTypesXml.Trim(), [System.Text.Encoding]::UTF8)

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
        # Remove old control entries
        Remove-ZipEntriesByPrefix -Zip $zip -EntryPrefix "Controls/sa_Hermes.Controls.UserManagePcf1"
        Remove-ZipEntriesByPrefix -Zip $zip -EntryPrefix "Controls/sa_Hermes.Controls8.UserManagerPcf8"
        Remove-ZipEntriesByPrefix -Zip $zip -EntryPrefix "Controls/sa_Hermes.Controls8.UserManagerPcf"

        # Update common files
        Update-ZipEntry -Zip $zip -EntryPath "[Content_Types].xml" -SourceFile $contentTypesTmp
        Update-ZipEntry -Zip $zip -EntryPath "customizations.xml" -SourceFile $customizationsXmlPath
        Update-ZipEntry -Zip $zip -EntryPath "solution.xml" -SourceFile $solutionXmlPath

        # Add all controls
        foreach ($control in $controls) {
            Update-ZipEntry -Zip $zip -EntryPath "Controls/$($control.ZipFolder)/bundle.js" -SourceFile $control.BundlePath
            Update-ZipEntry -Zip $zip -EntryPath "Controls/$($control.ZipFolder)/ControlManifest.xml" -SourceFile $control.ManifestPath
            
            $bundleSize = (Get-Item $control.BundlePath).Length
            Write-Host "  Added control: $($control.Name) -> $($control.ZipFolder) ($bundleSize bytes)"
        }

        Write-Host "Updated: $zipPath"
    }
    finally {
        $zip.Dispose()
    }
}

Write-Host "Packaging complete."

Remove-Item $contentTypesTmp -Force -ErrorAction SilentlyContinue
