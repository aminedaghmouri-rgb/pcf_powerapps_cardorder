$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcFolder = Join-Path $scriptDir "src"
$zipPath = Join-Path $scriptDir "bin\CardOrderAgentSolution.zip"

# Create bin folder
$binDir = Split-Path -Parent $zipPath
if (-not (Test-Path $binDir)) {
    New-Item -ItemType Directory -Path $binDir -Force | Out-Null
}

# Remove old zip if exists
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Create zip archive
$zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Create)

try {
    # Add root files from Other folder to root of zip
    $customizationsPath = Join-Path $srcFolder "Other\Customizations.xml"
    $solutionPath = Join-Path $srcFolder "Other\Solution.xml"
    
    if (Test-Path $customizationsPath) {
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $customizationsPath, "customizations.xml") | Out-Null
    }
    
    if (Test-Path $solutionPath) {
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $solutionPath, "solution.xml") | Out-Null
    }
    
    # Create [Content_Types].xml
    $contentTypesXml = @'
<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/octet-stream" />
  <Default Extension="js" ContentType="application/octet-stream" />
</Types>
'@
    $contentTypesEntry = $zip.CreateEntry("[Content_Types].xml")
    $writer = New-Object System.IO.StreamWriter($contentTypesEntry.Open())
    $writer.Write($contentTypesXml)
    $writer.Close()
    
    # Add Controls folder
    $controlsFolder = Join-Path $srcFolder "Controls"
    if (Test-Path $controlsFolder) {
        Get-ChildItem -Path $controlsFolder -Recurse -File | ForEach-Object {
            $relativePath = $_.FullName.Substring($srcFolder.Length + 1).Replace("\", "/")
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $relativePath) | Out-Null
        }
    }
    
    # Add ControlExtensions folder
    $controlExtFolder = Join-Path $srcFolder "ControlExtensions"
    if (Test-Path $controlExtFolder) {
        Get-ChildItem -Path $controlExtFolder -Recurse -File | ForEach-Object {
            $relativePath = $_.FullName.Substring($srcFolder.Length + 1).Replace("\", "/")
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $relativePath) | Out-Null
        }
    }
    
    Write-Host "Solution package created: $zipPath" -ForegroundColor Green
}
finally {
    $zip.Dispose()
}

# Verify contents
$zipRead = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
try {
    Write-Host "`nPackage contents:"
    $zipRead.Entries | ForEach-Object {
        Write-Host "  $($_.FullName)"
    }
}
finally {
    $zipRead.Dispose()
}
