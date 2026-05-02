$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$zipPath = Join-Path $scriptDir "bin\Debug\CardOrderAgentSolution.zip"
$controlFolder = Join-Path $scriptDir "src\Controls\sa_Hermes.Controls.CardOrderAgentPcfLast66"

if (-not (Test-Path $zipPath)) {
    throw "Solution zip not found: $zipPath"
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$zip = [System.IO.Compression.ZipFile]::Open($zipPath, "Update")
try {
    # Remove old control entries if any
    $toDelete = @($zip.Entries | Where-Object { $_.FullName.StartsWith("Controls/", [System.StringComparison]::OrdinalIgnoreCase) })
    foreach ($entry in $toDelete) {
        $entry.Delete()
    }

    # Add bundle.js
    $bundlePath = Join-Path $controlFolder "bundle.js"
    if (Test-Path $bundlePath) {
        $entry = $zip.GetEntry("Controls/sa_Hermes.Controls.CardOrderAgentPcfLast66/bundle.js")
        if ($entry) { $entry.Delete() }
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $bundlePath, "Controls/sa_Hermes.Controls.CardOrderAgentPcfLast66/bundle.js") | Out-Null
        Write-Host "Added bundle.js"
    }

    # Add ControlManifest.xml
    $manifestPath = Join-Path $controlFolder "ControlManifest.xml"
    if (Test-Path $manifestPath) {
        $entry = $zip.GetEntry("Controls/sa_Hermes.Controls.CardOrderAgentPcfLast66/ControlManifest.xml")
        if ($entry) { $entry.Delete() }
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $manifestPath, "Controls/sa_Hermes.Controls.CardOrderAgentPcfLast66/ControlManifest.xml") | Out-Null
        Write-Host "Added ControlManifest.xml"
    }

    Write-Host "Control files added to solution successfully"
}
finally {
    $zip.Dispose()
}
