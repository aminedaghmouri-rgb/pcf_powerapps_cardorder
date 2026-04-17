$file = "c:\Users\c_adaghm\GIT-Perso-Projects\pcf_powerapps_cardorder\pcf\CardOrderAgentPcf\CardOrderAgentPcf\index.ts"
$lines = [System.IO.File]::ReadAllLines($file)

# Find the line index of the FIRST occurrence of 'const hideItemCount'
# which is in createOrderCard (line ~911, 0-indexed ~910)
$targetIdx = -1
for ($i = 0; $i -lt $lines.Length; $i++) {
    if ($lines[$i] -match "const hideItemCount = this\.shouldHideItemCount" ) {
        $targetIdx = $i
        break
    }
}

if ($targetIdx -lt 0) { Write-Host "Not found!"; exit 1 }

# From targetIdx, find the line with 'record.createdTime' (should be ~7 lines later)
$createdTimeIdx = -1
for ($i = $targetIdx; $i -lt [Math]::Min($targetIdx + 15, $lines.Length); $i++) {
    if ($lines[$i] -match "record\.createdTime") {
        $createdTimeIdx = $i
        break
    }
}

if ($createdTimeIdx -lt 0) { Write-Host "createdTime line not found!"; exit 1 }

Write-Host "Found createdTime at line $($createdTimeIdx + 1)"
Write-Host "Line before: $($lines[$createdTimeIdx - 1])"
Write-Host "Line: $($lines[$createdTimeIdx])"
Write-Host "Line after: $($lines[$createdTimeIdx + 1])"

# Lines to replace: createdTimeIdx-1 (the span open), createdTimeIdx (fontWeight), createdTimeIdx+1 (createdTime close), createdTimeIdx+2 (details.appendChild)
# Replace those 4 lines with the conditional block
$newBlock = @(
    '        if (showActions && record.createdOn) {',
    '            meta.appendChild(this.createElement("span", {',
    '                fontWeight: "400"',
    '            }, record.createdOn.toLocaleDateString([], { day: "2-digit", month: "2-digit", year: "numeric" })));',
    '        } else {',
    '            meta.appendChild(this.createElement("span", {',
    '                fontWeight: "400"',
    '            }, record.createdTime));',
    '        }',
    '        details.appendChild(meta);'
)

$spanOpenIdx = $createdTimeIdx - 1  # 'meta.appendChild(this.createElement("span", {'
$detailsIdx = $createdTimeIdx + 2   # 'details.appendChild(meta);'

$result = $lines[0..($spanOpenIdx - 1)] + $newBlock + $lines[($detailsIdx + 1)..($lines.Length - 1)]
[System.IO.File]::WriteAllLines($file, $result, [System.Text.UTF8Encoding]::new($false))
Write-Host "Done! Replaced lines $($spanOpenIdx+1) to $($detailsIdx+1)"
