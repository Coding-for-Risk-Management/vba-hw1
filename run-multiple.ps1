param (
    [Parameter(Mandatory = $true)]
    [string[]]$MacroNames,
    [Parameter(Mandatory = $true)]
    [string]$BasFile
)

$failures = @()

# Check if BasFile exists
if (-Not (Test-Path $BasFile)) {
    Write-Host "BasFile path not found: $BasFile"
    exit 1
}

$basItem = Get-Item $BasFile
if ($basItem.PSIsContainer) {
    # Folder provided: run for each .bas file in the folder
    $basFiles = Get-ChildItem -Path $BasFile -Filter "*.bas"
    if ($basFiles.Count -eq 0) {
        Write-Host "No .bas files found in folder: $BasFile"
        exit 1
    }
    foreach ($bas in $basFiles) {
        Write-Host ""
        Write-Host ""
        Write-Host "Using VBA file: $($bas.FullName)"
        Write-Host "==== Starting Macro Execution ===="
        foreach ($macro in $MacroNames) {
            Write-Host "Running macro: $macro with file $($bas.Name)"
            & "$PSScriptRoot\excel-testing.ps1" -MacroName $macro -BasFile $bas.FullName
            if ($LASTEXITCODE -ne 0) {
                Write-Host "Macro $macro FAILED with file $($bas.Name)."
                $failures += "[$($bas.Name)] $macro"
            }
            else {
                Write-Host "Macro $macro executed successfully with file $($bas.Name)."
            }
            Write-Host "----------------------------------"
        }
    }
}
else {
    Write-Host ""
    Write-Host ""
    Write-Host "==== Starting Macro Execution ===="
    # Single file case
    foreach ($macro in $MacroNames) {
        Write-Host "Running macro: $macro with file $BasFile"
        & "$PSScriptRoot\excel-testing.ps1" -MacroName $macro -BasFile $BasFile
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Macro $macro FAILED."
            $failures += $macro
        }
        else {
            Write-Host "Macro $macro executed successfully."
        }
        Write-Host "----------------------------------"
    }
}

if ($failures.Count -gt 0) {
    Write-Host "Some macros failed:"
    $failures | ForEach-Object { Write-Host $_ }
    exit 1
}
else {
    Write-Host "All macros executed successfully."
}
