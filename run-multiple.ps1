param (
    [Parameter(Mandatory = $true)]
    [string[]]$MacroNames,
    [Parameter(Mandatory = $true)]
    [string]$BasFile
)

Write-Host "==== Starting Macro Execution ===="

foreach ($macro in $MacroNames) {
    Write-Host "Running macro: $macro"
    & "$PSScriptRoot\excel-testing.ps1" -MacroName $macro -BasFile $BasFile
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Macro $macro failed."
        exit $LASTEXITCODE
    } else {
        Write-Host "Macro $macro executed successfully."
    }
    Write-Host "----------------------------------"
}

Write-Host "All macros executed successfully."
