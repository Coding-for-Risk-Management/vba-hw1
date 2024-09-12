<#
.SYNOPSIS
This script automates testing in Excel by running a specified VBA macro.

.DESCRIPTION
The script performs the following steps:
1. Creates a new instance of Excel application.
2. Opens the specified Excel file.
3. Imports a VBA script into the Excel file.
4. Runs the specified macro in the Excel file.
5. Removes the imported VBA module.
6. Closes the workbook and quits Excel application.

.PARAMETER FilePath
The file path of the PowerShell script.

.PARAMETER MacroName
The name of the macro to be executed.

.EXAMPLE
.\excel-testing.ps1 -MacroName "test_PriceBond"
Runs the script to automate testing in Excel for the specified macro.

.NOTES
- This script requires Excel to be installed on the system.
- The file paths for the Excel file and VBA script should be provided.
- The macro name can be customized.
#>

param (
    [string]$MacroName
)

if (-not $MacroName) {
    Write-Host "Please provide a macro name to run using the -MacroName parameter."
    exit 1
}

# Print the current directory
Write-Host "Current Directory: $PSScriptRoot"

$excel = New-Object -ComObject Excel.Application

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name AccessVBOM -Value 1 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name VBAWarnings -Value 1 -Force | Out-Null

$excel.Visible = $False # Set to $False to run in background
$Excel.DisplayAlerts = $false
$Excel.AskToUpdateLinks = $false
$Excel.EnableEvents = $false

# Set security level to low to enable macros
$excel.AutomationSecurity = 1

$currentDirectory = (Get-Location).Path

# Open Excel file
$fileName = "C4RM_Class2_UnitTests.xlsm"
$fullPath = Join-Path -Path $currentDirectory -ChildPath $fileName
Write-Host "Opening Excel file: $fullPath"
$workbook = $excel.Workbooks.Open($fullPath)

# Disable Protected View
$workbook.CheckCompatibility = $False

Write-Host "Excel Ready: $($excel.Ready)"

# Find a file that starts with "hw1" and ends with ".bas"
$files = Get-ChildItem -Path $currentDirectory -Filter "hw1*.bas"

try {
    # Check if any file found
    if ($files.Count -gt 0) {
        $vbaScriptFileName = $files[0].Name
        $fullPath = Join-Path -Path $currentDirectory -ChildPath $vbaScriptFileName
        Write-Host "Importing VBA script: $fullPath"
        $module = $workbook.VBProject.VBComponents.Import($fullPath)
    } else {
        Write-Host "No file found that starts with 'hw1' and ends with '.bas'"
        exit 1
    }

    $moduleName = $module.Name
    Write-Host "Module Name: $moduleName"

    # Run the specified macro
    Write-Host "Running macro: $MacroName"
    $result = $excel.Run($MacroName)
    Write-Host "$MacroName result: $result"

    if ($result -eq "FAIL") {
        Write-Host "Macro failed."
        exit 1
    }

} catch {
    Write-Host "Error running macro: $MacroName"
    exit 1

} finally {
    # Ensure that Excel closes even if an error occurs
    if ($module) {
        $workbook.VBProject.VBComponents.Remove($module)
    }

    # Close the workbook without saving changes and quit Excel
    if ($workbook) {
        $workbook.Close($False)
    }

    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    Write-Host "Excel instance closed successfully."
}

Write-Host "Macro executed successfully."
exit 0
