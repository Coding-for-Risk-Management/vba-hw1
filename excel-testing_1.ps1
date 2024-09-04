<#
.SYNOPSIS
This script automates testing in Excel by running a VBA macro and checking the result.

.DESCRIPTION
The script performs the following steps:
1. Creates a new instance of Excel application.
2. Opens the specified Excel file.
3. Imports a VBA script into the Excel file.
4. Runs a specified macro in the Excel file.
6. Checks the result of the macro execution.
7. Removes the imported VBA module.
8. Closes the workbook and quits Excel application.

.PARAMETER FilePath
The file path of the PowerShell script.

.EXAMPLE
.\excel-testing.ps1
Runs the script to automate testing in Excel.

.NOTES
- This script requires Excel to be installed on the system.
- The file paths for the Excel file and VBA script should be provided.
- The macro name and the expected result can be customized.
#>

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

# # Add a delay to let Excel finish loading
# Write-Host "Waiting for Excel to load..."
# Start-Sleep -Seconds 30

Write-Host "Excel Ready: $($excel.Ready)"

# Find a file that starts with "hw1" and ends with ".bas"
$files = Get-ChildItem -Path $currentDirectory -Filter "hw1*.bas"

# Check if any file found
if ($files.Count -gt 0) {
    $vbaScriptFileName = $files[0].Name
    $fullPath = Join-Path -Path $currentDirectory -ChildPath $vbaScriptFileName
    Write-Host "Importing VBA script: $fullPath"
    $module = $workbook.VBProject.VBComponents.Import($fullPath)
} else {
    Write-Host "No file found that starts with 'hw1' and ends with '.bas'"
}

$moduleName = $module.Name
Write-Host "Module Name: $moduleName"

# Define a list of macro names and their points
$macros = @{
    "test_PriceBond" = 3
    "test_FizzBuzz" = 3
    "test_MyMatMult" = 4
}

# Calculate the total maximum points
$totalMaxPoints = 0
foreach ($macro in $macros.GetEnumerator()) {
    $totalMaxPoints += $macro.Value
}

# Initialize total points
$totalPoints = 0

# Loop through each macro in the hashtable
foreach ($macro in $macros.GetEnumerator()) {
    Write-Host "Running macro: $($macro.Name)"

    try {
        $result = $excel.Run($macro.Name)
        Write-Host "$($macro.Name) $result"

        # If the macro didn't fail, add its points to the total
        if ($result -ne "FAIL") {
            $totalPoints += $macro.Value
            Write-Host "Points: $($macro.Value) of $($macro.Value)"
        }
        else {
            Write-Host "Points: 0 of $($macro.Value)"
        }
    } catch {
        Write-Host "Error running macro: $($macro.Name)"
    }
}

Write-Host "TOTAL POINTS: $totalPoints"


$workbook.VBProject.VBComponents.Remove($module)

$workbook.Close($False)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

if ($totalPoints -lt $totalMaxPoints) {
    exit 1
} else {
    exit 0
}