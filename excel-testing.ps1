# excel-testing.ps1

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

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\Excel\Security" -Name AccessVBOM -Value 1 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\Excel\Security" -Name VBAWarnings -Value 1 -Force | Out-Null

$excel.Visible = $False # Run in background
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$excel.EnableEvents = $false

# Set security level to low to enable macros
$excel.AutomationSecurity = 1

$currentDirectory = $PSScriptRoot

# Open Excel file
$fileName = "C4RM_Class2_UnitTests.xlsm"
$fullPath = Join-Path -Path $currentDirectory -ChildPath $fileName
Write-Host "Opening Excel file: $fullPath"

try {
    $workbook = $excel.Workbooks.Open($fullPath)

    # Disable Protected View
    $workbook.CheckCompatibility = $False

    Write-Host "Excel Ready: $($excel.Ready)"

    # Find a file that starts with "hw1" and ends with ".bas"
    $files = Get-ChildItem -Path $currentDirectory -Filter "hw1*.bas"

    if ($files.Count -gt 0) {
        $vbaScriptFileName = $files[0].Name
        $vbaFullPath = Join-Path -Path $currentDirectory -ChildPath $vbaScriptFileName
        Write-Host "Importing VBA script: $vbaFullPath"
        $module = $workbook.VBProject.VBComponents.Import($vbaFullPath)
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
    Write-Host $_.Exception.Message
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
