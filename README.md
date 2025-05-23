# VBA Homework 1 - Automated Testing

This project provides automated testing for the first VBA assignment in Coding for Risk Management using GitHub Actions.

## Student Instructions

### Submitting Your Work

1. **Clone this repository** to your local machine
2. **Create your VBA file** following the naming convention: `hw1_xx1234.bas` where `xx1234` is your UNI (e.g., `hw1_ab1234.bas`)
3. **Add your VBA file** to the root directory of the repository
4. **Commit and push** your changes to GitHub:
   ```bash
   git add hw1_xx1234.bas
   git commit -m "Submit homework 1"
   git push origin main
   ```
5. **Check your results** in the "Actions" tab of your GitHub repository

### Understanding Test Results

Your code will be automatically tested for three functions:

- **PriceBond Test** (3 points) - Tests your `PriceBond` function
- **FizzBuzz Test** (4 points) - Tests your `FizzBuzz` function
- **MyMatMult Test** (3 points) - Tests your `MyMatMult` function

Each test will show as ✅ **PASS** or ❌ **FAIL** in the Actions tab. Click on any failed test to see detailed error messages.

### Running Tests Locally

You can test your code locally before submitting using PowerShell:

#### Prerequisites

- Windows computer with Excel installed
- PowerShell execution policy set to allow scripts

#### Running Individual Tests

```powershell
# Test a specific function (replace with your actual filename)
powershell -ExecutionPolicy Bypass -File excel-testing.ps1 -MacroName test_PriceBond
powershell -ExecutionPolicy Bypass -File excel-testing.ps1 -MacroName test_FizzBuzz
powershell -ExecutionPolicy Bypass -File excel-testing.ps1 -MacroName test_MyMatMult
```

#### Running All Tests

```powershell
# Run all tests for your file
powershell -ExecutionPolicy Bypass -File run-multiple.ps1 -MacroNames @("test_PriceBond","test_FizzBuzz","test_MyMatMult") -BasFile "hw1_xx1234.bas"
```

#### Setup for Local Testing

1. **Enable macros** in Excel security settings
2. **Set PowerShell execution policy** (run as administrator):
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

### File Requirements

Your VBA file must contain the following functions with exact names:

- `PriceBond()` - Bond pricing function
- `FizzBuzz()` - FizzBuzz implementation
- `MyMatMult()` - Matrix multiplication function

Make sure your functions return the expected data types and handle the test cases correctly.

## Technical Details

### How the Testing System Works

1. **GitHub Workflow**: Triggered automatically on every push
2. **Windows VM Setup**: GitHub Actions spins up a Windows virtual machine
3. **Office Installation**: Installs Office 365 Business via Chocolatey
4. **Test Execution**:
   - Opens `C4RM_Class2_UnitTests.xlsm` Excel file
   - Imports your VBA file as a module
   - Runs each test macro and captures results
   - Reports PASS/FAIL for each function

### Files in This Repository

- `excel-testing.ps1` - Main PowerShell script for running individual tests
- `run-multiple.ps1` - Script for running multiple tests locally
- `unit_tests_1.bas` - Contains the test functions that validate your code
- `C4RM_Class2_UnitTests.xlsm` - Excel workbook with test framework
- `.github/workflows/classroom.yml` - GitHub Actions workflow configuration
- `.github/classroom/autograding.json` - Test configuration and point values

## Troubleshooting

### Common Issues

1. **File naming**: Ensure your file follows `hw1_xx1234.bas` format exactly
2. **Function names**: Your functions must match the expected names exactly
3. **Compilation errors**: Fix any VBA syntax errors before submitting
4. **Return values**: Ensure your functions return the correct data types

### Getting Help

- Check the Actions tab for detailed error messages
- Review the test functions in `unit_tests_1.bas` to understand expected behavior
- Test locally before submitting to GitHub
