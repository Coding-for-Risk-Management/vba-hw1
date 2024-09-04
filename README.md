# vba-hw1

This project aims to automate testing for the first VBA class in Coding for Risk Management

## Instructions

1. Students clone repo
2. Students delete hw1_blank.bas
3. Students add their .bas file in the format of hw1_xx1234.bas where xx1234 is their UNI
4. Students commit and push updates
5. Students and teachers can check the success of their code in the actions section of github

## TODO

- Code will fail if there is a compile error with the code. Could separate each test into it's own github action to isolate the bad code
- Incorporate test results with github classroom

## How does this work

- This repo executes a github workflow
- The github workflow initaites a Windows VM and installs Office
- Next it will run the excel-testing powershell script
- The powershell script will load the xlsm file and load in the students homework as a module
- The powershell will then execute the macros listed in the script and print out the final score for the student
