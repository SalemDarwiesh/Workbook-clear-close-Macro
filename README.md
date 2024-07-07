# Workbook-clear-and-close-Macro

## Overview
This repository contains a VBA macro for Excel designed to automate the process of clearing the content of a specified sheet before closing the workbook. This script helps streamline repetitive tasks and ensures that sensitive or temporary data is cleared before the workbook is closed.

## Contents
- **Workbook BeforeClose Macro:** A script to clean the content of the "AgentPerformance Macro" sheet before closing the workbook if certain conditions are met.

## Usage
1. Download or clone the repository to your local machine.
2. Open the Excel workbook where you want to add the macro.
3. Press `Alt + F11` to open the VBA editor.
4. Copy the desired script from the repository and paste it into the VBA editor.
5. Customize the file paths and any other parameters as needed.
6. Save the workbook as a macro-enabled file (`.xlsm`).

## Example Code

### Workbook BeforeClose Macro
```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Select the "SHEETNAME" sheet
    Sheets("SHEETNAME").Select

    ' Check if there is no data in the range A1:AB1
    If WorksheetFunction.CountA(Range("A1:AB1")) = 0 Then
        Exit Sub
    Else
        ' Select all used cells in the sheet
        Cells.Select
        
        ' Clear the content of all selected cells
        Selection.ClearContents
    End If
End Sub
