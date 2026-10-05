# Excel VBA Automation

A collection of VBA modules developed to automate and manage Excel-based business workflows.

## Overview

This project contains VBA automation for managing structured Excel workbooks with multiple worksheets, automated worksheet protection, cell locking, and workbook event handling.

The automation is designed to reduce repetitive manual operations and maintain consistent workbook behavior across multiple sheets.

## Features

* Automatic worksheet protection when the workbook is opened
* Automatic locking of populated cells before saving
* Multi-worksheet automation
* Handling of merged cells
* Workbook event handling
* Error handling
* Excel application state management

## Structure

```text
excel-vba-automation/
│
├── README.md
│
└── src/
    ├── Workbook_Open.bas
    └── Workbook_BeforeSave.bas
```

## Modules

### Workbook_Open.bas

Handles actions performed when the workbook is opened.

* Iterates through all worksheets
* Applies worksheet protection
* Uses `UserInterfaceOnly` protection so VBA automation can continue modifying protected worksheets
* Temporarily disables screen updating during the operation

### Workbook_BeforeSave.bas

Handles automatic cell locking before the workbook is saved.

* Iterates through all worksheets
* Checks predefined cell ranges
* Locks populated cells
* Skips merged cells
* Temporarily disables Excel events and screen updating
* Includes error handling and user notification

## Technologies

* Microsoft Excel
* VBA (Visual Basic for Applications)
