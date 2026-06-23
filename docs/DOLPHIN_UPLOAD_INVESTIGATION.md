# Dolphin Upload Investigation

## Problem

Generated XLSX files were rejected by the Dolphin portal.

## Investigation Summary

Tested:

* Template workbook approach
* Formatting replication
* Header validation
* Row cleanup
* Workbook structure comparison

## Findings

OpenPyXL generated XLSX files without:

xl/sharedStrings.xml

Microsoft Excel and XlsxWriter generated XLSX files with:

xl/sharedStrings.xml

The Dolphin portal accepted files containing sharedStrings.xml.

## Resolution

Switched Dolphin XLSX generation from:

engine="openpyxl"

to:

engine="xlsxwriter"

## Result

* Dolphin upload works
* Batch Dolphin upload works
* Manual Excel save step eliminated

## Date Resolved

21-Jun-2026
