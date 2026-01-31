# SharePoint Online – Fast CSV Import Script

A reusable PowerShell script for importing large CSV datasets into SharePoint Online lists using PnP PowerShell and CSOM batching.

Built for real-world scenarios where standard imports become slow or unreliable.

---

## Why this script

When working with large SharePoint lists, common import methods struggle with:

- Performance on large datasets  
- Throttling issues  
- Choice and multi-choice fields  
- Re-running imports during testing or migration  

This script solves those problems by using batch processing, dynamic column mapping, and safe throttling handling.

---

## What it does

- Imports CSV data into an existing SharePoint Online list  
- Supports 10k–100k+ records  
- Maps CSV headers to SharePoint internal field names  
- Validates Choice and Multi-Choice fields  
- Supports test runs before full execution  
- Optionally preserves Created / Modified dates  

This script focuses only on data import. The target list must already exist.

---

## Key features

- High-performance CSOM batching  
- Dynamic CSV → SharePoint column mapping  
- Choice & Multi-Choice field support  
- Throttling protection  
- Test mode for safe validation  
- Reusable across projects  

---

## Basic usage

```powershell
$ColumnMap = @{
  "Employee Name" = "Title"
  "Status"        = "Status"
  "Tags"          = "Tags"
}

.\Import-CSVToSharePoint.ps1 `
  -SiteUrl "https://tenant.sharepoint.com/sites/TargetSite" `
  -ListName "Employees" `
  -CsvPath "employees.csv" `
  -ClientId "YOUR-CLIENT-ID" `
  -TenantName "tenant.onmicrosoft.com" `
  -ColumnMap $ColumnMap `
  -TestRun
```

---

## Recommended approach

1. Run with `-TestRun`  
2. Verify data and mappings  
3. Remove `-TestRun` for full import  
4. Monitor logs during execution  

---

## Notes

- Target list and columns must already exist  
- Invalid choice values are logged and skipped  
- Designed for repeated use in production scenarios  
