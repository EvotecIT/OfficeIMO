---
title: Protect Sheets and Workbooks
description: Protect Excel worksheets and workbook structure, and encrypt workbooks with a password to open, using OfficeIMO.Excel and PSWriteOffice.
order: 29
packages:
  - OfficeIMO.Excel
cmdlets:
  - Protect-OfficeDocument
limits: "In PowerShell, `Get-OfficeDocumentCapability -Protection` lists the protection operations each format supports."
search_terms:
  - protect workbook
  - password excel
  - lock sheet
---

Excel has two different kinds of protection, and they solve different problems:

- **Sheet and workbook protection** stops accidental edits in Excel: locked cells, inserting or deleting sheets. It is not security; the data remains readable to anyone who opens the file.
- **Encryption** (password to open) encrypts the whole package. Without the password the file cannot be read.

## Protect a worksheet

`Protect` locks the sheet's cells. Each `Allow*` option keeps one operation available to users. Remove protection again with `sheet.Unprotect()` or `Unprotect-OfficeExcelSheet`.

### C#

```csharp
using OfficeIMO.Excel;

using var book = ExcelDocument.Create("Directory.xlsx");
ExcelSheet sheet = book.AddSheet("Directory");
sheet.AddTable(new[] {
    new { Name = "Adele Vance", Department = "Finance" },
    new { Name = "Alex Wilber", Department = "Sales" }
});

sheet.Protect(new ExcelSheetProtectionOptions {
    AllowSort = true,
    AllowAutoFilter = true,
    AllowFormatColumns = true
});

book.Save();
```

### PowerShell

```powershell
Open-OfficeExcel -Path .\HR\Staff-Directory.xlsx {
    ExcelSheet 'Directory' {
        Protect-OfficeExcelSheet -AllowSort -AllowAutoFilter -AllowFormatColumns
    }
}
```

## Protect the workbook structure

Structure protection prevents adding, deleting, renaming, hiding, or moving sheets.

### C#

```csharp
using OfficeIMO.Excel;

using var book = ExcelDocument.Load("Directory.xlsx");
book.ProtectWorkbook(new ExcelWorkbookProtectionOptions { ProtectStructure = true });
book.Save();
```

## Encrypt a workbook with a password to open

Pass a password in `ExcelSaveOptions`. Load the file again with the same password in `ExcelLoadOptions`.

### C#

```csharp
using OfficeIMO.Excel;

string password = Environment.GetEnvironmentVariable("PAYROLL_PASSWORD")!;

using (var book = ExcelDocument.Create("Payroll.xlsx")) {
    book.AddSheet("Payroll").Cell("A1").SetValue("Confidential");
    book.Save("Payroll.xlsx", new ExcelSaveOptions { Password = password });
}

using var reopened = ExcelDocument.Load("Payroll.xlsx", new ExcelLoadOptions { Password = password });
Console.WriteLine(reopened["Payroll"].Cell("A1").GetValue());
```

To remove encryption, save with `new ExcelSaveOptions { Password = null }`. An unset password keeps the encryption the workbook was loaded with.

### PowerShell

```powershell
Get-ChildItem .\Payroll -Filter *.xlsx |
    Protect-OfficeDocument -Encrypt -Password (Read-Host -AsSecureString) -PassThru
```

`Protect-OfficeDocument -Restriction ReadOnlyRecommended` marks a workbook as read-only recommended without encrypting it.
