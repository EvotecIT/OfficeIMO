---
title: "Fill and Flatten PDF Forms"
description: "List AcroForm fields, fill them by name, flatten the result, exchange XFDF data, and add form fields to new PDFs with OfficeIMO.Pdf and PSWriteOffice."
layout: docs
packages:
  - OfficeIMO.Pdf
cmdlets:
  - Set-OfficePdfFormField
limits: "On hybrid AcroForm/XFA forms the AcroForm fields are filled; XFA-aware viewers may keep showing the original XFA data. For signed or certified forms, `Forms.AppendRevision(values)` adds the values as a new revision after the original bytes instead of rewriting the file."
---

PDF forms (AcroForms) are named fields: text boxes, check boxes, choice lists, and radio groups. Fill them by field name. Filling keeps the fields editable; flattening turns their values into ordinary page content so they can no longer be changed.

## List the fields in a form

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocumentInfo info = PdfDocument.Load("Application.pdf").Inspect();
foreach (PdfFormField field in info.FormFields) {
    Console.WriteLine($"{field.Name} = {field.Value} ({field.Kind})");
}
```

### PowerShell

```powershell
Get-OfficePdfFormField -Path .\Application.pdf | Select-Object Name, FieldType, Value
```

## Fill a form

### C#

```csharp
using OfficeIMO.Pdf;

var values = new Dictionary<string, string> {
    ["Applicant.Name"] = "Adele Vance",
    ["Approval.Status"] = "Approved"
};

// Fill and keep the fields editable
PdfDocument.Load("Application.pdf").Forms.Fill(values).Save("Application-filled.pdf");

// Fill and flatten in one step
PdfDocument.Load("Application.pdf").Forms.FillAndFlatten(values).Save("Application-final.pdf");
```

### PowerShell

```powershell
Set-OfficePdfFormField -Path .\Forms\Access-Request.pdf -OutputPath .\Forms\Access-Request-Ada.pdf `
    -Value @{ Requester = 'Ada Lovelace'; System = 'ERP' }

Set-OfficePdfFormField -Path .\Forms\Access-Request.pdf -OutputPath .\Forms\Access-Request-Final.pdf `
    -Value @{ Requester = 'Ada Lovelace'; System = 'ERP' } -Flatten
```

## Flatten an already filled form

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocument.Load("Application-filled.pdf").Forms.Flatten().Save("Application-flat.pdf");
```

## Exchange form data as XFDF

XFDF is the XML form-data format used by Acrobat and many document workflows. Export the values from one PDF and import them into another copy of the same form.

### C#

```csharp
using OfficeIMO.Pdf;

string xfdf = PdfDocument.Load("Application-filled.pdf").Forms.ExportXfdf();
File.WriteAllText("Application.xfdf", xfdf);

PdfDocument.Load("Application.pdf")
    .Forms.ImportXfdf(File.ReadAllText("Application.xfdf"))
    .Save("Application-restored.pdf");
```

### PowerShell

```powershell
Export-OfficePdfXfdf -Path .\Application-filled.pdf -OutputPath .\Application.xfdf
Import-OfficePdfXfdf -Path .\Application.pdf -XfdfPath .\Application.xfdf -OutputPath .\Application-restored.pdf
```

## Add form fields to a new PDF

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Application")
        .AddParagraph("Applicant name")
        .TextField("Applicant.Name")
        .AddParagraph("Approval status")
        .TextField("Approval.Status")))
    .Save("Application.pdf");
```

### PowerShell

```powershell
New-OfficePdf -Path .\Review.pdf {
    PdfHeading 'Review'
    PdfFormField -Name Reviewer -Type Text -Value 'Unassigned' -Width 320 -Height 24
    PdfFormField -Name Decision -Type Choice -Option Approve,Reject,Defer -Value Defer
}
```
