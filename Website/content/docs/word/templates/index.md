---
title: Generate Word Documents from DOCX Templates
description: Fill DOCX templates from C# data with {{placeholders}}, repeated and conditional blocks, images, and links, run classic MERGEFIELD mail merge, and fill content-control forms with OfficeIMO.Word.
order: 17
packages:
  - OfficeIMO.Word
cmdlets:
  - Invoke-OfficeWordMailMerge
  - Update-OfficeWordText
  - Add-OfficeWordContentControl
  - Get-OfficeWordContentControl
limits: Repeated and conditional block markers must sit in their own paragraphs; a marker inside running text is not expanded.
search_terms:
  - word template
  - mail merge
  - docx placeholders
  - generate docx from template
  - content controls
meta.seo_title: "Generate DOCX from Word templates in .NET | OfficeIMO"
verified: 2026-11-14
---

OfficeIMO.Word supports three template styles. Choose the one that matches who owns the document and how much structure the application needs.

| Template style | Use it when | Entry point |
|---|---|---|
| Plain `{{Name}}` placeholders | A designer owns a DOCX layout and application data should bind in one pass | `WordTemplate.Apply(...)` |
| Word `MERGEFIELD` fields | Existing mail-merge templates and Word field formatting are part of the contract | `WordMailMerge` |
| Tagged content controls | The document is a typed form with checkboxes, dates, choices, pictures, or repeating sections | `FillContentControlValues(...)` |

## Bind a model to ordinary placeholders

Write placeholders directly in Word:

```text
Service summary for {{Client.Name}}

{{#each Services}}
{{Name}} — {{Hours}} hours
{{#Priority}}
Priority delivery
{{/Priority}}
{{/each Services}}

Portal: {{Portal}}
Logo: {{Logo}}
```

Repeated and conditional markers must be placed on their own paragraphs. The content between them can contain normal paragraphs, styled runs, lists, tables, text boxes, and nested blocks. Scalar placeholders can be split across Word runs; the replacement inherits the formatting of the first run containing the marker.

The dictionary overload avoids reflection and is the recommended path for trimming and NativeAOT. `WordTemplateHyperlink` and `WordTemplateImage` values become a link or a picture in place of the placeholder.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("service-template.docx");

var values = new Dictionary<string, object?> {
    ["Client"] = new Dictionary<string, object?> {
        ["Name"] = "Northwind Traders"
    },
    ["Services"] = new object[] {
        new Dictionary<string, object?> {
            ["Name"] = "Assessment",
            ["Hours"] = 8,
            ["Priority"] = true
        },
        new Dictionary<string, object?> {
            ["Name"] = "Implementation",
            ["Hours"] = 24,
            ["Priority"] = false
        }
    },
    ["Portal"] = new WordTemplateHyperlink(
        "Open customer portal",
        new Uri("https://example.com/customer")),
    ["Logo"] = new WordTemplateImage(
        File.ReadAllBytes("logo.png"),
        "logo.png",
        width: OfficeLength.Pixels(96),
        height: OfficeLength.Pixels(32),
        altText: "Company logo")
};

WordTemplateResult result = WordTemplate.Apply(document, values).EnsureComplete();
document.Save("service-summary.docx");
```

`WordTemplate.Apply(document, model)` also accepts a POCO or anonymous object. That overload reflects over public properties and is annotated accordingly; use dictionaries when the application publishes with trimming or NativeAOT.

## Template syntax

| Syntax | Meaning |
|---|---|
| `{{Name}}` | Scalar value; `null` becomes empty text |
| `{{Customer.Name}}` | Nested dictionary key or public property |
| `{{this}}` | Current scalar item inside a repeated block |
| `{{#each Items}}` … `{{/each Items}}` | Repeat any enclosed Word blocks for an enumerable value |
| `{{#Approved}}` … `{{/Approved}}` | Include enclosed content when the Boolean value is `true` |

Nested blocks inherit their parent scope. An item can therefore use both `{{Name}}` from the current item and `{{Client.Name}}` from the outer model.

## Detect incomplete output

Missing scalar values are preserved in the DOCX by default and reported through `WordTemplateResult`. This makes a bad binding visible instead of silently turning it into blank content.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("service-template.docx");
var values = new Dictionary<string, object?> { ["Client"] = new Dictionary<string, object?> { ["Name"] = "Northwind" } };

WordTemplateResult result = WordTemplate.Apply(document, values);
if (!result.IsComplete) {
    Console.Error.WriteLine("Missing: " + string.Join(", ", result.MissingValueNames));
    return;
}
document.Save("service-summary.docx");
```

Set `WordTemplateOptions.RemoveMissingPlaceholders` only when blank output is an intentional application policy. Missing or non-Boolean block values and non-enumerable repeated values fail immediately because leaving half-expanded document structure would be ambiguous.

## Run a MERGEFIELD mail merge

`WordMailMerge.PreflightTemplate` inspects Word `MERGEFIELD` fields and reports conditional blocks, repeating blocks, malformed markers, and requested names that are missing. Run preflight when a template is uploaded or promoted, not only after a production batch has failed.

### C#

```csharp
using OfficeIMO.Word;

using (WordDocument template = WordDocument.Load("contract-template.docx")) {
    WordTemplatePreflightReport report = WordMailMerge.PreflightTemplate(
        template,
        mergeFieldNames: new[] { "CustomerName", "ContractNumber", "EffectiveDate" });
    if (!report.CanBindTemplate) throw new InvalidOperationException(report.ToJson());
}

var records = new List<IDictionary<string, string>> {
    new Dictionary<string, string> { ["CustomerName"] = "Northwind", ["ContractNumber"] = "C-1001", ["EffectiveDate"] = "2026-12-01" },
    new Dictionary<string, string> { ["CustomerName"] = "Contoso", ["ContractNumber"] = "C-1002", ["EffectiveDate"] = "2027-01-01" }
};
IReadOnlyList<string> written = WordMailMerge.ExecuteBatch("contract-template.docx", records,
    (index, record) => $"contract-{record["ContractNumber"]}.docx");
Console.WriteLine($"{written.Count} contracts written");
```

### PowerShell

```powershell
$document = Open-OfficeWord -Path .\letter-template.docx
Invoke-OfficeWordMailMerge -Document $document -InputObject @{ FirstName = 'John'; OrderId = 12345 }
Save-OfficeWord -Document $document -Path .\letter-john.docx
Close-OfficeWord -Document $document
```

`WordMailMerge.Execute` replaces fields in an already loaded document. `ExecuteBatch` loads the template for every record and writes one output per value set. Table-row and grouped-row helpers repeat structured table regions without rebuilding layout in code.

## Fill a content-control form

Tagged content controls make a typed form: text, drop-down, date, check box, picture, and repeating-section controls. `ExtractContentControlValues` reads them into a dictionary, `ValidateContentControlValues` reports missing, unused, duplicate, and invalid values, and `FillContentControlValues` writes values into the matching controls. Pick one key policy (tag, alias, or tag then alias) and use it for creating, validating, and filling.

### C#

```csharp
using OfficeIMO.Word;

using (WordDocument form = WordDocument.Create("request-form.docx")) {
    form.AddParagraph("Requester: ").AddStructuredDocumentTag("Name", alias: "Requester", tag: "Requester");
    form.AddParagraph("Approved: ").AddCheckBox(false, alias: "Approved", tag: "Approved");
    form.AddParagraph("Due: ").AddDatePicker(null, alias: "Due date", tag: "DueDate");
    form.Save();
}

using WordDocument document = WordDocument.Load("request-form.docx");
var values = new Dictionary<string, object?> {
    ["Requester"] = "Alice Smith",
    ["Approved"] = true,
    ["DueDate"] = new DateTime(2026, 12, 15)
};

WordContentControlFormValidationResult validation = document.ValidateContentControlValues(values);
if (!validation.IsValid) {
    Console.WriteLine(validation.ToMarkdown());
    return;
}
document.FillContentControlValues(values);
document.Save("request-filled.docx");

foreach (var pair in document.ExtractContentControlValues()) {
    Console.WriteLine($"{pair.Key} = {pair.Value}");
}
```

### PowerShell

```powershell
New-OfficeWord -Path .\request-form.docx {
    Add-OfficeWordParagraph { Add-OfficeWordText 'Requester: '; Add-OfficeWordContentControl -Text 'Name' -Alias 'Requester' -Tag 'Requester' }
    Add-OfficeWordParagraph { Add-OfficeWordText 'Approved: '; Add-OfficeWordCheckBox -Alias 'Approved' -Tag 'Approved' }
}
Get-OfficeWordContentControl -Path .\request-form.docx
```

## Replace text in an existing document

For simple one-off substitutions without template markers, replace text directly.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("report.docx");
int count = document.ReplaceAll("FY25", "FY26");
Console.WriteLine($"{count} replacements");
document.Save();
```

### PowerShell

```powershell
Update-OfficeWordText -Path .\report.docx -OldValue 'FY25' -NewValue 'FY26'
```

## Keep templates reliable

- Generate into a new file and keep the approved template unchanged.
- Validate the data before changing the document, and check `WordTemplateResult` or the preflight report after.
- Refresh fields and the table of contents if the template has them (see [fields and table of contents](/docs/word/fields-toc/)).
- Apply [protection or encryption](/docs/word/protection/) last, after all content is in place.
