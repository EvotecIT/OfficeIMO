# OfficeIMO.Invoices

`OfficeIMO.Invoices` loads existing UN/CEFACT CrossIndustryInvoice XML in namespace version 100, reads selected header fields, and edits an existing invoice identifier or format-102 issue date. It also owns the CII inspection primitives used by `OfficeIMO.Pdf`.

```csharp
using OfficeIMO.Invoices;

var original = CiiInvoiceDocument.Load(File.ReadAllBytes("invoice.xml"));
Console.WriteLine($"{original.DocumentId}: {original.CurrencyCode}");

var edited = original
    .WithDocumentId("INV-2026-0002")
    .WithIssueDate(new DateTime(2026, 9, 5));
File.WriteAllBytes("edited-invoice.xml", edited.ToBytes());
```

The model is immutable. `Load` and `ToBytes` make defensive copies. Saving an unedited model returns the original bytes, including encoding and formatting. Editing serializes UTF-8 XML, retaining unknown elements, attributes, comments, and whitespace outside the replaced field. XML character content is preserved, but original quotation choices, entity spelling, and other lexical formatting can change.

Field access follows exact namespace-qualified paths. Duplicate fields throw instead of selecting an arbitrary value. Edits require an existing scalar field; nested elements, comments, or processing instructions inside that field cause rejection. Unknown date representations remain untouched. XML containing a standard XML Signature element can be loaded and saved unchanged, but cannot be edited through this model. Signature verification is outside this API.

Input is limited to 16 MiB and XML depth 128. DTDs and external entity resolution are disabled. Other CII root namespace versions and UBL are outside this model's contract.

## PDF attachment

Use the PDF package to attach a snapshot of the model:

```csharp
using OfficeIMO.Pdf;

var options = new PdfOptions().UseFacturXDocument(edited);
```

`UseFacturXDocument` uses the same PDF/A-3 groundwork as `UseFacturX(byte[])`. It does not generate the visible invoice. Changing the identifier does not change payment references, other XML fields, or an existing PDF page. The application must keep those representations consistent and run the PDF and invoice validation appropriate to its declared profile.

Loading XML is not a schema or business-rule validation result. This package does not author complete invoices, calculate taxes or totals, or certify EN 16931, Factur-X, ZUGFeRD, XRechnung, or Peppol compliance. Existing PDF compliance assessment and independent-validator workflows remain separate evidence.

Targets: .NET Standard 2.0, .NET 8, .NET 10, and .NET Framework 4.7.2.
