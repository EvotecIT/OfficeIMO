# OfficeIMO.OpenDocument

`OfficeIMO.OpenDocument` is a dependency-free native OpenDocument engine for ODT text documents, ODS spreadsheets, and ODP presentations.

The package owns ZIP/XML packaging, metadata, styles, preservation-aware editing, and typed document APIs. It does not require LibreOffice, Microsoft Office, UNO, or third-party runtime packages.

```csharp
using OfficeIMO.OpenDocument;

using OdtDocument document = OdtDocument.Create();
document.AddHeading("Summary", 1);
document.AddParagraph("Created with OfficeIMO.OpenDocument.");
document.Save("summary.odt");
```

The first implementation targets ODF 1.4 and can write an ODF 1.3 compatibility profile. Unknown package entries and extension XML are preserved by default when existing documents are edited.
