# OfficeIMO.Reader.All

`OfficeIMO.Reader.All` is the explicit composition package for applications that want every local OfficeIMO format handler without registering individual adapters.

## Install

```powershell
dotnet add package OfficeIMO.Reader.All
```

## Use the preset

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddAllOfficeIMOHandlers()
    .WithMaxConcurrentReads(4)
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument("input.epub");
```

## Convert Email, EPUB, or Visio directly to PDF

The all-adapters package also exposes thin direct façades. They normalize through
the owning Reader adapter, then project through `OfficeIMO.Reader.Pdf`; they do
not introduce a second renderer.

```csharp
using OfficeIMO.Reader.All;

PdfDocumentConversionResult email = OfficeDocumentPdfConverter.EmailToPdf("message.eml");
PdfDocumentConversionResult book = OfficeDocumentPdfConverter.EpubToPdf("book.epub");
PdfDocumentConversionResult diagram = OfficeDocumentPdfConverter.VisioToPdf("diagram.vsdx");

email.Save("message.pdf");
book.Save("book.pdf");
diagram.Save("diagram.pdf");
```

Email attachments, EPUB resources and chapter pagination, and Visio
preview-versus-semantic fallback behavior follow `ReaderPdfProjectionOptions`
and are reported as structured conversion evidence.

The preset adds Word, Excel, PowerPoint, Markdown, direct email artifacts, Outlook stores and OAB address books, plus AsciiDoc, CSV/TSV, EPUB, HTML/MHTML, standalone images, JSON, LaTeX, Jupyter Notebook, offline OneNote, OpenDocument, PDF, RTF, subtitles, Visio, XML, YAML, and ZIP handlers. `OfficeIMO.Reader.Core` itself contains no format parser.

Configure a format through one options object:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddAllOfficeIMOHandlers(new ReaderAllOptions {
        Csv = new OfficeIMO.Reader.Csv.CsvReadOptions {
            ChunkRows = 100
        },
        Email = new OfficeIMO.Reader.Email.ReaderEmailHandlersOptions {
            Stores = new OfficeIMO.Reader.Email.ReaderEmailStoreOptions {
                StoreOptions = new OfficeIMO.Email.Store.EmailStoreReaderOptions(
                    retainAttachmentContent: false)
            },
            AddressBooks = new OfficeIMO.Reader.Email.ReaderEmailAddressBookOptions {
                MaxEntries = 10_000
            }
        },
        OneNote = new OfficeIMO.Reader.OneNote.ReaderOneNoteOptions {
            IncludeConflictPages = true,
            IncludeVersionHistory = true
        },
        ZipTraversal = new OfficeIMO.Zip.ZipTraversalOptions {
            MaxEntries = 500,
            MaxTotalUncompressedBytes = 64L * 1024 * 1024
        }
    })
    .Build();
```

Registrations are copied into the builder's immutable snapshot. The preset does not mutate process-wide reader state.

## Dependency boundary

This package contains no parser, provider, model, native binary, process launcher, or network client. It references OfficeIMO's existing local adapter packages and therefore carries their established managed dependency graph. OCR packages are deliberately excluded because they require an engine or executable; add one explicitly only when the host chooses it.
