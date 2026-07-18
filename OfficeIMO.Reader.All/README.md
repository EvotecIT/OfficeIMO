# OfficeIMO.Reader.All

`OfficeIMO.Reader.All` is a thin composition package for applications that want OfficeIMO's local format handlers without registering every adapter separately.

The project is packable but is not currently published on NuGet. Until publication is enabled deliberately, use the
individual `OfficeIMO.Reader.*` packages or a source project reference. The modular package roadmap requires
`Reader.All` to become a tested, explicit meta package or be removed rather than remain ambiguous.

## Install

```powershell
dotnet add package OfficeIMO.Reader.All # available after deliberate publication
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

The preset adds AsciiDoc, CSV/TSV, Outlook OAB address books, PST/OST/OLM/EMLX email stores, EPUB, HTML/MHTML, standalone image, JSON, LaTeX, Jupyter Notebook, offline OneNote (`.one`, `.onetoc2`, and `.onepkg`), OpenDocument, PDF, RTF, SRT/WebVTT subtitle, Visio, XML, YAML, and ZIP handlers. Word, Excel, PowerPoint, Markdown, individual email artifacts, and plain text remain built into `OfficeIMO.Reader`.

Configure a format through one options object:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddAllOfficeIMOHandlers(new ReaderAllOptions {
        Csv = new OfficeIMO.Reader.Csv.CsvReadOptions {
            ChunkRows = 100
        },
        EmailStore = new OfficeIMO.Reader.EmailStore.ReaderEmailStoreOptions {
            StoreOptions = new OfficeIMO.Email.Store.EmailStoreReaderOptions(
                retainAttachmentContent: false)
        },
        EmailAddressBook = new OfficeIMO.Reader.EmailAddressBook.ReaderEmailAddressBookOptions {
            MaxEntries = 10_000
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
