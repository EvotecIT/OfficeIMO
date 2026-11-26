# OfficeIMO.Mobi

Read PalmDOC, Mobipocket and Kindle Format 8 ebooks. The reader selects the
KF8 section when a dual-format file contains MOBI and KF8.

```csharp
using OfficeIMO.Mobi;

using var source = File.OpenRead("book.azw3");
var imported = MobiReader.Import(source, "book.azw3");
imported.Value.Save("book.epub");
```

DRM-protected input is refused. Review `imported.Report.Diagnostics` for source
features that cannot be carried forward.

`OfficeIMO.Reader.All` supplies the rich Reader projection. Workflows and the CLI
can write EPUB, PDF, HTML or Markdown.

```sh
officeimo convert book.mobi book.pdf
```

[Supported features and limits](SUPPORT.md)
