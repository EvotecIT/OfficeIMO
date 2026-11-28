# OfficeIMO.Mobi

Read PalmDOC, Mobipocket and Kindle Format 8 ebooks. The reader selects the
KF8 section when a dual-format file contains MOBI and KF8.

Install the package from NuGet.

```sh
dotnet add package OfficeIMO.Mobi
```

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

The Reader preset handles PalmDOC `.pdb` and Mobipocket `.prc` files by default.
Application handlers registered for either extension take precedence. MOBI
imports check the Palm database signature before reading text.

[Supported features and limits](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Mobi/SUPPORT.md)
