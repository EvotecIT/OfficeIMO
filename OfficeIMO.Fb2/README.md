# OfficeIMO.Fb2

Read FictionBook 2 from an `.fb2` file or an `.fb2.zip` archive with one book.
The reader builds an EPUB publication from the title and author metadata, body
sections, notes, links and embedded images.

```csharp
using OfficeIMO.Fb2;

using var source = File.OpenRead("book.fb2.zip");
var imported = Fb2Reader.Import(source, "book.fb2.zip");
imported.Value.Save("book.epub");
```

Use `imported.Report.Diagnostics` to review omitted or approximated content.
`OfficeIMO.Reader.All` supplies the rich Reader projection. The Workflows package
and the CLI can also convert the book to PDF, HTML or Markdown.

```sh
officeimo convert book.fb2.zip book.pdf
```

[Supported features and limits](SUPPORT.md)
