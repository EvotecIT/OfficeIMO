# OfficeIMO.Comic

Read a CBZ as a sequence of numbered page images. An optional `ComicInfo.xml`
provides the title, contributors, cover and page bookmarks.

Install the package from NuGet.

```sh
dotnet add package OfficeIMO.Comic
```

```csharp
using OfficeIMO.Comic;

using var source = File.OpenRead("comic.cbz");
var imported = ComicReader.Import(source, "comic.cbz");
imported.Value.Save("comic.epub");
```

The EPUB retains the page dimensions in a fixed-layout spine. Use
`imported.Report.Diagnostics` to inspect skipped pages or metadata.

`OfficeIMO.Reader.All` exposes the pages and images. Workflows and the CLI can
convert the comic to EPUB, PDF, HTML or Markdown.

```sh
officeimo convert comic.cbz comic.pdf
```

[Supported features and limits](SUPPORT.md)
