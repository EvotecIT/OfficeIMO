# OfficeIMO.Comic

Read CBZ, CB7 and stored RAR5 CBR comics as sequences of numbered page images. An optional `ComicInfo.xml`
provides the title, contributors, cover and page bookmarks.
A comic named `.cbr` can also contain ZIP or 7z data.

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
officeimo convert comic.cb7 comic.pdf
officeimo convert comic.cbr comic.epub
```

[Supported features and limits](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Comic/SUPPORT.md)
