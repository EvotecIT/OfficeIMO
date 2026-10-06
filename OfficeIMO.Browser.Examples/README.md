# Offline current-view exports

This minimal HtmlForgeX example generates a single HTML file with the OfficeIMO runtime embedded, plus an HTML/JavaScript bundle using the .NET asset's content-hashed filename. Open either page directly from disk. It needs no server or internet connection.

From the repository root:

```sh
dotnet run --project OfficeIMO.Browser.Examples/OfficeIMO.Browser.Examples.csproj -c Release -- /path/to/example-output
```

Open `current-view.html` or `current-view-bundle.html` in that output directory. Filter the three controllers, reverse their name order, hide Site or move it to the first/last column, then export Excel or CSV. Exports contain the current displayed rows and visible columns in their current order. Excel uses a native table, a blue header, frozen row/column, numeric latency values with two decimal places and warning highlighting above 20 ms. CSV uses a UTF-8 BOM and the package's default formula protection. An empty view exports just the visible headers.

The page uses HtmlForgeX's existing default styles and a plain typed table. Its small script projects the displayed DOM; a data-grid integration should supply the grid's own current-view iterable, especially when the grid virtualizes rows. The [OfficeIMO.Browser README](../OfficeIMO.Browser/README.md) owns the writer API and limits.

This example stays outside the default solution. HtmlForgeX is an example-only dependency and is absent from the OfficeIMO.Browser NuGet and npm artifacts.
