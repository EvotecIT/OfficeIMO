# OfficeIMO.Drawing.CodeGlyphX integration tests

This non-packable test project verifies the neutral SVG boundary and the optional typed adapter without adding a dependency to either core package.

By default, it tests against the published CodeGlyphX package. To validate an adjacent CodeGlyphX source checkout, pass its project explicitly:

```powershell
dotnet test -f net10.0 `
  -p:CodeGlyphXProjectPath="<path-to-CodeGlyphX.csproj>"
```

The smoke suite covers default and styled QR output, Data Matrix, DataBar Expanded stacked output, linear barcode shapes, and searchable human-readable text.
