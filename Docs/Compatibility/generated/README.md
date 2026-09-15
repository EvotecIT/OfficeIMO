# Generated Office compatibility contracts

These files are generated from the public format and capability catalogs. Do not edit them by hand.

Regenerate:

```powershell
dotnet run --framework net8.0 --project Build/CompatibilityCatalog/OfficeIMO.CompatibilityCatalog.Tool.csproj -- --output Docs/Compatibility/generated --converter-sample Website/Apps/OfficeIMO.Web.Converter/wwwroot/samples/conversion-proof.pptx
```

Use the [conversion route catalog](conversion-routes.md) to find the focused package, output model, proven support level, known limits, browser availability, and result type for each route.
Use the [package-neutral operation catalog](package-operations.md) to compare create, read, edit, preserve, inspect, validate, remove, convert, and export outcomes without losing the detailed owning contract.

Verify:

```powershell
dotnet run --framework net8.0 --project Build/CompatibilityCatalog/OfficeIMO.CompatibilityCatalog.Tool.csproj -- --output Docs/Compatibility/generated --converter-sample Website/Apps/OfficeIMO.Web.Converter/wwwroot/samples/conversion-proof.pptx --verify
```

| Contract | Schema | Rows | JSON | Markdown |
| --- | ---: | ---: | --- | --- |
| OfficeIMO.Word.LegacyDoc | 1 | 33 | [JSON](word-legacy-doc.json) | [Markdown](word-legacy-doc.md) |
| OfficeIMO.Excel.LegacyXls | 1 | 28 | [JSON](excel-legacy-xls.json) | [Markdown](excel-legacy-xls.md) |
| OfficeIMO.Excel.Xlsb | 1 | 20 | [JSON](excel-xlsb.json) | [Markdown](excel-xlsb.md) |
| OfficeIMO.PowerPoint.LegacyPpt | 1 | 56 | [JSON](powerpoint-legacy-ppt.json) | [Markdown](powerpoint-legacy-ppt.md) |
| OfficeIMO.Operations | 1 | 1302 | [JSON](package-operations.json) | [Markdown](package-operations.md) |
| OfficeIMO.Provenance | 1 | 11 | [JSON](provenance.json) | [Markdown](provenance.md) |
| OfficeIMO.ProtectedContent | 1 | 18 | [JSON](protected-content.json) | [Markdown](protected-content.md) |

`office-formats.json` is the concrete extension, document-kind, encoding, and macro-carrier inventory used by conversion routing.
