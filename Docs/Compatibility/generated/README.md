# Generated Office compatibility contracts

These files are generated from the public format and capability catalogs. Do not edit them by hand.

Regenerate:

```powershell
dotnet run --project Build/CompatibilityCatalog/OfficeIMO.CompatibilityCatalog.Tool.csproj -- --output Docs/Compatibility/generated
```

Verify:

```powershell
dotnet run --project Build/CompatibilityCatalog/OfficeIMO.CompatibilityCatalog.Tool.csproj -- --output Docs/Compatibility/generated --verify
```

| Contract | Schema | Rows | JSON | Markdown |
| --- | ---: | ---: | --- | --- |
| OfficeIMO.Word.LegacyDoc | 1 | 33 | [JSON](word-legacy-doc.json) | [Markdown](word-legacy-doc.md) |
| OfficeIMO.Excel.LegacyXls | 1 | 28 | [JSON](excel-legacy-xls.json) | [Markdown](excel-legacy-xls.md) |
| OfficeIMO.Excel.Xlsb | 1 | 20 | [JSON](excel-xlsb.json) | [Markdown](excel-xlsb.md) |
| OfficeIMO.PowerPoint.LegacyPpt | 1 | 56 | [JSON](powerpoint-legacy-ppt.json) | [Markdown](powerpoint-legacy-ppt.md) |

`office-formats.json` is the concrete extension, document-kind, encoding, and macro-carrier inventory used by conversion routing.
