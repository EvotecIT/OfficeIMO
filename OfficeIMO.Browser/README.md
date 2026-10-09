# OfficeIMO.Browser

OfficeIMO.Browser embeds the portable scripts built by [OfficeIMO JavaScript](../OfficeIMO.JavaScript/README.md). It exposes exact script content and SHA-256-based filenames to .NET applications that generate offline reports. It does not execute JavaScript and has no runtime dependencies.

The TypeScript package owns XLSX/CSV/PDF behavior, public core/ZIP/XML/OPC layers, npm exports, generated declarations and browser qualification. This project consumes its committed build output directly; there is no second JavaScript source tree or copied bundle directory.

The XLSX assets support [live conditional formatting](../OfficeIMO.JavaScript/README.md#live-excel-conditional-formatting), including threshold/formula highlights, color scales and gradient data bars. Rules use bounded workbook metadata and preserve existing number/date formats unless an override is supplied.

## Use embedded assets

```csharp
using OfficeIMO.Browser;

string script = BrowserAssets.Script.Content;
string immutableName = BrowserAssets.Script.HashedFileName;
```

Each `BrowserAsset` exposes `FileName`, `Content`, `ContentHash` and `HashedFileName`. The hash is the first 16 lowercase hex digits of SHA-256 over UTF-8 bytes without BOM. Write `Content` as UTF-8 without BOM when delivering that filename.

| Asset | Surface |
| --- | --- |
| `Script` | Combined classic script, global `OfficeIMO` with all seven public namespaces |
| `XlsxScript` | Classic XLSX plus core/ZIP/XML/OPC/XLSX namespaces |
| `CsvScript` | Classic CSV plus core/CSV namespaces |
| `PdfScript` | Classic PDF tables plus core/PDF namespaces |
| `Module` | Combined standalone ES module |
| `XlsxModule` | Standalone XLSX ES module |
| `CsvModule` | Standalone CSV ES module |
| `PdfModule` | Standalone PDF ES module |
| `DataTablesScript` | Optional classic DataTables/Buttons bridge with XLSX/CSV/PDF writers |
| `DataTablesModule` | Optional standalone DataTables bridge ES module |
| `CanopyXScript` | Optional classic CanopyX record capture adapter with XLSX/CSV/PDF writers |
| `CanopyXModule` | Optional standalone CanopyX capture adapter ES module |

Classic scripts compose in either order and work in a plain `file://` page or a host-owned worker. The `Workbook`, `writeCsv` and `saveBlob` root helpers remain available. Browser local-file policy can block ES module imports; serve modules through HTTP(S). The standalone `.mjs` assets contain their module graph and require no relative runtime imports.

The [offline current-view example](../OfficeIMO.Browser.Examples/README.md) produces an inline HTML report and an HTML/script bundle. Both export filtered, sorted, reordered and visible table data without a server. Full-dataset Excel generation at report-build time remains on OfficeIMO.Excel. Grid, iframe, CSP and download policy belong to the host integration.

The PDF assets provide paginated tables, repeated headings, totals, display styling, page decorations and native TrueType subsetting. Supply embedding-permitted font bytes for Unicode text; no fonts or font loader are embedded in the package. The [PDF contract](../OfficeIMO.JavaScript/README.md#pdf-tables) owns supported text/layout profiles and resource limits.

The DataTables assets expose `registerDataTablesButtons`, `createDataTablesExport`, `exportDataTable` and `writeDataTableTo`. They use an already installed DataTables/Buttons host and do not embed DataTables, jQuery or a third-party document writer. Loading an asset does not register buttons automatically. The [JavaScript integration guide](../OfficeIMO.JavaScript/README.md#datatables-excel-csv-and-pdf-exports) owns options and compatibility limits.

The CanopyX assets expose `createCanopyExport`, `exportCanopy` and `writeCanopyTo`. They consume an immutable native grid capture without embedding CanopyX or registering UI controls. The [CanopyX integration guide](../OfficeIMO.JavaScript/README.md#canopyx-record-exports) owns raw/display values, semantic tones, datetime preservation, streaming and presentation diagnostics.

## Build or install a local asset package

From the OfficeIMO repository root:

```sh
dotnet pack OfficeIMO.Browser/OfficeIMO.Browser.csproj -c Release -o /path/to/local-packages
dotnet add /path/to/report/Report.csproj package OfficeIMO.Browser --version 3.4.4 --source /path/to/local-packages
```

The package targets .NET Standard 2.0, .NET 8, .NET 10 and .NET Framework 4.7.2 on Windows. Its build embeds `OfficeIMO.JavaScript/bundles/*.js` and `*.mjs`; Node is unnecessary for a .NET build or consuming application. Contributors changing TypeScript run `npm ci`, `npm run build` and `npm test` in `OfficeIMO.JavaScript` and commit the regenerated bundles.

The [JavaScript package guide](../OfficeIMO.JavaScript/README.md) owns installation, examples, options and format limits. The [architecture guide](../Docs/officeimo.javascript-architecture.md) explains how npm and the asset adapter map to each other. [Browser verification](../Build/BrowserExports/README.md) tests exact embedded bytes/hashes, C# readers, Open XML validation and isolated npm/NuGet consumers.
