# Structured-data conversions design

OfficeIMO converts structured data between JSON, YAML, TOML, XML, CSV/TSV and Excel worksheets. Conversions stream with bounded memory, follow the [unified API design](officeimo.api-design.md#conversion), and report every loss through `OfficeConversionReport`. This document owns the architecture and the mapping rules. Package READMEs will own the supported API, and open work belongs in the [roadmap](ROADMAP.md#yaml-engine).

## Decisions

| # | Topic | Decision |
|---|---|---|
| 1 | Architecture | A **streaming event pipeline**. Each format has a source that drives events and a sink that consumes them, so N formats need 2N pieces, not N² converters. Models (`YamlDocument`, `TomlDocument`, `JsonNode`, `XDocument`) are walked into the same sinks. There is no neutral tree and no `JsonNode` hub. |
| 2 | Packages | One hub, **`OfficeIMO.Data.Conversions`**, for JSON, YAML, TOML, XML and CSV/TSV, plus **`OfficeIMO.Data.Conversions.Excel`** for worksheets. Format packages stay free of each other and of `System.Text.Json`. The routes are built into the `OfficeIMO` front door. |
| 3 | Tree ⇄ table | Rows come from an array (the root or a `Select` path). Nested objects flatten into path headers (`address.city`, `tags[0]`) in the YAML path grammar and unflatten on the way back. Untyped sources (CSV/TSV) are type-inferred per column with round-trip-safe rules. |
| 4 | XML | The Newtonsoft-style convention: `@attr`, `#text`, kept root, repeated siblings become arrays, namespace prefixes and `xmlns` declarations kept. Two deliberate differences: `xsi:nil` for null and no `#comment`/`?pi` properties. Order the shape can't carry is reported. |
| 5 | Loss | Every lossy mapping has a defined behavior and a catalog code. Findings are aggregated per code (count plus the first locations), so a million-row file produces a short report. `FailOnLoss` throws before anything is published. |
| 6 | API | `To{Target}`/`SaveAs{Target}` extension methods on each source model, `OfficeConverter.Convert` for files and batches, one `DataConversionOptions` type, and `Convert-OfficeDocument` in PowerShell. Excel's existing `ToJson`/`TableToJson`/`FromJson` move out of `OfficeIMO.Excel`. |
| 7 | Proof | Opt-in benchmark lanes against YamlDotNet + System.Text.Json, Tomlyn, Newtonsoft's XML conversion and CsvHelper. Correctness from the yaml-test-suite `in.json`, toml-test JSON fixtures, JSONTestSuite and independent readers of every output. |
| 8 | Order | YAML→JSON and CSV/Excel→JSON first (they need only shipped readers), then JSON→YAML when `YamlWriter` lands, then XML, TOML and TSV as their engines and profiles land. |

## Architecture

### Options considered

| Option | Memory | Fidelity | Dependencies | Verdict |
|---|---|---|---|---|
| Streaming events between format readers and writers | Constant for tree⇄tree; one row for tables | Each sink sees the source's real types, comments and tags, so it reports exactly what *it* can't carry | Each piece needs only its own format engine | **Chosen** |
| Neutral in-memory value tree | O(input), about one object per value | Good, if the tree carries comments, tags and date kinds | None | Kept only as bounded local buffers (below) |
| `System.Text.Json.Nodes` as the hub | O(input), heavier than an owned tree | Lossy at the hub: no comments, tags, date-time kinds or YAML key types, so loss is detected at the wrong point | `System.Text.Json` for every route, including YAML⇄TOML | Rejected |

A 1 GB JSON array converts to YAML or CSV without holding the array. Models remain the right input for small and random-access jobs: `YamlDocument.ToJson()` walks its tree into the same JSON sink the file route uses, so both paths share one mapping.

### The event contract (internal)

```
JSON source (Utf8JsonReader)  ─┐                        ┌─ JSON sink (Utf8JsonWriter)
YAML source (YamlReader)      ─┤                        ├─ YAML sink (YamlWriter)
TOML source (TomlReader)      ─┼─► data events ──────► ─┼─ TOML sink (TomlWriter)
XML source (XmlReader)        ─┤   + loss collector     ├─ XML sink (XmlWriter)
Table source (DbDataReader)   ─┘                        └─ Table sink (DbDataReader → CsvWriter / ExcelWorkbookWriter)
Model walkers (YamlDocument, TomlDocument, JsonNode, XDocument) drive the same events.
```

- **Events:** start and end of documents, objects and arrays; property names; scalars; comments. Optional metadata travels with them: YAML tag, anchor, style and source position.
- **Scalar** is a struct with a kind: null, boolean, integer, float, string, offset date-time, local date-time, local date or local time. Numbers and date-times also keep their **source text**, so `1.10`, `12345678901234567890123` and `1979-05-27T07:32:00.999999-07:00` pass through exactly when the target can hold them. Sinks decide from the kind and the text; nothing is parsed to `double` on the way.
- **Pull, not push:** a converter pulls from the source and writes to the sink in one loop. Table sinks invert this: the tree source is wrapped as a `DbDataReader` that materializes one row at a time, which plugs straight into the existing `CsvWriter` and `ExcelWorkbookWriter.WriteTable(IDataReader)` streaming paths. In the other direction `CsvDataReader` and `ExcelWorkbookDataReader` are already `DbDataReader`s.
- **Visibility:** internal in v1, shared between the two packages through `InternalsVisibleTo`. Making the contract or the `DbDataReader` bridge public waits for a real consumer (see [open questions](#open-questions)).

### Where buffering is unavoidable

Each buffer is bounded by a limit and reported in the README:

| Case | Buffer | Bound |
|---|---|---|
| XML repeated siblings | Adjacent same-name siblings stream with one-sibling lookahead. Non-adjacent repeats (`a, b, a`) buffer that parent to group them | One record, or one parent |
| YAML aliases expanded for non-YAML targets | The anchored subtree | `MaxAliasExpansion` from `YamlReaderOptions` |
| TOML sink | Within a table, scalars must precede sub-tables, so object members are held until the table's scalars are written | One table's subtree; TOML files are small configuration |
| XML root children | A first pass reads only the root's child names; interleaved names (`a, b, a`) buffer the root to group them, otherwise records stream | One record, or the root when its children interleave |
| Tree → table column discovery | Pass 1 scans events for column paths and type evidence; pass 2 writes rows | The column set only. The input is re-read; `OfficeConverter` already captures input to re-readable bounded storage. An explicit `Columns` list fixes the columns, but pass 1 still decides the shape |
| CSV/TSV type inference | Pass 1 decides each column's type | Per-column state only |

### Performance rules

- `Utf8JsonReader`/`Utf8JsonWriter` over pooled buffers. No `JsonDocument` on the file routes.
- No per-scalar strings unless the target needs them. YAML keys go through the engine's string cache. JSON property names that repeat (every row of an array of records) are encoded once with `JsonEncodedText`.
- Header paths for table routes are compiled once into a write plan, so each row is a flat loop.
- No LINQ, closures or boxing in event loops. The engines' AOT and trimming annotations carry through; no reflection anywhere in the hub.

## Packages

| Package | Depends on | Owns |
|---|---|---|
| `OfficeIMO.Data.Conversions` | `OfficeIMO.Core`, `OfficeIMO.Yaml`, `OfficeIMO.Toml`, `OfficeIMO.CSV`; `System.Text.Json` on netstandard2.0/net472 only (in the box on net8.0+); `System.Xml` (in the box) | Event contract, every source and sink except Excel, tree ⇄ table mapping, `DataConversionOptions`, `To{Target}`/`SaveAs{Target}` extensions on `YamlDocument`, `TomlDocument`, `CsvDocument`, `JsonNode` and `XDocument` |
| `OfficeIMO.Data.Conversions.Excel` | The hub, `OfficeIMO.Excel` | Worksheet and workbook sources and sinks; `ToJson`/`ToYaml`/`ToToml`/`ToXml` and `SaveAs…` on `ExcelSheet`/`ExcelDocument`; `ToExcel`/`SaveAsExcel` on the tree sources |
| `OfficeIMO` (front door) | Adds both packages | Registers the routes as built-ins, so `OfficeConverter.Convert("a.yaml", "a.json")` works with one package reference |

Why not the alternatives:

- **N² adapters** (`OfficeIMO.Yaml.Json`, `OfficeIMO.Json.Toml`, …): six formats would need about 30 packages that share one mapping. Principle 7 exists to stop heavy engines from leaking into each other. JSON, YAML, TOML, XML and CSV are small, dependency-free text engines (JSON and XML come from the BCL), so one hub costs a few hundred kilobytes, not a third-party dependency. Excel is the one heavy engine (Open XML SDK), so it stays behind its own package. This deliberate exception to the pairwise-adapter rule should be recorded in the API design's Packages section.
- **Conversions inside each format** (`YamlDocument.ToJson()` in `OfficeIMO.Yaml`): this adds `System.Text.Json` to `OfficeIMO.Yaml` on old targets, breaking the YAML design's no-dependency goal. It also can't host YAML→TOML without one format package depending on another.
- **Workflows-only routes:** this gives file conversion but no `To{Target}` model API. Workflows stays a host; the mapping lives in the hub.

Namespaces follow the package (`OfficeIMO.Data.Conversions`), as `OfficeIMO.Excel.Csv` does. Targets are netstandard2.0, net8.0, net10.0 and net472, matching the engines, with `IsAotCompatible` on net8.0+. `OfficeIMO.Reader.Json` and `OfficeIMO.Reader.Xml` stay ingestion adapters; they can adopt the hub's sources later without API change.

## Tree ⇄ table mapping

### Which shapes are tabular

The rows are taken from the root, or from `Select` (a path in the shared YAML path grammar, such as `data.items`), which covers the very common `{ "data": [...], "meta": {...} }` API response.

| Shape at the root or `Select` | Table result |
|---|---|
| Array of objects | One row per object; columns are the union of flattened paths in first-seen order |
| Array of scalars | One column named `value` |
| Array of arrays | Positional rows with no header row. The reverse is a header-less CSV, so the pair is symmetric |
| Object whose every member is an array | One table per member: Excel sheets named by key. CSV/TSV hold one table, so the conversion fails and the message lists the `Select` paths to choose from |
| Any other object | One row |
| Multi-document YAML stream | An array of the documents, so Kubernetes manifests become one row per resource |

Workbooks map back symmetrically. A workbook with one sheet (or a chosen `Sheet`) becomes an array. With several sheets it becomes an object keyed by sheet name. Sheets read with first-row headers through the existing `ExcelReaderOptions`. Formulas contribute their cached values, which is reported.

### Flattening

- **Objects** flatten to dotted paths, `ObjectFlattener`'s convention: `address.city`.
- **Arrays inside a row** flatten to indexed columns by default: `tags[0]`, `tags[1]`, `items[0].sku`. This is lossless, and the path syntax is the YAML engine's path grammar, so a header can be fed to `Select` or `YamlNode.Select` unchanged. Ragged arrays produce the union of indexes.
- `ArrayColumns = Join` writes scalar arrays into one cell joined by `ArraySeparator` (default `", "`) for human-readable sheets, as `ObjectFlattener` does today. Joining is lossy and reported.
- If the column count would exceed `MaxColumns` (Excel's 16,384 by default), the conversion fails and suggests `Select` or `Join`. It never silently truncates.
- **Unflattening (table → tree)** applies when every header is a well-formed path: no empty segments, numeric indexes, no conflicts such as both `a` and `a.b`. Otherwise headers stay literal, and a conflict is reported. Contiguous indexes `[0]..[n]` rebuild arrays, and missing trailing cells shorten them.
- A tree key containing `.`, `[` or `]` can't round-trip as a header. It is written literally and reported (`header-path-ambiguous`).

### Types

**Typed sources** (JSON, YAML, TOML, Excel) keep their types. Strings are never re-interpreted, so `"0042"` in JSON is a text cell in Excel.

**Untyped sources** (CSV/TSV) are inferred **per column**, so one column never mixes numbers and strings:

| Column type | Every non-empty cell must be | Notes |
|---|---|---|
| Integer, then decimal | A canonical number under `Culture` (invariant by default): no leading zeros except `0`, no leading `+`, no thousands separators | `007`, `+1`, `1,000` keep the column as strings, which protects IDs, ZIP codes and phone numbers |
| Boolean | `true` or `false`, case-insensitive | `yes`, `Y` and `1` stay strings |
| Date-time | ISO 8601 / RFC 3339 | Locale formats such as `10/11/2026` stay strings |
| String | Anything else | |

- Empty cells are `null` in typed columns and `""` in string columns.
- `InferTypes = false` keeps every value a string.
- XML text is untyped too, but XML defaults to strings because per-path inference would need a pre-pass over the whole document. With `InferTypes = true`, XML uses the same canonical rules per value.

**Writing tables:**

- Numbers are written from their source text, formatted for `Culture` in CSV.
- Booleans are `true`/`false`, date-times are ISO 8601 and `null` is an empty field.
- Into Excel, numbers become numeric cells when the value survives a double round-trip. Otherwise they become text cells with `number-precision`, which protects 16+ digit IDs from becoming `1.23457E+15`.
- Date kinds become date-formatted cells; an offset is dropped with `datetime-offset-dropped`, keeping the wall-clock time.
- A null in a text column becomes an empty field, which reads back as `""`; that is reported (`null-as-empty`) when it occurs.
- An empty array or object inside a row has no cells, so it becomes an empty cell and is reported (`empty-collection-dropped`). An empty trailing array item reads back as absent.

## XML ⇄ JSON, YAML and TOML

### Conventions compared

| Convention | Attributes | Repeated elements | Mixed content and order | Round-trip of data XML |
|---|---|---|---|---|
| Parker | Dropped | Arrays | Dropped | No |
| BadgerFish | `@attr`; text always `$` | Arrays | Lost | Yes, but every value becomes an object (`{"$": "x"}`), which nobody wants as data |
| JsonML | Attribute object | Positional | Kept exactly | Yes, but the output is a markup tree, not data |
| Newtonsoft (`SerializeXNode`) | `@attr`; text `#text` only when needed | Arrays | `#text` merged; comments as `#comment` | Yes for data XML, and the convention .NET developers already know |

**Chosen: the Newtonsoft-style convention,** because data-shaped XML round-trips exactly and the output is the JSON people expect. The rules:

- **Root:** the root element is kept as the single top-level property: `<order>` becomes `{ "order": { … } }`. From JSON, a single-property object supplies the root name. Any other root (an array, or several properties) is wrapped in `RootName`, which defaults to `root`.
- **Attributes:** `@name`. Namespace declarations are attributes too (`@xmlns`, `@xmlns:x`), and names keep their prefixes as written (`x:item`). The same prefixes and declarations are written back, so namespaces round-trip without a separate mapping.
- **Text:** an element with only text and no attributes becomes a string. With attributes or children, its text is `#text`. CDATA is merged into text, because CDATA boundaries are not part of the XML Infoset.
- **Repeated elements:** same-name siblings become an array. Non-adjacent repeats are grouped, and the change in sibling order is reported (`xml-sibling-order-changed`).
- **Empty and null:** `<a/>` and `<a></a>` both become `""`, since the Infoset doesn't distinguish them. `xsi:nil="true"` becomes `null`, and `null` is written as `xsi:nil="true"`. This is the standard XML Schema marker, where Newtonsoft writes an ambiguous empty element.
- **Mixed content:** text around child elements becomes `#text` (an array if there are several runs), and its interleaving is reported (`xml-mixed-content`). Document-centric XML (XHTML, DocBook) is not a data-conversion input; OfficeIMO's document formats handle it.
- **Comments:** carried to YAML and TOML as comments. JSON reports `comments-omitted`; there are no `#comment` properties polluting the data.
- **Processing instructions, the XML declaration and DOCTYPE:** not carried. Processing instructions are reported. DTDs are prohibited (`DtdProcessing.Prohibit`), as in every untrusted-input path, and the error says so.
- **Insignificant whitespace** between elements is dropped unless `xml:space="preserve"` is set.
- **JSON → XML:**
  - Array items repeat the property's element name, and items of a nested array are named `item`.
  - Invalid XML names are encoded with `XmlConvert.EncodeName` (`first name` → `first_x0020_name`) and decoded on the way back.
  - All scalars become text, so their types are reported once per conversion (`scalar-type-dropped`).
  - Arrays with zero or one item can't be told apart from absent or single elements, so `xml-array-shape` reports them when they occur.

A typed, order-exact representation (W3C XPath 3.1 `json-to-xml`, or JsonML) is an option only when a scenario such as XSLT 3.0 interchange needs it. It is not in v1.

## Type fidelity and loss

Each case has one behavior. Codes are added to `OfficeIMO.Core/Diagnostics/diagnostic-code-map.json` and generated into `OfficeDiagnosticCodes`. Loss kinds: **A** is Approximation, **O** is Omission, **–** is no loss (Information severity). Each finding is aggregated per code with a count and up to ten locations. Locations are source positions: `line 12, column 5` for YAML and XML, `byte 1834` for JSON and `row 12` for tables, or a header path for column findings. Tracking a path for every value would cost an allocation per key.

| Case | Behavior | Code | Loss |
|---|---|---|---|
| Comments into JSON, CSV, TSV or Excel | Dropped | `comments-omitted` (existing) | O |
| Comments between YAML, TOML and XML | Carried, attached to the following node | – | – |
| JSON comments | Read only when the caller's `JsonReaderOptions` allow them (JSONC such as `tsconfig.json`), then carried as above | – | – |
| YAML anchors and aliases into other formats | Expanded within `MaxAliasExpansion`. The YAML specification calls anchors a serialization detail, so the data is unchanged | `yaml-alias-expanded` | – |
| YAML merge keys `<<` (1.1) | Resolved into the mapping | – | – |
| YAML standard tags (`!!str`, `!!int`, …) | Applied as the value's type | – | – |
| YAML custom tags (`!Ref`, `!Sub`) | Value kept, tag dropped | `yaml-tag-dropped` | A |
| YAML complex keys (a mapping or sequence as a key) | Key written as its flow-YAML text | `yaml-complex-key` | A |
| Non-string keys (`1:`, `true:`) into JSON, TOML or XML | Written as strings | `key-type-changed` | A |
| YAML plain scalars shaped like timestamps | Follow the schema: strings under the 1.2 Core schema; date-times under the 1.1 schema (a `%YAML 1.1` directive or a forced schema) or an explicit `!!timestamp` tag | – | – |
| Multi-document YAML into JSON | An array of the documents | `yaml-documents-combined` | A |
| Duplicate keys | Error by default, as YAML and TOML require. A source's `LastWins`/`FirstWins` policy reports each dropped value | `duplicate-key` | O |
| Big or high-precision numbers | Exact text into JSON and YAML | – | – |
| Integers outside int64 into TOML | Written as strings | `number-out-of-range` | A |
| More than 15 significant digits into Excel | Text cell | `number-precision` | A |
| NaN and ±Infinity into JSON | Strings `"NaN"`, `"Infinity"`, `"-Infinity"`, which `System.Text.Json`'s named floating-point literals read back as numbers | `non-finite-number` | A |
| TOML date kinds into JSON | RFC 3339 text | `datetime-as-string` | A |
| TOML date kinds into YAML | Plain scalars; local times are quoted, because YAML 1.1 reads `07:32:00` as a base-60 integer | `datetime-as-string` for local times only | A |
| Date-times with an offset into Excel | Wall-clock time kept, offset dropped | `datetime-offset-dropped` | A |
| Fractional seconds beyond the target's precision (Excel: milliseconds) | Rounded | `datetime-precision` | A |
| `null` into TOML | Key omitted | `null-dropped` | O |
| Non-table root into TOML; array root into XML | Wrapped in `RootName` | `root-wrapped` | A |
| Scalar types into XML | Text | `scalar-type-dropped` | A |
| Strings into CSV that inference would read back as other types (`"42"`) | Written as text | `scalar-type-dropped` | A |
| XML non-adjacent repeated siblings | Grouped | `xml-sibling-order-changed` | A |
| XML mixed content | `#text` runs | `xml-mixed-content` | A |
| XML processing instructions | Dropped | `xml-processing-instruction-dropped` | O |
| Zero- or one-item arrays into XML | Repeated or absent elements | `xml-array-shape` | A |
| Arrays joined into one cell | `ArrayColumns = Join` | `array-joined` | A |
| Keys containing `.`, `[` or `]` into a table | Literal header | `header-path-ambiguous` | A |
| Excel formulas | Cached values; the worksheet reader does not expose whether a cell held a formula, so this is not reported | – | – |
| Characters XML 1.0 cannot hold (NUL, most C0 controls, U+FFFE, U+FFFF) | Replaced with U+FFFD | `xml-character-replaced` | A |
| Empty keys into XML | `<_>` elements | `xml-name-changed` | A |
| Empty arrays and objects inside table rows | Empty cells | `empty-collection-dropped` | O |
| Nulls in text columns into CSV | Empty fields | `null-as-empty` | A |

`FailOnLoss` (in `DataConversionOptions` and `OfficeConversionOptions`) throws `OfficeConversionException` carrying the report before the output is published. Conversion is streaming, so this means `SaveAs*` writes to a temporary target and publishes only on success, as `OfficeConverter` already does. An `Information` finding never trips it.

## Public API

### C#

```csharp
using OfficeIMO.Data.Conversions;

// YAML ⇄ JSON
using var values = YamlDocument.Load("values.yaml");
JsonNode json = values.ToJson();                        // the JSON model; .ToJsonString() for text
OfficeConversionReport report = values.SaveAsJson("values.json");
JsonNode.Parse(File.ReadAllText("app.json"))!.SaveAsYaml("app.yaml");

// TOML, XML
using var cargo = TomlDocument.Load("Cargo.toml");
cargo.SaveAsJson("cargo.json");
XDocument.Load("orders.xml").SaveAsYaml("orders.yaml");
XDocument xml = values.ToXml(new DataConversionOptions { RootName = "values" });

// Tables
using var csv = CsvDocument.Load("people.csv");
csv.SaveAsJson("people.json");                         // inferred types, nested headers unflattened
JsonNode.Parse(apiResponse)!.SaveAsCsv("items.csv", new DataConversionOptions { Select = "data.items" });
values.SaveAsExcel("values.xlsx");                       // OfficeIMO.Data.Conversions.Excel
using var book = ExcelDocument.Load("Sales.xlsx");
book["Q3"].SaveAsYaml("q3.yaml");
YamlDocument q3 = book["Q3"].ToYaml(out OfficeConversionReport loss);
```

- Every `To{Target}` has the `out OfficeConversionReport` overload. Async forms are not in the first release: the pipeline is synchronous and CPU-bound, and an async wrapper would only move it to another thread.
- `StructuredDataReader` is public: the rows of JSON, JSON Lines, YAML or XML as a `DbDataReader`, for `SqlBulkCopy`, `OfficeIMO.Data.Arrow`, `DataTable.Load` or `ExcelSheet.InsertDataReader`.

```csharp
using StructuredDataReader rows = StructuredDataReader.Open("orders.json", new DataConversionOptions { Select = "data.items" });
using var bulk = new SqlBulkCopy(connection) { DestinationTableName = "Orders" };
bulk.WriteToServer(rows);   // columns such as id, customer.name and lines[0].sku
```
- `To{Target}` returns the target's model: `JsonNode` (the BCL's mutable JSON model), `YamlDocument`, `TomlDocument`, `XDocument`, `CsvDocument` or `ExcelDocument`. This keeps the API design's rule without adding a second `ToHtml`-style string exception. Large data goes through `SaveAs{Target}(stream)`, which streams.
- **Breaking:** `ExcelSheet.ToJson`, `TableToJson` and `FromJson` are removed from `OfficeIMO.Excel`. They returned strings through `DataTable`, and `FromJson` breaks the no-`From…` rule. Their replacements are `sheet.ToJson()`/`SaveAsJson` and `JsonNode.SaveAsExcel`/`ToExcel` in the Excel adapter. `MIGRATION.md` gets the mapping. `OfficeIMO.Excel` keeps `System.Text.Json`, which it uses elsewhere.

Files and batches:

```csharp
OfficeConverter.Convert("a.yaml", "a.json");
OfficeConverter.Convert("export.json", "export.xlsx", new OfficeConversionOptions {
    FailOnLoss = true,
    Data = new DataConversionOptions { Select = "data.items" }
});

OfficeBatchReport batch = await OfficeConverter.ConvertAsync(
    Directory.EnumerateFiles("manifests", "*.yaml"), outputDirectory: "json",
    target: OfficeDocumentFormat.Json, new OfficeBatchOptions { MaxParallelism = 8 });
```

- Routes come from the capability catalog: `json-yaml`, `yaml-json`, `json-toml`, …, `xlsx-yaml`, generated over the format set. `csv-xlsx` and `xlsx-csv` stay with `OfficeIMO.Excel.Csv`.
- `OfficeDocumentFormat` gains `Toml` (`Json`, `Yaml`, `Xml` and `Csv` exist). TSV is CSV with the tab profile, selected by the `.tsv` extension.

### Options

One type, because the routes share one mapping. Each property exists for a named scenario:

`OfficeConversionOptions.Data` carries the same object to `OfficeConverter` routes, as the other built-in routes carry theirs.

| Property | Scenario |
|---|---|
| `FailOnLoss` | Pipelines that must not lose data |
| `Select` | API responses and nested documents whose rows sit below the root |
| `Sheet` | One worksheet out of a workbook |
| `Columns` | A fixed column order or subset |
| `Range` | Part of a worksheet, such as an Excel table's range |
| `ArrayColumns` (`Indexed`, `Join`), `ArraySeparator` | Human-readable sheets of tags and lists |
| `InferTypes` | Keep CSV as strings, or type XML values |
| `NestedHeaders` | Turn off unflattening for CSVs with dotted headers that aren't paths (`Ver. 1.2`) |
| `Culture` | European CSV (`1,5`), matching PowerShell's `-Culture` |
| `RootName` | XML root element, and the wrapping key for TOML |
| `JsonLines` | JSON Lines output; a `.jsonl` or `.ndjson` path selects it too |
| `JsonReader`, `JsonWriter`, `YamlReader`, `CsvReader`, `CsvSave`, `XmlReader`, `XmlWriter` | Each format's own reader and writer options, passed through: comments, indentation, encoder, delimiter, limits. DTD processing stays prohibited |

- Writer defaults: JSON is indented with 2 spaces and uses the relaxed encoder, so non-ASCII text stays readable. The output is a JSON file, not HTML. YAML, TOML and CSV use their engines' defaults, including YAML's 1.1-safe quoting.
- Key order is always preserved.

### PowerShell

| Job | Command |
|---|---|
| File to file, one or many | `Convert-OfficeDocument -Path *.yaml -To Json -OutputDirectory json`; `-Options @{ Select = 'data.items' }`, `-Sheet`, `-Culture`, `-NoClobber`, `-Resume` as everywhere |
| Text and objects | `ConvertFrom-OfficeYaml`/`ConvertTo-OfficeYaml` (YAML design), `ConvertFrom-OfficeToml`/`ConvertTo-OfficeToml`, following `ConvertFrom-Json`. PowerShell's own `ConvertTo-Json` covers JSON |
| Tables | `Import-OfficeExcel`/`Import-OfficeCsv` piped into `ConvertTo-OfficeYaml`, and the reverse into `Export-OfficeExcel` |

No new `ConvertTo-`/`ConvertFrom-` file-conversion commands are added, as the API design requires. `Convert-OfficeDocument` writes the `OfficeConversionReport` per file, and loss codes become error IDs under `-FailOnLoss`.

## Performance targets and proof

**Benchmarks:** an opt-in `OfficeIMO.Data.Conversions.Benchmarks` project outside the normal solution, `Category=Performance`. Third-party libraries are referenced only there. Every lane validates its output against the same expected data.

| Lane | Baseline pipeline |
|---|---|
| YAML → JSON, JSON → YAML | YamlDotNet deserialize + `System.Text.Json` serialize, and the reverse; SharpYaml as a second baseline |
| TOML ⇄ JSON | Tomlyn `TomlSerializer`/model + `System.Text.Json` |
| XML ⇄ JSON | Newtonsoft `JsonConvert.SerializeXNode`/`DeserializeXNode` |
| CSV ⇄ JSON | CsvHelper `GetRecords<dynamic>` + `System.Text.Json`, and `JsonDocument` + CsvHelper writer |
| Excel ⇄ JSON | The existing Excel benchmark baselines + `System.Text.Json` |

**Workloads:** many small configuration files (1–10 KB), a 10 MB multi-document Kubernetes stream, a 100 MB and a 1 GB JSON array of records, deep nesting, a 50 MB data-XML export and a 1M-row CSV.

**Targets:**

- Lead every lane in elapsed time and allocation.
- Streaming routes keep peak managed memory flat as input grows from 10 MB to 1 GB (measured, with a fixed ceiling recorded per route).
- Small-file batches are bound by I/O, not converter setup.

Windows, Linux and macOS evidence, as for the YAML engine.

**Correctness:**

- **yaml-test-suite:** YAML → JSON equals `in.json` for every case that has one.
- **toml-test:** TOML → JSON equals the fixtures' tagged JSON after mapping their type tags. The encoder fixtures check JSON → TOML.
- **JSONTestSuite:** every `y_` case converts JSON → YAML → JSON to equal data. Every `n_` case fails.
- **Round-trip identity** where the table above says lossless:
  - JSON ⇄ YAML on the YAML corpus's JSON-compatible files;
  - data XML → JSON → XML (Infoset-equal), checked against Newtonsoft on the same corpus, with each convention difference listed;
  - CSV → JSON → CSV byte-equal under the same dialect.
- **Independent readers accept every output:**
  - JSON: `System.Text.Json` strict and Newtonsoft;
  - YAML: YamlDotNet and SharpYaml, plus PyYAML/ruamel opt-in;
  - TOML: Tomlyn;
  - XML: `XmlReader` with no recovery;
  - CSV: CsvHelper;
  - xlsx: the Open XML validator, plus Excel through REA on Windows opt-in.
- **Loss codes:** one fixture per row of the loss table proves the behavior and the code at the converter boundary, following the test-audit authoring gate.

## Delivery order

| Step | Content | Waits for | Size |
|---|---|---|---|
| D0 | Hub package, event contract, scalar struct, loss collector and catalog codes, `DataConversionOptions`, catalog routes, benchmark and oracle harness | – | M |
| D1 | YAML → JSON (reader and model), JSON source, the yaml-test-suite `in.json` oracle | – (Y1/Y2 shipped) | M |
| D2 | Tree → table and table → tree: CSV ⇄ JSON, CSV/Excel → JSON, JSON → Excel, `Select`, flattening and inference; the Excel adapter; removal of Excel's `ToJson`/`FromJson` | – | L |
| D3 | JSON → YAML, CSV/Excel ⇄ YAML | Y3 `YamlWriter` | S–M |
| D4 | XML ⇄ JSON/YAML/tables | – (can run in parallel with D3) | M |
| D5 | TOML ⇄ everything, toml-test oracle | `OfficeIMO.Toml` reader and writer | S–M |
| D6 | TSV routes | The strict TSV profile in `OfficeIMO.CSV` | S |
| D7 | Front-door registration, PSWriteOffice `Convert-OfficeDocument` routes and `ConvertFrom/To-OfficeToml`, READMEs, `MIGRATION.md`, three-OS benchmark evidence | D1–D6 | M |

**Out of scope:**

- JSON5 and HJSON;
- INI, `.properties` and plist;
- BSON, MessagePack and CBOR;
- JSON Schema validation or inference;
- jq/JSONPath query languages beyond `Select`'s path grammar;
- schema-driven typed mapping (that is `YamlSerializer`/`System.Text.Json`'s job);
- XML Schema-aware typing;
- carrying Excel styles or formulas;
- ODS (reachable later through `OfficeIMO.Excel.OpenDocument`).

## Decisions after review

Decided on 2026-11-10 under the direction to follow best practices:

1. **Package names:** `OfficeIMO.Data.Conversions` and `OfficeIMO.Data.Conversions.Excel`, next to `OfficeIMO.Data.Arrow`. The exception to pairwise adapters is recorded in the [API design](officeimo.api-design.md#packages).
2. **Front door:** the `OfficeIMO` package references both and runs the routes as built-ins.
3. **YAML timestamps:** follow the schema, with no conversion-specific override (see the loss table).
4. **JSON Lines:** supported as a JSON variant (`.jsonl`, `.ndjson`) in D1 and D2. Several documents written as JSON Lines are one line each, and a single array document is one line per item.
5. **`StructuredDataReader`** is public, for DbaClientX bulk copy and `OfficeIMO.Data.Arrow`.
