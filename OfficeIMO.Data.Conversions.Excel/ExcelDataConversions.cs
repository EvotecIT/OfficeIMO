using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Toml;
using OfficeIMO.Yaml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Converts worksheets and workbooks to JSON, JSON Lines, YAML, TOML and XML, and the rows of JSON, YAML, TOML and XML
/// data to workbooks. A worksheet is an array of row objects read with first-row headers; a workbook with several sheets is an object of
/// such arrays, keyed by sheet name. Formulas contribute their cached values.
/// </summary>
public static class ExcelDataConversions {
    // ------------------------------------------------------------------ worksheets and workbooks to data

    /// <summary>Returns the worksheet's rows as a JSON array.</summary>
    public static JsonNode? ToJson(this ExcelSheet sheet, DataConversionOptions? options = null) => ToJson(sheet, options, out _);

    /// <summary>Returns the worksheet's rows as a JSON array, with the conversion report.</summary>
    public static JsonNode? ToJson(this ExcelSheet sheet, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToJson(Open(sheet), 1, DataFormat.Csv, options, out report).Rename(ref report);

    /// <summary>Writes the worksheet's rows as JSON, or JSON Lines when the path ends in <c>.jsonl</c>.</summary>
    public static OfficeConversionReport SaveAsJson(this ExcelSheet sheet, string path, DataConversionOptions? options = null) =>
        Excel(DataConversionPublishing.SaveAs(Open(sheet), 1, DataFormat.Csv, DataFormat.Json, path, options));

    /// <summary>Writes the worksheet's rows as UTF-8 JSON (or JSON Lines) to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsJson(this ExcelSheet sheet, Stream stream, DataConversionOptions? options = null) =>
        Excel(DataConversionPublishing.SaveAs(Open(sheet), 1, DataFormat.Csv, DataFormat.Json, stream, options));

    /// <summary>Returns the worksheet's rows as XML.</summary>
    public static XDocument ToXml(this ExcelSheet sheet, DataConversionOptions? options = null) =>
        DataConversionPublishing.ToXml(Open(sheet), 1, DataFormat.Csv, options, out _);

    /// <summary>Writes the worksheet's rows as XML.</summary>
    public static OfficeConversionReport SaveAsXml(this ExcelSheet sheet, string path, DataConversionOptions? options = null) =>
        Excel(DataConversionPublishing.SaveAs(Open(sheet), 1, DataFormat.Csv, DataFormat.Xml, path, options));

    /// <summary>
    /// Returns the workbook as JSON: an array of rows for one sheet (or <see cref="DataConversionOptions.Sheet"/>), or an
    /// object of arrays keyed by sheet name.
    /// </summary>
    public static JsonNode? ToJson(this ExcelDocument document, DataConversionOptions? options = null) => ToJson(document, options, out _);

    /// <summary>Returns the workbook as JSON, with the conversion report.</summary>
    public static JsonNode? ToJson(this ExcelDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToJson(Open(document, options), 1, DataFormat.Csv, options, out report).Rename(ref report);

    /// <summary>Writes the workbook as JSON, or JSON Lines when the path ends in <c>.jsonl</c>.</summary>
    public static OfficeConversionReport SaveAsJson(this ExcelDocument document, string path, DataConversionOptions? options = null) =>
        Excel(DataConversionPublishing.SaveAs(Open(document, options), 1, DataFormat.Csv, DataFormat.Json, path, options));

    /// <summary>Writes the workbook as UTF-8 JSON (or JSON Lines) to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsJson(this ExcelDocument document, Stream stream, DataConversionOptions? options = null) =>
        Excel(DataConversionPublishing.SaveAs(Open(document, options), 1, DataFormat.Csv, DataFormat.Json, stream, options));

    /// <summary>Writes the workbook as XML.</summary>
    public static OfficeConversionReport SaveAsXml(this ExcelDocument document, string path, DataConversionOptions? options = null) =>
        Excel(DataConversionPublishing.SaveAs(Open(document, options), 1, DataFormat.Csv, DataFormat.Xml, path, options));

    /// <summary>Returns the worksheet's rows as a YAML sequence of mappings.</summary>
    public static YamlDocument ToYaml(this ExcelSheet sheet, DataConversionOptions? options = null) =>
        DataConversionPublishing.ToYaml(Open(sheet), 1, DataFormat.Csv, options, out _);

    /// <summary>Writes the worksheet's rows as YAML.</summary>
    public static OfficeConversionReport SaveAsYaml(this ExcelSheet sheet, string path, DataConversionOptions? options = null) =>
        Excel(DataConversionPublishing.SaveAs(Open(sheet), 1, DataFormat.Csv, DataFormat.Yaml, path, options));

    /// <summary>Returns the workbook as YAML: a sequence of rows for one sheet, or a mapping of sequences keyed by sheet name.</summary>
    public static YamlDocument ToYaml(this ExcelDocument document, DataConversionOptions? options = null) =>
        DataConversionPublishing.ToYaml(Open(document, options), 1, DataFormat.Csv, options, out _);

    /// <summary>Writes the workbook as YAML.</summary>
    public static OfficeConversionReport SaveAsYaml(this ExcelDocument document, string path, DataConversionOptions? options = null) =>
        Excel(DataConversionPublishing.SaveAs(Open(document, options), 1, DataFormat.Csv, DataFormat.Yaml, path, options));

    /// <summary>Returns the worksheet's rows as a TOML document: an array of tables under <see cref="DataConversionOptions.RootName"/>.</summary>
    public static TomlDocument ToToml(this ExcelSheet sheet, DataConversionOptions? options = null) => ToToml(sheet, options, out _);

    /// <summary>Returns the worksheet's rows as a TOML document, with the conversion report.</summary>
    public static TomlDocument ToToml(this ExcelSheet sheet, DataConversionOptions? options, out OfficeConversionReport report) {
        TomlDocument document = DataConversionPublishing.ToToml(Open(sheet), 1, DataFormat.Csv, options, out report);
        report = Excel(report);
        return document;
    }

    /// <summary>Writes the worksheet's rows as TOML.</summary>
    public static OfficeConversionReport SaveAsToml(this ExcelSheet sheet, string path, DataConversionOptions? options = null) =>
        Excel(DataConversionPublishing.SaveAs(Open(sheet), 1, DataFormat.Csv, DataFormat.Toml, path, options));

    /// <summary>Writes the worksheet's rows as UTF-8 TOML to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsToml(this ExcelSheet sheet, Stream stream, DataConversionOptions? options = null) =>
        Excel(DataConversionPublishing.SaveAs(Open(sheet), 1, DataFormat.Csv, DataFormat.Toml, stream, options));

    /// <summary>
    /// Returns the workbook as TOML: an array of tables under <see cref="DataConversionOptions.RootName"/> for one sheet,
    /// or one array of tables per sheet, keyed by sheet name.
    /// </summary>
    public static TomlDocument ToToml(this ExcelDocument document, DataConversionOptions? options = null) => ToToml(document, options, out _);

    /// <summary>Returns the workbook as TOML, with the conversion report.</summary>
    public static TomlDocument ToToml(this ExcelDocument document, DataConversionOptions? options, out OfficeConversionReport report) {
        TomlDocument toml = DataConversionPublishing.ToToml(Open(document, options), 1, DataFormat.Csv, options, out report);
        report = Excel(report);
        return toml;
    }

    /// <summary>Writes the workbook as TOML.</summary>
    public static OfficeConversionReport SaveAsToml(this ExcelDocument document, string path, DataConversionOptions? options = null) =>
        Excel(DataConversionPublishing.SaveAs(Open(document, options), 1, DataFormat.Csv, DataFormat.Toml, path, options));

    /// <summary>Writes the workbook as UTF-8 TOML to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsToml(this ExcelDocument document, Stream stream, DataConversionOptions? options = null) =>
        Excel(DataConversionPublishing.SaveAs(Open(document, options), 1, DataFormat.Csv, DataFormat.Toml, stream, options));

    // ------------------------------------------------------------------ data to workbooks

    /// <summary>Returns the rows of a JSON value as a workbook: one sheet, or one per member when every root member is an array.</summary>
    public static ExcelDocument ToExcel(this JsonNode? node, DataConversionOptions? options = null) => ToExcel(node, options, out _);

    /// <summary>Returns the rows of a JSON value as a workbook, with the conversion report.</summary>
    public static ExcelDocument ToExcel(this JsonNode? node, DataConversionOptions? options, out OfficeConversionReport report) =>
        LoadWorkbook(stream => Write(JsonDataConversions.Open(node), 1, DataFormat.Json, stream, options), out report);

    /// <summary>Writes the rows of a JSON value as an XLSX workbook.</summary>
    public static OfficeConversionReport SaveAsExcel(this JsonNode? node, string path, DataConversionOptions? options = null) =>
        DataConversionEngine.SaveToPath(path, stream => Write(JsonDataConversions.Open(node), 1, DataFormat.Json, stream, options));

    /// <summary>Writes the rows of a JSON value as an XLSX workbook to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsExcel(this JsonNode? node, Stream stream, DataConversionOptions? options = null) =>
        Write(JsonDataConversions.Open(node), 1, DataFormat.Json, stream, options);

    /// <summary>Returns the rows of a YAML document as a workbook. Each document of a stream is a row.</summary>
    public static ExcelDocument ToExcel(this YamlDocument document, DataConversionOptions? options = null) => ToExcel(document, options, out _);

    /// <summary>Returns the rows of a YAML document as a workbook, with the conversion report.</summary>
    public static ExcelDocument ToExcel(this YamlDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        LoadWorkbook(stream => Write(YamlDataConversions.Open(document), YamlDataConversions.Count(document), DataFormat.Yaml, stream, options), out report);

    /// <summary>Writes the rows of a YAML document as an XLSX workbook.</summary>
    public static OfficeConversionReport SaveAsExcel(this YamlDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionEngine.SaveToPath(path, stream => Write(YamlDataConversions.Open(document), YamlDataConversions.Count(document), DataFormat.Yaml, stream, options));

    /// <summary>Writes the rows of a YAML document as an XLSX workbook to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsExcel(this YamlDocument document, Stream stream, DataConversionOptions? options = null) =>
        Write(YamlDataConversions.Open(document), YamlDataConversions.Count(document), DataFormat.Yaml, stream, options);

    /// <summary>Returns the rows of an XML document as a workbook.</summary>
    public static ExcelDocument ToExcel(this XDocument document, DataConversionOptions? options = null) => ToExcel(document, options, out _);

    /// <summary>Returns the rows of an XML document as a workbook, with the conversion report.</summary>
    public static ExcelDocument ToExcel(this XDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        LoadWorkbook(stream => Write(XmlDataConversions.Open(document), 1, DataFormat.Xml, stream, options), out report);

    /// <summary>Writes the rows of an XML document as an XLSX workbook.</summary>
    public static OfficeConversionReport SaveAsExcel(this XDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionEngine.SaveToPath(path, stream => Write(XmlDataConversions.Open(document), 1, DataFormat.Xml, stream, options));

    /// <summary>Writes the rows of an XML document as an XLSX workbook to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsExcel(this XDocument document, Stream stream, DataConversionOptions? options = null) =>
        Write(XmlDataConversions.Open(document), 1, DataFormat.Xml, stream, options);

    /// <summary>Returns the rows of a TOML document as a workbook, such as the tables of an <c>[[array]]</c> chosen with <see cref="DataConversionOptions.Select"/>.</summary>
    public static ExcelDocument ToExcel(this TomlDocument document, DataConversionOptions? options = null) => ToExcel(document, options, out _);

    /// <summary>Returns the rows of a TOML document as a workbook, with the conversion report.</summary>
    public static ExcelDocument ToExcel(this TomlDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        LoadWorkbook(stream => Write(TomlDataConversions.Open(document), 1, DataFormat.Toml, stream, options), out report);

    /// <summary>Writes the rows of a TOML document as an XLSX workbook.</summary>
    public static OfficeConversionReport SaveAsExcel(this TomlDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionEngine.SaveToPath(path, stream => Write(TomlDataConversions.Open(document), 1, DataFormat.Toml, stream, options));

    /// <summary>Writes the rows of a TOML document as an XLSX workbook to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsExcel(this TomlDocument document, Stream stream, DataConversionOptions? options = null) =>
        Write(TomlDataConversions.Open(document), 1, DataFormat.Toml, stream, options);

    // ------------------------------------------------------------------ plumbing

    internal static DataConversionPublishing.ReaderFactory Open(ExcelSheet sheet) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        return context => new TableEventReader(() => sheet.CreateDataReader(new ExcelReaderOptions { A1Range = context.Options.Range }), infer: context.Options.InferTypes ?? false, positional: false, null, context);
    }

    internal static DataConversionPublishing.ReaderFactory Open(ExcelDocument document, DataConversionOptions? options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var reader = new ExcelReaderOptions { SheetName = options?.Sheet, A1Range = options?.Range };
        return context => new TableEventReader(() => document.CreateDataReader(reader), context.Options.InferTypes ?? false,
            positional: false, options?.Sheet == null ? SheetNames(document) : null, context);
    }

    internal static IReadOnlyList<string> SheetNames(ExcelDocument document) => document.Sheets.Select(sheet => sheet.Name).ToArray();

    /// <summary>Opens a workbook file or stream for the file routes.</summary>
    internal static DataEventReader OpenWorkbook(Stream input, DataConversionContext context) {
        var reader = new ExcelReaderOptions { SheetName = context.Options.Sheet, A1Range = context.Options.Range };
        IReadOnlyList<string>? names = null;
        long start = input.Position;
        if (context.Options.Sheet == null) {
            using ExcelWorkbookDataReader probe = ExcelDocument.OpenReader(input, reader);
            names = probe.SheetNames;
            input.Position = start;
        }

        return new TableEventReader(() => {
            input.Position = start;
            return ExcelDocument.OpenReader(input, reader);
        }, context.Options.InferTypes ?? false, positional: false, names, context);
    }

    /// <summary>Writes the rows of a tree source as a workbook through the streaming writer.</summary>
    internal static OfficeConversionReport Write(DataConversionPublishing.ReaderFactory open, long? documents, DataFormat source, Stream output, DataConversionOptions? options) {
        if (output == null) throw new ArgumentNullException(nameof(output));
        var context = new DataConversionContext(options ?? DataConversionOptions.Default);
        using var rows = new StructuredDataReader(open, documents, context, text: false, owned: null);
        return WriteRows(rows, output, DataFormats.Name(source));
    }

    internal static OfficeConversionReport WriteRows(StructuredDataReader rows, Stream output, string source) {
        DataConversionContext context = rows.Context;
        var cells = new ExcelValueReader(rows);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (ExcelWorkbookWriter writer = ExcelDocument.CreateWriter(output, leaveOpen: true)) {
            do {
                writer.WriteTable(SheetName(rows.TableName.Length == 0 ? context.Options.Sheet ?? "Data" : rows.TableName, used), cells, context.CancellationToken);
            } while (rows.NextResult());

            context.ThrowIfLossNotAllowed(source, "Excel");
        }

        return context.CreateReport(source, "Excel");
    }

    // Excel sheet names: at most 31 characters, none of []:*?/\, unique without regard to case.
    private static string SheetName(string name, HashSet<string> used) {
        var builder = new System.Text.StringBuilder();
        foreach (char c in name) builder.Append(c is '[' or ']' or ':' or '*' or '?' or '/' or '\\' ? '_' : c);
        string candidate = builder.Length == 0 ? "Data" : builder.ToString();
        if (candidate.Length > 31) candidate = candidate.Substring(0, 31);
        string unique = candidate;
        for (int i = 2; !used.Add(unique); i++) {
            string suffix = " (" + i.ToString(CultureInfo.InvariantCulture) + ")";
            unique = (candidate.Length + suffix.Length > 31 ? candidate.Substring(0, 31 - suffix.Length) : candidate) + suffix;
        }

        return unique;
    }

    private static ExcelDocument LoadWorkbook(Func<Stream, OfficeConversionReport> write, out OfficeConversionReport report) {
        var buffer = new MemoryStream();
        report = write(buffer);
        buffer.Position = 0;
        return ExcelDocument.Load(buffer);
    }

    private static JsonNode? Rename(this JsonNode? node, ref OfficeConversionReport report) {
        report = Excel(report);
        return node;
    }

    // Worksheet sources run through the table reader the CSV route uses; reports name Excel as the source.
    private static OfficeConversionReport Excel(OfficeConversionReport report) =>
        new OfficeConversionReport("Excel", report.Target, report.OutputPath, report.ItemsWritten, report.Elapsed, report.Diagnostics);
}
