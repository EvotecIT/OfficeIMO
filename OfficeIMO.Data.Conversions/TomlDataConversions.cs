using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeIMO.CSV;
using OfficeIMO.Toml;
using OfficeIMO.Yaml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Converts a loaded <see cref="TomlDocument"/> to JSON, JSON Lines, YAML, XML and CSV or TSV. Tables are objects, arrays
/// of tables are arrays, and integers, floats and booleans keep their types. Date-times become RFC 3339 text in JSON,
/// YAML and XML, which is reported; YAML and XML carry the comments.
/// </summary>
public static class TomlDataConversions {
    /// <summary>Returns the document as a JSON model.</summary>
    public static JsonNode? ToJson(this TomlDocument document, DataConversionOptions? options = null) => ToJson(document, options, out _);

    /// <summary>Returns the document as a JSON model, with the conversion report.</summary>
    public static JsonNode? ToJson(this TomlDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToJson(Open(document), 1, DataFormat.Toml, options, out report);

    /// <summary>Writes JSON, or JSON Lines when the path ends in <c>.jsonl</c> or <see cref="DataConversionOptions.JsonLines"/> is set.</summary>
    public static OfficeConversionReport SaveAsJson(this TomlDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Toml, DataFormat.Json, path, options);

    /// <summary>Writes UTF-8 JSON (or JSON Lines) to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsJson(this TomlDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Toml, DataFormat.Json, stream, options);

    /// <summary>Returns the rows of the document as CSV: see <see cref="StructuredDataReader"/> for how rows and columns are found.</summary>
    public static CsvDocument ToCsv(this TomlDocument document, DataConversionOptions? options = null) => ToCsv(document, options, out _);

    /// <summary>Returns the rows of the document as CSV, with the conversion report.</summary>
    public static CsvDocument ToCsv(this TomlDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToCsv(Open(document), 1, DataFormat.Toml, options, out report);

    /// <summary>Writes the rows of the document as CSV, or as TSV when the path ends in <c>.tsv</c> or <c>.tab</c>.</summary>
    public static OfficeConversionReport SaveAsCsv(this TomlDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Toml, DataFormat.Csv, path, options);

    /// <summary>Writes the rows of the document as CSV to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsCsv(this TomlDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Toml, DataFormat.Csv, stream, options);

    /// <summary>Returns the document as XML.</summary>
    public static XDocument ToXml(this TomlDocument document, DataConversionOptions? options = null) => ToXml(document, options, out _);

    /// <summary>Returns the document as XML, with the conversion report.</summary>
    public static XDocument ToXml(this TomlDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToXml(Open(document), 1, DataFormat.Toml, options, out report);

    /// <summary>Writes the document as XML.</summary>
    public static OfficeConversionReport SaveAsXml(this TomlDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Toml, DataFormat.Xml, path, options);

    /// <summary>Writes the document as UTF-8 XML to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsXml(this TomlDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Toml, DataFormat.Xml, stream, options);

    /// <summary>Returns the document as a YAML document, keeping its comments.</summary>
    public static YamlDocument ToYaml(this TomlDocument document, DataConversionOptions? options = null) => ToYaml(document, options, out _);

    /// <summary>Returns the document as a YAML document, with the conversion report.</summary>
    public static YamlDocument ToYaml(this TomlDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToYaml(Open(document), 1, DataFormat.Toml, options, out report);

    /// <summary>Writes the document as YAML.</summary>
    public static OfficeConversionReport SaveAsYaml(this TomlDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Toml, DataFormat.Yaml, path, options);

    /// <summary>Writes the document as UTF-8 YAML to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsYaml(this TomlDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Toml, DataFormat.Yaml, stream, options);

    internal static DataConversionPublishing.ReaderFactory Open(TomlDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return context => new TomlModelEventReader(document, context);
    }
}
