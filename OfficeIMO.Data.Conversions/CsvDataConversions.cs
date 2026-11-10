using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeIMO.CSV;
using OfficeIMO.Yaml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Converts a <see cref="CsvDocument"/> to JSON, JSON Lines and XML: an array of row objects. Column types are inferred
/// from canonical spellings only (<c>007</c> stays text), and path headers (<c>address.city</c>, <c>tags[0]</c>) rebuild
/// nested data.
/// </summary>
public static class CsvDataConversions {
    /// <summary>Returns the rows as a JSON array.</summary>
    public static JsonNode? ToJson(this CsvDocument document, DataConversionOptions? options = null) => ToJson(document, options, out _);

    /// <summary>Returns the rows as a JSON array, with the conversion report.</summary>
    public static JsonNode? ToJson(this CsvDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToJson(Open(document), 1, DataFormat.Csv, options, out report);

    /// <summary>Writes the rows as JSON, or as JSON Lines when the path ends in <c>.jsonl</c> or <see cref="DataConversionOptions.JsonLines"/> is set.</summary>
    public static OfficeConversionReport SaveAsJson(this CsvDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Csv, DataFormat.Json, path, options);

    /// <summary>Writes the rows as UTF-8 JSON (or JSON Lines) to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsJson(this CsvDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Csv, DataFormat.Json, stream, options);

    /// <summary>Returns the rows as XML.</summary>
    public static XDocument ToXml(this CsvDocument document, DataConversionOptions? options = null) => ToXml(document, options, out _);

    /// <summary>Returns the rows as XML, with the conversion report.</summary>
    public static XDocument ToXml(this CsvDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToXml(Open(document), 1, DataFormat.Csv, options, out report);

    /// <summary>Writes the rows as XML.</summary>
    public static OfficeConversionReport SaveAsXml(this CsvDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Csv, DataFormat.Xml, path, options);

    /// <summary>Writes the rows as UTF-8 XML to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsXml(this CsvDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Csv, DataFormat.Xml, stream, options);

    /// <summary>Returns the rows as a YAML document. Strings that a YAML 1.1 reader could take for another type are quoted.</summary>
    public static YamlDocument ToYaml(this CsvDocument document, DataConversionOptions? options = null) => ToYaml(document, options, out _);

    /// <summary>Returns the rows as a YAML document, with the conversion report.</summary>
    public static YamlDocument ToYaml(this CsvDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToYaml(Open(document), 1, DataFormat.Csv, options, out report);

    /// <summary>Writes the rows as YAML.</summary>
    public static OfficeConversionReport SaveAsYaml(this CsvDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Csv, DataFormat.Yaml, path, options);

    /// <summary>Writes the rows as UTF-8 YAML to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsYaml(this CsvDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Csv, DataFormat.Yaml, stream, options);

    internal static DataConversionPublishing.ReaderFactory Open(CsvDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return context => new TableEventReader(() => document.CreateDataReader(), context.Options.InferTypes ?? true, positional: false, null, context);
    }
}
