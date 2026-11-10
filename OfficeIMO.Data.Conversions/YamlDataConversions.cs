using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeIMO.CSV;
using OfficeIMO.Yaml;

namespace OfficeIMO.Data.Conversions;

/// <summary>Converts a loaded <see cref="YamlDocument"/> to JSON, JSON Lines, CSV and XML.</summary>
public static class YamlDataConversions {
    /// <summary>
    /// Returns the document as a JSON model. A stream of several documents becomes a JSON array of them; an empty stream
    /// is JSON <c>null</c>. Comments, custom tags and other YAML-only content are reported.
    /// </summary>
    public static JsonNode? ToJson(this YamlDocument document, DataConversionOptions? options = null) => ToJson(document, options, out _);

    /// <summary>Returns the document as a JSON model, with the conversion report.</summary>
    public static JsonNode? ToJson(this YamlDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToJson(Open(document), Count(document), DataFormat.Yaml, options, out report);

    /// <summary>Writes JSON, or JSON Lines when the path ends in <c>.jsonl</c> or <see cref="DataConversionOptions.JsonLines"/> is set.</summary>
    public static OfficeConversionReport SaveAsJson(this YamlDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), Count(document), DataFormat.Yaml, DataFormat.Json, path, options);

    /// <summary>Writes UTF-8 JSON (or JSON Lines) to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsJson(this YamlDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), Count(document), DataFormat.Yaml, DataFormat.Json, stream, options);

    /// <summary>Returns the rows of the document as CSV: see <see cref="StructuredDataReader"/> for how rows and columns are found.</summary>
    public static CsvDocument ToCsv(this YamlDocument document, DataConversionOptions? options = null) => ToCsv(document, options, out _);

    /// <summary>Returns the rows of the document as CSV, with the conversion report.</summary>
    public static CsvDocument ToCsv(this YamlDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToCsv(Open(document), Count(document), DataFormat.Yaml, options, out report);

    /// <summary>Writes the rows of the document as CSV.</summary>
    public static OfficeConversionReport SaveAsCsv(this YamlDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), Count(document), DataFormat.Yaml, DataFormat.Csv, path, options);

    /// <summary>Writes the rows of the document as CSV to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsCsv(this YamlDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), Count(document), DataFormat.Yaml, DataFormat.Csv, stream, options);

    /// <summary>Returns the document as XML.</summary>
    public static XDocument ToXml(this YamlDocument document, DataConversionOptions? options = null) => ToXml(document, options, out _);

    /// <summary>Returns the document as XML, with the conversion report.</summary>
    public static XDocument ToXml(this YamlDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToXml(Open(document), Count(document), DataFormat.Yaml, options, out report);

    /// <summary>Writes the document as XML.</summary>
    public static OfficeConversionReport SaveAsXml(this YamlDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), Count(document), DataFormat.Yaml, DataFormat.Xml, path, options);

    /// <summary>Writes the document as UTF-8 XML to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsXml(this YamlDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), Count(document), DataFormat.Yaml, DataFormat.Xml, stream, options);

    internal static DataConversionPublishing.ReaderFactory Open(YamlDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return context => new YamlModelEventReader(document, context);
    }

    internal static long Count(YamlDocument document) => document.Documents.Count;
}
