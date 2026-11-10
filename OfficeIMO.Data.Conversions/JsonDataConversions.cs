using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeIMO.CSV;

namespace OfficeIMO.Data.Conversions;

/// <summary>Converts a <see cref="JsonNode"/> to JSON Lines, CSV and XML.</summary>
public static class JsonDataConversions {
    /// <summary>Writes the value as JSON Lines: one line per item of a root array, otherwise one line.</summary>
    public static OfficeConversionReport SaveAsJsonLines(this JsonNode? node, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(node), 1, DataFormat.Json, DataFormat.JsonLines, path, options);

    /// <summary>Writes the value as JSON Lines to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsJsonLines(this JsonNode? node, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(node), 1, DataFormat.Json, DataFormat.JsonLines, stream, options);

    /// <summary>Returns the rows of the value as CSV: see <see cref="StructuredDataReader"/> for how rows and columns are found.</summary>
    public static CsvDocument ToCsv(this JsonNode? node, DataConversionOptions? options = null) => ToCsv(node, options, out _);

    /// <summary>Returns the rows of the value as CSV, with the conversion report.</summary>
    public static CsvDocument ToCsv(this JsonNode? node, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToCsv(Open(node), 1, DataFormat.Json, options, out report);

    /// <summary>Writes the rows of the value as CSV.</summary>
    public static OfficeConversionReport SaveAsCsv(this JsonNode? node, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(node), 1, DataFormat.Json, DataFormat.Csv, path, options);

    /// <summary>Writes the rows of the value as CSV to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsCsv(this JsonNode? node, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(node), 1, DataFormat.Json, DataFormat.Csv, stream, options);

    /// <summary>Returns the value as XML.</summary>
    public static XDocument ToXml(this JsonNode? node, DataConversionOptions? options = null) => ToXml(node, options, out _);

    /// <summary>Returns the value as XML, with the conversion report.</summary>
    public static XDocument ToXml(this JsonNode? node, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToXml(Open(node), 1, DataFormat.Json, options, out report);

    /// <summary>Writes the value as XML.</summary>
    public static OfficeConversionReport SaveAsXml(this JsonNode? node, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(node), 1, DataFormat.Json, DataFormat.Xml, path, options);

    /// <summary>Writes the value as UTF-8 XML to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsXml(this JsonNode? node, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(node), 1, DataFormat.Json, DataFormat.Xml, stream, options);

    internal static DataConversionPublishing.ReaderFactory Open(JsonNode? node) =>
        context => new JsonEventReader(DataConversionPublishing.Serialize(node), jsonLines: false, context, leaveOpen: false);
}
