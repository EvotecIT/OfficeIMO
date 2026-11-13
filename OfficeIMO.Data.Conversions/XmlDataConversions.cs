using System.Text.Json.Nodes;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.CSV;
using OfficeIMO.Toml;
using OfficeIMO.Yaml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Converts an <see cref="XDocument"/> to JSON, JSON Lines, CSV or TSV, YAML and TOML with one convention: the root element is the single
/// top-level property, attributes are <c>@name</c>, text beside attributes or elements is <c>#text</c>, repeated
/// elements are arrays and <c>xsi:nil="true"</c> is null.
/// </summary>
public static class XmlDataConversions {
    /// <summary>Returns the document as a JSON model.</summary>
    public static JsonNode? ToJson(this XDocument document, DataConversionOptions? options = null) => ToJson(document, options, out _);

    /// <summary>Returns the document as a JSON model, with the conversion report.</summary>
    public static JsonNode? ToJson(this XDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToJson(Open(document), 1, DataFormat.Xml, options, out report);

    /// <summary>Writes JSON, or JSON Lines when the path ends in <c>.jsonl</c> or <see cref="DataConversionOptions.JsonLines"/> is set.</summary>
    public static OfficeConversionReport SaveAsJson(this XDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Xml, DataFormat.Json, path, options);

    /// <summary>Writes UTF-8 JSON (or JSON Lines) to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsJson(this XDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Xml, DataFormat.Json, stream, options);

    /// <summary>Returns the rows of the document as CSV: see <see cref="StructuredDataReader"/> for how rows and columns are found.</summary>
    public static CsvDocument ToCsv(this XDocument document, DataConversionOptions? options = null) => ToCsv(document, options, out _);

    /// <summary>Returns the rows of the document as CSV, with the conversion report.</summary>
    public static CsvDocument ToCsv(this XDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToCsv(Open(document), 1, DataFormat.Xml, options, out report);

    /// <summary>Writes the rows of the document as CSV, or as TSV when the path ends in <c>.tsv</c> or <c>.tab</c>.</summary>
    public static OfficeConversionReport SaveAsCsv(this XDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Xml, DataFormat.Csv, path, options);

    /// <summary>Writes the rows of the document as CSV to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsCsv(this XDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Xml, DataFormat.Csv, stream, options);

    /// <summary>Returns the document as a YAML document. Strings that a YAML 1.1 reader could take for another type are quoted.</summary>
    public static YamlDocument ToYaml(this XDocument document, DataConversionOptions? options = null) => ToYaml(document, options, out _);

    /// <summary>Returns the document as a YAML document, with the conversion report.</summary>
    public static YamlDocument ToYaml(this XDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToYaml(Open(document), 1, DataFormat.Xml, options, out report);

    /// <summary>Writes the document as YAML.</summary>
    public static OfficeConversionReport SaveAsYaml(this XDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Xml, DataFormat.Yaml, path, options);

    /// <summary>Writes the document as UTF-8 YAML to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsYaml(this XDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Xml, DataFormat.Yaml, stream, options);

    /// <summary>Returns the document as a TOML document. Nulls are omitted and a root that is not an object is wrapped in <see cref="DataConversionOptions.RootName"/>; both are reported.</summary>
    public static TomlDocument ToToml(this XDocument document, DataConversionOptions? options = null) => ToToml(document, options, out _);

    /// <summary>Returns the document as a TOML document, with the conversion report.</summary>
    public static TomlDocument ToToml(this XDocument document, DataConversionOptions? options, out OfficeConversionReport report) =>
        DataConversionPublishing.ToToml(Open(document), 1, DataFormat.Xml, options, out report);

    /// <summary>Writes the document as TOML.</summary>
    public static OfficeConversionReport SaveAsToml(this XDocument document, string path, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Xml, DataFormat.Toml, path, options);

    /// <summary>Writes the document as UTF-8 TOML to a stream, which stays open.</summary>
    public static OfficeConversionReport SaveAsToml(this XDocument document, Stream stream, DataConversionOptions? options = null) =>
        DataConversionPublishing.SaveAs(Open(document), 1, DataFormat.Xml, DataFormat.Toml, stream, options);

    internal static DataConversionPublishing.ReaderFactory Open(XDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return context => {
            XmlReaderSettings settings = XmlEventReader.Settings(context);
            return new XmlEventReader(() => XmlReader.Create(document.CreateReader(), settings), context);
        };
    }
}
