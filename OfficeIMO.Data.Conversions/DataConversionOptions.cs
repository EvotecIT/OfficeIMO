using System.Xml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Settings for structured-data conversions between JSON, JSON Lines, YAML, XML, CSV and Excel. One type serves every
/// route, because the routes share one mapping. Format options pass through to each format's own reader and writer.
/// </summary>
public sealed class DataConversionOptions {
    private string _rootName = "root";
    private string _arraySeparator = ", ";
    private int _maxColumns = 16_384;

    internal static DataConversionOptions Default { get; } = new DataConversionOptions();

    /// <summary>Throws <see cref="OfficeConversionException"/> instead of writing output when any loss is reported.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>
    /// The path to the rows of a table, in the YAML path grammar (<c>data.items</c>, <c>results[0].rows</c>), for sources
    /// whose rows sit below the root, such as API responses. Null takes the rows from the root.
    /// </summary>
    public string? Select { get; set; }

    /// <summary>The worksheet to read from a workbook, by name. Null reads every sheet.</summary>
    public string? Sheet { get; set; }

    /// <summary>The A1 range to read from a worksheet (<c>B2:D40</c>), such as an Excel table's range. Null reads the used range.</summary>
    public string? Range { get; set; }

    /// <summary>The table columns as header paths, in order. Fixes the column set and lets table output be written in one pass.</summary>
    public IReadOnlyList<string>? Columns { get; set; }

    /// <summary>How arrays inside a row become columns. Defaults to <see cref="DataArrayColumns.Indexed"/>.</summary>
    public DataArrayColumns ArrayColumns { get; set; }

    /// <summary>The separator for <see cref="DataArrayColumns.Join"/>. Defaults to <c>", "</c>.</summary>
    public string ArraySeparator {
        get => _arraySeparator;
        set => _arraySeparator = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>
    /// Whether untyped values (CSV fields, XML text) are read as numbers, booleans and date-times when their text is
    /// canonical. Null uses each source's default: on for CSV, off for XML.
    /// </summary>
    public bool? InferTypes { get; set; }

    /// <summary>Whether path-shaped table headers (<c>address.city</c>, <c>tags[0]</c>) rebuild nested data. Defaults to true.</summary>
    public bool NestedHeaders { get; set; } = true;

    /// <summary>The culture for reading and writing CSV numbers and dates. Defaults to the invariant culture.</summary>
    public CultureInfo? Culture { get; set; }

    /// <summary>The root element name for XML output when the data has no single named root. Defaults to <c>root</c>.</summary>
    public string RootName {
        get => _rootName;
        set {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("The root name cannot be empty.", nameof(value));
            _rootName = value;
        }
    }

    /// <summary>The most columns a table may have before the conversion fails. Defaults to 16,384, Excel's limit.</summary>
    public int MaxColumns {
        get => _maxColumns;
        set {
            if (value < 1) throw new ArgumentOutOfRangeException(nameof(value), value, "The limit must be positive.");
            _maxColumns = value;
        }
    }

    /// <summary>Writes JSON output as JSON Lines: one line per document, or per item of a single array document.</summary>
    public bool JsonLines { get; set; }

    /// <summary>JSON parsing settings (comments, trailing commas, depth). Null uses strict RFC 8259 parsing with depth 256.</summary>
    public JsonReaderOptions? JsonReader { get; set; }

    /// <summary>JSON output settings. Null writes indented JSON (2 spaces) that keeps non-ASCII text readable.</summary>
    public JsonWriterOptions? JsonWriter { get; set; }

    /// <summary>YAML parsing settings, including the schema and limits. Null uses the YAML 1.2 core schema.</summary>
    public OfficeIMO.Yaml.YamlReaderOptions? YamlReader { get; set; }

    /// <summary>CSV parsing settings. Null uses the strict profile with a header row.</summary>
    public OfficeIMO.CSV.CsvReaderOptions? CsvReader { get; set; }

    /// <summary>CSV output settings. Null writes comma-separated UTF-8 with a header row and LF line endings.</summary>
    public OfficeIMO.CSV.CsvSaveOptions? CsvSave { get; set; }

    /// <summary>XML parsing settings. DTD processing stays prohibited whatever this sets.</summary>
    public XmlReaderSettings? XmlReader { get; set; }

    /// <summary>XML output settings. Null writes indented UTF-8 XML.</summary>
    public XmlWriterSettings? XmlWriter { get; set; }

    /// <summary>Returns an independent copy.</summary>
    public DataConversionOptions Clone() {
        var copy = (DataConversionOptions)MemberwiseClone();
        copy.Columns = Columns?.ToArray();
        copy.CsvReader = CsvReader?.Clone();
        copy.CsvSave = CsvSave?.Clone();
        copy.XmlReader = XmlReader?.Clone();
        copy.XmlWriter = XmlWriter?.Clone();
        return copy;
    }
}
