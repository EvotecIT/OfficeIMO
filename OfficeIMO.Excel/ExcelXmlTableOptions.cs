using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel;

/// <summary>Scalar types admitted by explicit XML table column mappings.</summary>
public enum ExcelXmlTableValueType {
    /// <summary>Literal text, including leading zeroes and formula-looking strings.</summary>
    Text,
    /// <summary>A finite invariant-culture floating-point number.</summary>
    Number,
    /// <summary>An XML Boolean: true, false, 1 or 0.</summary>
    Boolean,
    /// <summary>An XML date/time converted to an Excel date value.</summary>
    DateTime
}

/// <summary>A named direct child element or attribute mapped to one worksheet column.</summary>
public sealed class ExcelXmlTableColumn {
    /// <summary>Worksheet header and exported table column name.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>XML local name of the direct child element or row attribute.</summary>
    public string XmlName { get; set; } = string.Empty;
    /// <summary>Reads and writes a row attribute instead of a child element.</summary>
    public bool IsAttribute { get; set; }
    /// <summary>Explicit scalar conversion; text is the default.</summary>
    public ExcelXmlTableValueType ValueType { get; set; }
    internal ExcelXmlTableColumn Snapshot() => (ExcelXmlTableColumn)MemberwiseClone();
}

/// <summary>Schema and resource policy for XML tables, separate from XML Spreadsheet workbooks.</summary>
public sealed class ExcelXmlTableOptions {
    /// <summary>Required root local name. Default: Table.</summary>
    public string RootElementName { get; set; } = "Table";
    /// <summary>Direct child local name for rows. Default: Row.</summary>
    public string RowElementName { get; set; } = "Row";
    /// <summary>Repeated cell local name in positional mode. Default: Data.</summary>
    public string CellElementName { get; set; } = "Data";
    /// <summary>Namespace for root, row and child elements. Attributes remain unqualified. Default: empty.</summary>
    public string NamespaceUri { get; set; } = string.Empty;
    /// <summary>Explicit named columns. An empty list selects positional repeated-cell mode.</summary>
    public IList<ExcelXmlTableColumn> Columns { get; } = new List<ExcelXmlTableColumn>();
    /// <summary>Maximum input bytes. Default: 16 MiB.</summary>
    public int MaxInputBytes { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum exported bytes. Default: 64 MiB.</summary>
    public int MaxOutputBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Maximum data rows. Default: 100,000.</summary>
    public int MaxRows { get; set; } = 100_000;
    /// <summary>Maximum columns. Default: 256.</summary>
    public int MaxColumns { get; set; } = 256;
    /// <summary>Maximum rectangular table cells, including missing values. Default: 1,000,000.</summary>
    public int MaxCells { get; set; } = 1_000_000;
    /// <summary>Maximum input XML elements. Default: 1,000,000.</summary>
    public int MaxXmlElements { get; set; } = 1_000_000;
    /// <summary>Maximum XML nesting depth. Default: 64.</summary>
    public int MaxXmlDepth { get; set; } = 64;
    /// <summary>Cancellation checked during parsing, conversion and writing.</summary>
    public CancellationToken CancellationToken { get; set; }

    internal ExcelXmlTableOptions Snapshot() {
        if (MaxInputBytes < 1 || MaxOutputBytes < 1 || MaxRows < 1 || MaxRows > ExcelA1.MaxRows || MaxColumns < 1 || MaxColumns > ExcelA1.MaxColumns || MaxCells < 1 || MaxXmlElements < 1 || MaxXmlDepth < 1)
            throw new ArgumentOutOfRangeException(nameof(ExcelXmlTableOptions));
        XmlConvert.VerifyNCName(RootElementName); XmlConvert.VerifyNCName(RowElementName); XmlConvert.VerifyNCName(CellElementName);
        if (NamespaceUri == null) throw new ArgumentNullException(nameof(NamespaceUri));
        if (Columns.Count > MaxColumns) throw new ArgumentException("XML table mappings exceed MaxColumns.");
        var copy = new ExcelXmlTableOptions { RootElementName = RootElementName, RowElementName = RowElementName, CellElementName = CellElementName,
            NamespaceUri = NamespaceUri, MaxInputBytes = MaxInputBytes, MaxOutputBytes = MaxOutputBytes, MaxRows = MaxRows, MaxColumns = MaxColumns,
            MaxCells = MaxCells, MaxXmlElements = MaxXmlElements, MaxXmlDepth = MaxXmlDepth, CancellationToken = CancellationToken };
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase); var selectors = new HashSet<string>(StringComparer.Ordinal);
        foreach (var column in Columns) {
            if (column == null || string.IsNullOrWhiteSpace(column.Name) || !names.Add(column.Name)) throw new ArgumentException("XML table columns require unique nonempty names.");
            XmlConvert.VerifyNCName(column.XmlName);
            if (!selectors.Add((column.IsAttribute ? "@" : "") + column.XmlName) || !Enum.IsDefined(typeof(ExcelXmlTableValueType), column.ValueType)) throw new ArgumentException("Invalid or duplicate XML table column mapping.");
            copy.Columns.Add(column.Snapshot());
        }
        CancellationToken.ThrowIfCancellationRequested(); return copy;
    }
}
