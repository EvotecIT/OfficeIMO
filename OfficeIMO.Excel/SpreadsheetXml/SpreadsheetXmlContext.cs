using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Excel.SpreadsheetXml;

internal sealed class SpreadsheetXmlContext {
    internal static readonly XNamespace Ss = "urn:schemas-microsoft-com:office:spreadsheet";
    internal static readonly XNamespace X = "urn:schemas-microsoft-com:office:excel";
    internal readonly ExcelSpreadsheetXmlOptions Options;
    private readonly Dictionary<string, OfficeConversionFidelityDiagnostic> _diagnostics = new(StringComparer.Ordinal);
    private int _cells;
    private int _names;
    private readonly Dictionary<string, HashSet<uint?>> _definedNameScopes = new(StringComparer.OrdinalIgnoreCase);
    internal readonly Dictionary<uint, uint> DateStyles = new();

    internal SpreadsheetXmlContext(ExcelSpreadsheetXmlOptions options) => Options = options;
    internal void Cell() {
        Options.CancellationToken.ThrowIfCancellationRequested();
        if (_cells >= Options.MaxCells) throw new InvalidDataException("XML Spreadsheet exceeds MaxCells.");
        _cells++;
    }
    internal void DefinedName(string name, uint? scope) {
        Options.CancellationToken.ThrowIfCancellationRequested();
        if (_names >= Options.MaxNames) throw new InvalidDataException("XML Spreadsheet exceeds MaxNames.");
        if (!_definedNameScopes.TryGetValue(name, out HashSet<uint?>? scopes))
            _definedNameScopes[name] = scopes = new HashSet<uint?>();
        if (!scopes.Add(scope)) throw new InvalidDataException("Duplicate XML Spreadsheet defined name.");
        _names++;
    }
    internal void Omitted(string feature) {
        if (_diagnostics.ContainsKey(feature)) return;
        if (_diagnostics.Count >= 127) feature = "additional unsupported feature kinds";
        _diagnostics[feature] = new OfficeConversionFidelityDiagnostic("EXCEL_XML_OMITTED",
            "The data interchange profile omits " + feature + ".", OfficeConversionLossKind.Omission,
            "OfficeIMO.Excel.SpreadsheetXml");
    }
    internal void Approximate(string feature) {
        if (_diagnostics.ContainsKey(feature)) return;
        _diagnostics[feature] = new OfficeConversionFidelityDiagnostic("EXCEL_XML_APPROXIMATION", feature, OfficeConversionLossKind.Approximation, "OfficeIMO.Excel.SpreadsheetXml");
    }
    internal ExcelSpreadsheetXmlReport Report() => new(_diagnostics.Values);
    internal void Inspect(XElement element, string[] children, string[] attributes) {
        foreach (XElement child in element.Elements()) {
            if (!(child.Name.Namespace == Ss && children.Contains(child.Name.LocalName)) && !children.Contains(child.Name.ToString())) Omitted(child.Name.ToString());
        }
        foreach (XAttribute attribute in element.Attributes()) {
            if (attribute.IsNamespaceDeclaration || attribute.Name == XNamespace.Xml + "space") continue;
            if (!(attribute.Name.Namespace == Ss && attributes.Contains(attribute.Name.LocalName)) && !attributes.Contains(attribute.Name.ToString()))
                Omitted("attribute " + attribute.Name);
        }
    }
    internal static int Index(XElement element, string name, int fallback, int maximum) {
        string? text = (string?)element.Attribute(Ss + name);
        if (text == null) return fallback;
        if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int value) || value < 1 || value > maximum)
            throw new InvalidDataException("Invalid XML Spreadsheet " + name + ".");
        return value;
    }
    internal static void Coordinate(int row, int column) {
        if (row < 1 || row > 65536 || column < 1 || column > 256)
            throw new InvalidDataException("XML Spreadsheet 2003 supports at most 65,536 rows and 256 columns.");
    }
}
