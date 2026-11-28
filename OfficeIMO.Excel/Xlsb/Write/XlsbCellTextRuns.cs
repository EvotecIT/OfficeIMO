using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write;

/// <summary>Reads the formatting runs and stored text for one native save.</summary>
internal sealed class XlsbCellTextRuns {
    private const int Spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
    private readonly OfficeXmlNode _sharedStringTable;
    private readonly XlsbStylesheet? _sourceStylesheet;
    private List<OfficeXmlNode>? _sharedStrings;
    private SharedStringCache? _sharedStringValues;

    /// <param name="sharedStrings">The workbook shared-string table.</param>
    /// <param name="sourceStylesheet">The loaded XLSB stylesheet that native source runs refer to, when rewriting.</param>
    internal XlsbCellTextRuns(SharedStringTable? sharedStrings, XlsbStylesheet? sourceStylesheet = null) {
        _sharedStringTable = OfficeXmlNode.FromSdk(sharedStrings);
        _sourceStylesheet = sourceStylesheet;
    }

    internal static XlsbCellTextRuns ForWorkbook(ExcelDocument document) =>
        new XlsbCellTextRuns(document.WorkbookPartRoot.SharedStringTablePart?.SharedStringTable);

    /// <summary>Reads the stored SpreadsheetML type, retaining the SDK's error for malformed values on its model.</summary>
    internal static string? ReadCellType(OfficeXmlNode cell) {
        string? type = cell.Attribute("t");
        if (type is null or "n" or "b" or "d" or "e" or "s" or "str" or "inlineStr") return type;
        if (cell.SdkElement is Cell sdk) {
            _ = sdk.DataType?.Value;
            return sdk.DataType?.HasValue == true ? type : null;
        }
        throw new FormatException($"Invalid SpreadsheetML cell type '{type}'.");
    }

    /// <summary>Reads a cell's text using the native writer's workbook snapshot.</summary>
    internal string ReadText(ExcelSheet sheet, OfficeXmlNode cell, string? type, string rawValue) {
        switch (type) {
            case "s":
                if (!ExcelSheet.TryParseCellTextSharedStringIndex(rawValue, out int index)) return string.Empty;
                return (_sharedStringValues ??= sheet.BuildCellTextSharedStringSnapshot()).Get(index) ?? string.Empty;
            case "inlineStr":
                OfficeXmlNode inline = cell.Child(Spreadsheet, "is");
                OfficeXmlNode text = inline.Child(Spreadsheet, "t");
                if (text.Exists) return Utilities.ExcelSanitizer.DecodeXString(text.InnerText);
                var decoded = new StringBuilder();
                foreach (OfficeXmlNode run in inline.Children(Spreadsheet, "r")) {
                    decoded.Append(Utilities.ExcelSanitizer.DecodeXString(run.Child(Spreadsheet, "t").InnerText));
                }
                return decoded.ToString();
            case "str":
                return Utilities.ExcelSanitizer.DecodeXString(rawValue);
            case "e" when cell.Attribute("vm") != null && cell.SdkElement is Cell richError:
                return sheet.GetCellText(richError);
            default:
                return rawValue;
        }
    }

    /// <summary>Returns the formatted runs for the existing rich-text font encoder.</summary>
    internal Run[] GetRuns(OfficeXmlNode cell) {
        var runs = ReadRuns(cell);
        if (runs.Count == 0) return Array.Empty<Run>();
        var result = new Run[runs.Count];
        int i = 0;
        foreach (OfficeXmlNode run in runs) {
            result[i++] = run.SdkElement as Run ?? new Run(Encoding.UTF8.GetString(run.OuterXml()));
        }
        return result;
    }

    private OfficeXmlNode.ChildList ReadRuns(OfficeXmlNode cell) {
        OfficeXmlNode inline = cell.Child(Spreadsheet, "is");
        if (inline.Exists) return inline.Children(Spreadsheet, "r");
        if (ReadCellType(cell) != "s"
            || !ExcelSheet.TryParseCellTextSharedStringIndex(cell.Child(Spreadsheet, "v").InnerText, out int index)) return default;
        _sharedStrings ??= _sharedStringTable.Children(Spreadsheet, "si").ToList();
        return index >= 0 && index < _sharedStrings.Count ? _sharedStrings[index].Children(Spreadsheet, "r") : default;
    }

    /// <summary>Whether the cell text carries run formatting.</summary>
    internal bool HasFormattedRuns(OfficeXmlNode cell) {
        foreach (OfficeXmlNode run in ReadRuns(cell)) {
            if (run.Child(Spreadsheet, "rPr").FirstChild.Exists) return true;
        }
        return false;
    }

    internal static bool HasFormattedRuns(Run[] runs) => runs.Any(run => run.RunProperties?.HasChildren == true);

    /// <summary>Whether the cell's runs are the ones projected from the native source cell, so its record can be kept.</summary>
    internal bool MatchesSource(OfficeXmlNode cell, XlsbCell source) {
        InlineString? projected = XlsbWorkbookProjector.CreateRichText(source, _sourceStylesheet);
        if (projected == null) return !HasFormattedRuns(cell);
        var current = ReadRuns(cell).GetEnumerator();
        foreach (OfficeXmlNode expected in OfficeXmlNode.FromSdk(projected).Children(Spreadsheet, "r")) {
            if (!current.MoveNext()
                || !string.Equals(current.Current.ChildText(Spreadsheet, "t"), expected.ChildText(Spreadsheet, "t"), StringComparison.Ordinal)
                || !current.Current.Child(Spreadsheet, "rPr").OuterXml().AsSpan().SequenceEqual(expected.Child(Spreadsheet, "rPr").OuterXml())) {
                return false;
            }
        }
        return !current.MoveNext();
    }
}
