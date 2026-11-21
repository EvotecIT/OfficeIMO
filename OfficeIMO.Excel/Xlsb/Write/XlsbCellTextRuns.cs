using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write;

/// <summary>Reads the formatting runs of cell text for one save, resolving shared strings from a single snapshot.</summary>
internal sealed class XlsbCellTextRuns {
    private readonly SharedStringTable? _sharedStringTable;
    private readonly XlsbStylesheet? _sourceStylesheet;
    private SharedStringItem[]? _sharedStrings;

    /// <param name="sharedStrings">The workbook shared-string table.</param>
    /// <param name="sourceStylesheet">The loaded XLSB stylesheet that native source runs refer to, when rewriting.</param>
    internal XlsbCellTextRuns(SharedStringTable? sharedStrings, XlsbStylesheet? sourceStylesheet = null) {
        _sharedStringTable = sharedStrings;
        _sourceStylesheet = sourceStylesheet;
    }

    /// <summary>Returns the runs of an inline or shared-string cell; empty for other cells.</summary>
    internal Run[] GetRuns(Cell cell) {
        if (cell.InlineString != null) return cell.InlineString.Elements<Run>().ToArray();
        if (cell.DataType?.Value != CellValues.SharedString
            || !ExcelSheet.TryParseCellTextSharedStringIndex(cell.CellValue?.InnerText, out int index)) return Array.Empty<Run>();
        _sharedStrings ??= _sharedStringTable?.Elements<SharedStringItem>().ToArray() ?? Array.Empty<SharedStringItem>();
        return index >= 0 && index < _sharedStrings.Length ? _sharedStrings[index].Elements<Run>().ToArray() : Array.Empty<Run>();
    }

    /// <summary>Whether the cell text carries run formatting.</summary>
    internal bool HasFormattedRuns(Cell cell) => HasFormattedRuns(GetRuns(cell));

    internal static bool HasFormattedRuns(Run[] runs) => runs.Any(run => run.RunProperties?.HasChildren == true);

    /// <summary>Whether the cell's runs are the ones projected from the native source cell, so its record can be kept.</summary>
    internal bool MatchesSource(Cell cell, XlsbCell source) {
        Run[] current = GetRuns(cell);
        InlineString? projected = XlsbWorkbookProjector.CreateRichText(source, _sourceStylesheet);
        if (projected == null) return !HasFormattedRuns(current);
        Run[] expected = projected.Elements<Run>().ToArray();
        if (current.Length != expected.Length) return false;
        for (int index = 0; index < current.Length; index++) {
            if (!string.Equals(current[index].Text?.Text, expected[index].Text?.Text, StringComparison.Ordinal)
                || !string.Equals(current[index].RunProperties?.OuterXml, expected[index].RunProperties?.OuterXml, StringComparison.Ordinal)) {
                return false;
            }
        }
        return true;
    }
}
