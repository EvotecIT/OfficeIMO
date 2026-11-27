using System.Globalization;
using System.IO.Compression;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Write;

internal static partial class XlsbNativePackageWriter {
    /// <summary>Rewrites recorded scalar edits from the owned Excel model while retaining the native record graph.</summary>
    internal static bool TryRewriteOwnedCells(byte[] originalPackage, XlsbWorkbook source,
        IReadOnlyDictionary<string, IReadOnlyDictionary<(int Row, int Column), ExcelDeferredCellEdit>> worksheetEdits,
        long maximumBytes, out byte[] result) {
        result = originalPackage;
        if (source.Sheets.Any(sheet => sheet.State is 1 or 2)
            || source.ActiveTab.GetValueOrDefault() != 0 || source.FirstSheet.GetValueOrDefault() != 0) return false;

        using var input = new MemoryStream(originalPackage, writable: false);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true);
        if (archive.Entries.Any(entry => entry.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase))) {
            throw new NotSupportedException("Native XLSB rewriting is blocked because the source package is digitally signed. Rewriting worksheet parts would invalidate the signature.");
        }

        var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        int matchedSheets = 0;
        foreach (XlsbSheet sheet in source.Sheets) {
            if (!worksheetEdits.TryGetValue(sheet.Name, out var edits)) continue;
            matchedSheets++;
            if (sheet.IsChartSheet || sheet.Tables.Count != 0 || sheet.Comments != null
                || sheet.Cells.Any(cell => cell.FormulaBytes != null) || sheet.FormulaGroups.Count != 0) return false;

            if (!TryOverlayCells(sheet, edits, out IReadOnlyList<XlsbWriteCell> cells)) return false;
            string partName = sheet.PartName
                ?? throw new InvalidDataException($"The source XLSB worksheet '{sheet.Name}' has no resolved package part.");
            ZipArchiveEntry entry = FindEntry(archive, partName)
                ?? throw new InvalidDataException($"The source XLSB worksheet part '{partName}' is missing.");
            byte[] original = ReadEntry(entry, (int)Math.Min(maximumBytes, MaxWorksheetPartBytes));
            byte[] updated = XlsbWorksheetPartWriter.Rewrite(original, cells,
                Array.Empty<XlsbGeneratedRecord>(), rewriteAutoFilter: false,
                Array.Empty<XlsbGeneratedRecord>(), rewriteHyperlinks: false, maximumBytes);
            if (!original.AsSpan().SequenceEqual(updated)) replacements.Add(partName, updated);
        }

        if (matchedSheets != worksheetEdits.Count) return false;
        if (replacements.Count != 0) {
            result = RewritePackage(originalPackage, replacements, source.MaxPartBytes, source.MaxPackageBytes,
                new HashSet<string>(StringComparer.OrdinalIgnoreCase), maximumBytes);
        }
        return true;
    }

    private static bool TryOverlayCells(XlsbSheet sheet,
        IReadOnlyDictionary<(int Row, int Column), ExcelDeferredCellEdit> edits, out IReadOnlyList<XlsbWriteCell> cells) {
        var remaining = new Dictionary<(int Row, int Column), ExcelDeferredCellEdit>(edits);
        var result = new List<XlsbWriteCell>(checked(sheet.Cells.Count + remaining.Count));
        cells = result;
        foreach (XlsbCell source in sheet.Cells) {
            if (!remaining.TryGetValue((source.Row, source.Column), out ExcelDeferredCellEdit edit)) {
                result.Add(XlsbWriteCell.PreserveSource(source));
                continue;
            }

            remaining.Remove((source.Row, source.Column));
            XlsbWriteCell? replacement = ConvertScalarEdit(edit, source.StyleIndex);
            if (replacement == null || source.TextRuns != null) return false;
            bool unchanged = source.Kind switch {
                XlsbCellValueKind.Number => replacement.Kind == XlsbWriteCellKind.Number && Equals(source.Value, replacement.Value),
                XlsbCellValueKind.Text => replacement.Kind == XlsbWriteCellKind.Text && Equals(source.Value, replacement.Value),
                XlsbCellValueKind.Boolean => replacement.Kind == XlsbWriteCellKind.Boolean && Equals(source.Value, replacement.Value),
                _ => false,
            };
            result.Add(unchanged ? XlsbWriteCell.PreserveSource(source) : replacement);
        }

        if (remaining.Count != 0) {
            foreach (ExcelDeferredCellEdit edit in remaining.Values) {
                // A new cell in a formatted row or column can acquire a style from the Excel model.
                if (sheet.Rows.Any(row => row.Row == edit.Row && row.CustomFormat)
                    || sheet.Columns.Any(column => column.FirstColumn <= edit.Column && edit.Column <= column.LastColumn
                        && column.StyleIndex != 0)) return false;
                XlsbWriteCell? replacement = ConvertScalarEdit(edit, style: 0U);
                if (replacement == null) return false;
                result.Add(replacement);
            }
            result.Sort(static (left, right) => {
                int row = left.Row.CompareTo(right.Row);
                return row != 0 ? row : left.Column.CompareTo(right.Column);
            });
        }
        return true;
    }

    private static XlsbWriteCell? ConvertScalarEdit(in ExcelDeferredCellEdit edit, uint style) {
        switch (edit.Kind) {
            case ExcelDeferredCellKind.Text:
                if (edit.Text.IndexOf('\r') >= 0 || edit.Text.IndexOf('\n') >= 0) return null;
                CoerceValueHelper.ValidateSharedStringLength(edit.Text, nameof(edit));
                return new XlsbWriteCell(edit.Row, edit.Column, style, XlsbWriteCellKind.Text, edit.Text);
            case ExcelDeferredCellKind.Boolean:
                return new XlsbWriteCell(edit.Row, edit.Column, style, XlsbWriteCellKind.Boolean, edit.Number != 0D);
            case ExcelDeferredCellKind.Integer:
            case ExcelDeferredCellKind.Number:
                if (double.IsNaN(edit.Number) || double.IsInfinity(edit.Number)) return null;
                return new XlsbWriteCell(edit.Row, edit.Column, style, XlsbWriteCellKind.Number, edit.Number);
            case ExcelDeferredCellKind.Decimal:
                if (!OfficeRoundTripText.TryParseFinite(edit.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) return null;
                return new XlsbWriteCell(edit.Row, edit.Column, style, XlsbWriteCellKind.Number, number);
            default:
                // Date formatting, formula caches and named style operations need the general cell mapping.
                return null;
        }
    }
}
