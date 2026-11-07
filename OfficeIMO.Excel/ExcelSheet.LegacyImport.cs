using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Retains native custom height zero and explicit visibility independently of worksheet defaults.</summary>
    internal void SetImportedRowGeometry(int rowIndex, double? customHeight, bool hidden) {
        if (rowIndex <= 0) return;
        WriteLockConditional(() => {
            Row row = GetOrCreateRowElement(GetOrCreateSheetData(), rowIndex);
            if (customHeight.HasValue) {
                if (customHeight.Value == 0d) {
                    row.Height = 0d;
                    row.CustomHeight = true;
                } else {
                    SetRowHeightCore(row, customHeight.Value, roundToHundredths: false);
                }
                UpdateSheetFormat(customHeightSet: row.CustomHeight?.Value == true);
            }
            row.Hidden = hidden;
        });
    }

    /// <summary>Materializes bounded legacy column ranges once, retaining widths, visibility, outlines and styles.</summary>
    internal void SetImportedColumnLayouts(IReadOnlyList<LegacyXlsColumnLayout> source, Func<ushort, uint?> resolveStyle) {
        if (source.Count == 0) return;
        WriteLockConditional(() => {
            var definitions = new LegacyXlsColumnLayout?[256];
            foreach (LegacyXlsColumnLayout column in source) {
                if (column.StartColumn < 1 || column.EndColumn > definitions.Length || column.StartColumn > column.EndColumn) {
                    throw new InvalidDataException("Invalid imported legacy column range.");
                }
                for (int index = column.StartColumn - 1; index < column.EndColumn; index++) definitions[index] = column;
            }
            var columns = new Columns();
            for (int first = 0; first < definitions.Length; first++) {
                LegacyXlsColumnLayout? definition = definitions[first];
                if (definition == null) continue;
                int last = first;
                while (last + 1 < definitions.Length && SameLayout(definition, definitions[last + 1])) last++;
                var column = new Column {
                    Min = (uint)first + 1, Max = (uint)last + 1,
                    Hidden = definition.Hidden, OutlineLevel = definition.OutlineLevel, Collapsed = definition.Collapsed
                };
                if (definition.Width > 0) {
                    column.Width = NormalizeColumnWidth(definition.Width);
                    column.CustomWidth = true;
                }
                uint? style = definition.StyleIndex.HasValue ? resolveStyle(definition.StyleIndex.Value) : null;
                if (style.HasValue) column.Style = style.Value;
                columns.Append(column);
                first = last;
            }
            WorksheetRoot.GetFirstChild<Columns>()?.Remove();
            WorksheetRoot.AddChild(columns, true);
        });

        static bool SameLayout(LegacyXlsColumnLayout left, LegacyXlsColumnLayout? right) =>
            right != null && left.Width == right.Width && left.Hidden == right.Hidden
            && left.OutlineLevel == right.OutlineLevel && left.Collapsed == right.Collapsed && left.StyleIndex == right.StyleIndex;
    }

    /// <summary>Retains the source font zero used by native column-width units.</summary>
    internal void SetImportedDefaultFont(LegacyXlsWorkbook source) {
        if (source.Fonts.Count == 0) return;
        var stylesPart = _excelDocument.WorkbookPartRoot.WorkbookStylesPart!;
        Stylesheet styles = stylesPart.Stylesheet!;
        var candidate = new CellFormat { FontId = 0U };
        var normalized = new Stylesheet();
        EnsureDefaultStylePrimitives(normalized);
        ApplyLegacyFont(normalized, candidate, source, source.Fonts[0]);
        Font font = normalized.Fonts!.Elements<Font>().ElementAt((int)(candidate.FontId?.Value ?? 0));
        styles.Fonts!.ReplaceChild((Font)font.CloneNode(true), styles.Fonts.GetFirstChild<Font>()!);
        SaveStylesheet(stylesPart);
    }

    /// <summary>Imports inert formula text and its original typed cache without evaluating the expression.</summary>
    internal void SetImportedFormula(int row, int column, string formula, object? cachedValue, bool isError) {
        WriteLockConditional(() => {
            CellFormulaCore(row, column, formula);
            Cell cell = GetCell(row, column);
            ExcelWorksheetScan.RemoveInlineString(cell);
            if (cachedValue == null) return;
            FormulaArgumentValue cached = isError ? FormulaArgumentValue.Error((string)cachedValue!)
                : cachedValue is bool boolean ? FormulaArgumentValue.Boolean(boolean)
                : cachedValue is double number ? new FormulaArgumentValue(number, null)
                : new FormulaArgumentValue(null, cachedValue as string ?? string.Empty, sourceCellKind: ExcelCellDataKind.Text);
            SetFormulaCachedValue(cell, cached);
            cell.CellFormula!.CalculateCell = false;
        });
    }
}
