namespace OfficeIMO.Markup;

public sealed partial class OfficeMarkupCSharpEmitter {
    private static void EmitWorkbook(OfficeMarkupDocument document, OfficeMarkupEmitterOptions options, StringBuilder sb) {
        sb.AppendLine("using OfficeIMO.Excel;");
        sb.AppendLine("using OfficeIMO.Excel.Enums;");
        sb.AppendLine("using C = DocumentFormat.OpenXml.Drawing.Charts;");
        sb.AppendLine();
        sb.AppendLine($"using ExcelDocument workbook = ExcelDocument.Create({options.FilePathVariable});");
        sb.AppendLine("ExcelSheet? sheet = null;");
        sb.AppendLine("ExcelSheet GetOrAddSheet(string name) {");
        sb.AppendLine("    return workbook.TryGetSheet(name, out var existingSheet) ? existingSheet! : workbook.AddWorkSheet(name);");
        sb.AppendLine("}");
        var chartIndex = 0;
        foreach (var block in document.Blocks) {
            switch (block) {
                case OfficeMarkupSheetBlock sheetBlock:
                    sb.AppendLine($"sheet = GetOrAddSheet({CsString(sheetBlock.Name)});");
                    break;
                case OfficeMarkupRangeBlock range:
                    var (rangeSheetExpression, rangeAddress) = ResolveWorkbookTarget(range.Sheet, range.Address);
                    sb.AppendLine($"// Range {CsString(range.Address)}");
                    EmitRangeValues(range, sb, rangeSheetExpression, rangeAddress);
                    break;
                case OfficeMarkupFormulaBlock formula:
                    var (formulaSheetExpression, formulaCell) = ResolveWorkbookTarget(formula.Sheet, formula.Cell);
                    if (TryParseCellAddress(formulaCell, out var formulaRow, out var formulaColumn)) {
                        sb.AppendLine($"{formulaSheetExpression}.CellFormula({formulaRow}, {formulaColumn}, {CsString(formula.Expression)});");
                    } else {
                        sb.AppendLine($"// Set formula {CsString(formula.Expression)} in cell {CsString(formula.Cell)}.");
                    }

                    break;
                case OfficeMarkupNamedTableBlock table:
                    var (tableSheetExpression, tableRange) = ResolveWorkbookTarget(GetAttribute(table.Attributes, "sheet"), table.Range);
                    sb.AppendLine($"{tableSheetExpression}.AddTable({CsString(tableRange)}, hasHeader: {BoolLiteral(table.HasHeader)}, name: {CsString(table.Name)}, style: TableStyle.TableStyleMedium2);");
                    break;
                case OfficeMarkupChartBlock chart:
                    EmitExcelChart(chart, sb, ++chartIndex);
                    break;
                case OfficeMarkupFormattingBlock formatting:
                    var (formattingSheetExpression, formattingTarget) = ResolveWorkbookTarget(GetAttribute(formatting.Attributes, "sheet"), formatting.Target);
                    EmitWorkbookFormatting(formatting, formattingSheetExpression, formattingTarget, sb);
                    break;
                default:
                    sb.AppendLine($"// {block.Kind}: {CsString(Describe(block))}");
                    break;
            }
        }

        sb.AppendLine("workbook.Save();");
    }

    private static void EmitRangeValues(OfficeMarkupRangeBlock range, StringBuilder sb, string sheetExpression, string address) {
        int startRow;
        int startColumn;
        if (!TryParseCellAddress(address, out startRow, out startColumn)) {
            startRow = 1;
            startColumn = 1;
            sb.AppendLine($"// Could not parse range start {CsString(range.Address)}. Values are emitted from row 1, column 1.");
        }

        for (int row = 0; row < range.Values.Count; row++) {
            var values = range.Values[row];
            for (int column = 0; column < values.Count; column++) {
                sb.AppendLine($"{sheetExpression}.CellValue({startRow + row}, {startColumn + column}, {CsString(values[column])});");
            }
        }
    }

    private static void EmitWorkbookFormatting(OfficeMarkupFormattingBlock formatting, string sheetExpression, string target, StringBuilder sb) {
        var cells = EnumerateTargetCells(target).ToList();
        if (cells.Count == 0) {
            sb.AppendLine($"// Could not parse formatting target {CsString(target)} style={CsString(formatting.Style ?? string.Empty)} numberFormat={CsString(formatting.NumberFormat ?? string.Empty)}.");
            return;
        }

        var fill = GetAttribute(formatting.Attributes, "fill", "background");
        var fontColor = GetAttribute(formatting.Attributes, "color", "font-color", "fontColor", "text-color", "textColor", "textcolor");
        var bold = GetAttribute(formatting.Attributes, "bold");
        var italic = GetAttribute(formatting.Attributes, "italic");
        var underline = GetAttribute(formatting.Attributes, "underline");
        var alignment = GetAttribute(formatting.Attributes, "align", "alignment", "horizontal-align", "horizontalAlign", "horizontalalignment", "text-align", "textAlign");
        var verticalAlignment = GetAttribute(formatting.Attributes, "vertical-align", "verticalAlign", "verticalalignment", "valign");
        var wrap = GetAttribute(formatting.Attributes, "wrap", "wrap-text", "wrapText");
        var border = GetAttribute(formatting.Attributes, "border", "border-style", "borderStyle");
        var borderColor = GetAttribute(formatting.Attributes, "border-color", "borderColor", "line-color", "lineColor");

        foreach (var (row, column) in cells) {
            if (!string.IsNullOrWhiteSpace(formatting.NumberFormat)) {
                sb.AppendLine($"{sheetExpression}.FormatCell({row}, {column}, {CsString(formatting.NumberFormat!)});");
            }

            if (!string.IsNullOrWhiteSpace(fill)) {
                sb.AppendLine($"{sheetExpression}.CellBackground({row}, {column}, {CsString(fill!)});");
            }

            if (!string.IsNullOrWhiteSpace(fontColor)) {
                sb.AppendLine($"{sheetExpression}.CellFontColor({row}, {column}, {CsString(fontColor!)});");
            }

            if (!string.IsNullOrWhiteSpace(bold) && IsTruthy(bold!)) {
                sb.AppendLine($"{sheetExpression}.CellBold({row}, {column}, true);");
            }

            if (!string.IsNullOrWhiteSpace(italic) && IsTruthy(italic!)) {
                sb.AppendLine($"{sheetExpression}.CellItalic({row}, {column}, true);");
            }

            if (!string.IsNullOrWhiteSpace(underline) && IsTruthy(underline!)) {
                sb.AppendLine($"{sheetExpression}.CellUnderline({row}, {column}, true);");
            }

            if (TryGetHorizontalAlignmentIdentifier(alignment, out var alignmentIdentifier)) {
                sb.AppendLine($"{sheetExpression}.CellAlign({row}, {column}, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.{alignmentIdentifier});");
            }

            if (TryGetVerticalAlignmentIdentifier(verticalAlignment, out var verticalAlignmentIdentifier)) {
                sb.AppendLine($"{sheetExpression}.CellVerticalAlign({row}, {column}, DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.{verticalAlignmentIdentifier});");
            }

            if (!string.IsNullOrWhiteSpace(wrap) && IsTruthy(wrap!)) {
                sb.AppendLine($"{sheetExpression}.WrapCells({row}, {row}, {column});");
            }

            if (TryGetBorderStyleIdentifier(border, out var borderStyleIdentifier)) {
                var borderColorArgument = !string.IsNullOrWhiteSpace(borderColor)
                    ? $", {CsString(borderColor!)}"
                    : string.Empty;
                sb.AppendLine($"{sheetExpression}.CellBorder({row}, {column}, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.{borderStyleIdentifier}{borderColorArgument});");
            }
        }
    }
}
