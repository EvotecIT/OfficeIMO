namespace OfficeIMO.Markup;

/// <summary>
/// Emits starter C# code from the semantic OfficeIMO markup AST.
/// </summary>
public sealed class OfficeMarkupCSharpEmitter {
    public string Emit(OfficeMarkupDocument document, OfficeMarkupEmitterOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options ??= new OfficeMarkupEmitterOptions();
        var sb = new StringBuilder();
        if (options.IncludeHeader) {
            sb.AppendLine("// Generated from OfficeIMO.Markup semantic AST.");
            sb.AppendLine("// Extend the emitted code when the authored markup reaches beyond Markdown.");
        }

        switch (document.Profile) {
            case OfficeMarkupProfile.Presentation:
                EmitPresentation(document, options, sb);
                break;
            case OfficeMarkupProfile.Workbook:
                EmitWorkbook(document, options, sb);
                break;
            case OfficeMarkupProfile.Document:
            case OfficeMarkupProfile.Common:
            default:
                EmitWordDocument(document, options, sb);
                break;
        }

        return sb.ToString().TrimEnd();
    }

    private static void EmitPresentation(OfficeMarkupDocument document, OfficeMarkupEmitterOptions options, StringBuilder sb) {
        sb.AppendLine("using OfficeIMO.PowerPoint;");
        sb.AppendLine("using C = DocumentFormat.OpenXml.Drawing.Charts;");
        sb.AppendLine();
        sb.AppendLine($"using PowerPointPresentation presentation = PowerPointPresentation.Create({options.FilePathVariable});");
        var slideIndex = 0;
        var chartIndex = 0;
        string? activeSection = null;
        foreach (var block in document.Blocks) {
            if (block is OfficeMarkupSlideBlock slide) {
                slideIndex++;
                sb.AppendLine($"PowerPointSlide slide{slideIndex} = presentation.AddSlide();");
                if (!string.IsNullOrWhiteSpace(slide.Section)) {
                    var section = slide.Section!.Trim();
                    sb.AppendLine($"// slide{slideIndex}: section {CsString(section)}");
                    if (!string.Equals(activeSection, section, StringComparison.Ordinal)) {
                        sb.AppendLine($"presentation.AddSection({CsString(section)}, startSlideIndex: {slideIndex - 1});");
                        activeSection = section;
                    }
                }

                if (!string.IsNullOrWhiteSpace(slide.Transition)) {
                    var resolvedTransition = OfficeMarkupTransitionResolver.Parse(slide.Transition);
                    if (!string.IsNullOrWhiteSpace(resolvedTransition.ResolvedIdentifier)) {
                        sb.AppendLine($"slide{slideIndex}.Transition = SlideTransition.{resolvedTransition.ResolvedIdentifier};");
                        EmitTransitionAssignments(sb, $"slide{slideIndex}", resolvedTransition);
                    }

                    if (resolvedTransition.HasArguments || string.IsNullOrWhiteSpace(resolvedTransition.ResolvedIdentifier)) {
                        EmitTransitionComments(sb, resolvedTransition);
                    }
                }

                if (!string.IsNullOrWhiteSpace(slide.Background)) {
                    sb.AppendLine($"// slide{slideIndex}: background {CsString(slide.Background!)}");
                }

                if (!string.IsNullOrWhiteSpace(slide.Layout)) {
                    sb.AppendLine($"// slide{slideIndex}: layout {CsString(slide.Layout!)}");
                }

                if (!string.IsNullOrWhiteSpace(slide.Title)) {
                    sb.AppendLine($"slide{slideIndex}.AddTextBox({CsString(slide.Title!)});");
                }

                EmitContentBlocksForSlide(slide.Blocks, $"slide{slideIndex}", sb, ref chartIndex);
                if (!string.IsNullOrWhiteSpace(slide.Notes)) {
                    sb.AppendLine($"slide{slideIndex}.Notes.Text = {CsString(slide.Notes!)};");
                }

                sb.AppendLine();
            } else {
                sb.AppendLine($"// Presentation-level {block.Kind}: {CsString(Describe(block))}");
            }
        }

        sb.AppendLine("presentation.Save();");
    }

    private static void EmitContentBlocksForSlide(IEnumerable<OfficeMarkupBlock> blocks, string slideVariable, StringBuilder sb, ref int chartIndex) {
        foreach (var block in blocks) {
            switch (block) {
                case OfficeMarkupHeadingBlock heading:
                    sb.AppendLine($"{slideVariable}.AddTextBox({CsString(heading.Text)});");
                    break;
                case OfficeMarkupParagraphBlock paragraph:
                    sb.AppendLine($"{slideVariable}.AddTextBox({CsString(paragraph.Text)});");
                    break;
                case OfficeMarkupListBlock list:
                    sb.AppendLine($"{slideVariable}.AddTextBox({CsString(FormatList(list))});");
                    break;
                case OfficeMarkupImageBlock image:
                    EmitPlacementComment(image.Placement, sb);
                    sb.AppendLine($"{slideVariable}.AddPicture({CsString(image.Source)});");
                    break;
                case OfficeMarkupTableBlock table:
                    var rows = Math.Max(1, table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0));
                    var columns = Math.Max(1, table.Headers.Count > 0 ? table.Headers.Count : table.Rows.Select(r => r.Count).DefaultIfEmpty(1).Max());
                    sb.AppendLine($"var table = {slideVariable}.AddTable({rows}, {columns});");
                    sb.AppendLine("// Fill table cells from the AST before saving.");
                    break;
                case OfficeMarkupDiagramBlock diagram:
                    EmitPlacementComment(diagram.Placement, sb);
                    sb.AppendLine($"// Render {diagram.Language} diagram to an image, then call {slideVariable}.AddPicture(...).");
                    break;
                case OfficeMarkupChartBlock chart:
                    EmitPowerPointChart(chart, slideVariable, sb, ++chartIndex);
                    break;
                case OfficeMarkupTextBoxBlock textBox:
                    EmitPlacementComment(textBox.Placement, sb);
                    if (!string.IsNullOrWhiteSpace(textBox.Style)) {
                        sb.AppendLine($"// Textbox style: {CsString(textBox.Style!)}");
                    }

                    sb.AppendLine($"{slideVariable}.AddTextBox({CsString(textBox.Text)});");
                    break;
                case OfficeMarkupCardBlock card:
                    EmitPlacementComment(card.Placement, sb);
                    if (!string.IsNullOrWhiteSpace(card.Style)) {
                        sb.AppendLine($"// Card style: {CsString(card.Style!)}");
                    }

                    sb.AppendLine($"{slideVariable}.AddTextBox({CsString((card.Title ?? string.Empty) + Environment.NewLine + card.Body)});");
                    break;
                case OfficeMarkupColumnsBlock columnsBlock:
                    EmitPlacementComment(columnsBlock.Placement, sb);
                    sb.AppendLine($"// Start a semantic columns region; gap={CsString(columnsBlock.Gap ?? string.Empty)}. Place following Column blocks into separate slide regions.");
                    break;
                case OfficeMarkupColumnBlock column:
                    EmitComment(sb, $"Column {column.ColumnKind} width={column.Width ?? string.Empty}");
                    if (!string.IsNullOrWhiteSpace(column.Body)) {
                        EmitComment(sb, column.Body);
                    }

                    break;
                default:
                    sb.AppendLine($"// {block.Kind}: {CsString(Describe(block))}");
                    break;
            }
        }
    }

    private static void EmitTransitionComments(StringBuilder sb, OfficeMarkupResolvedTransition resolvedTransition) {
        sb.AppendLine($"// Transition details: {CsString(resolvedTransition.RawText ?? string.Empty)}");

        if (!string.IsNullOrWhiteSpace(resolvedTransition.Effect)) {
            sb.AppendLine($"// Transition effect: {CsString(resolvedTransition.Effect!)}");
        }

        if (!string.IsNullOrWhiteSpace(resolvedTransition.ResolvedIdentifier)) {
            sb.AppendLine($"// Transition native enum: {CsString(resolvedTransition.ResolvedIdentifier!)}");
        }

        var direction = GetTransitionAttribute(resolvedTransition, "direction", "dir", "orientation", "axis", "mode");
        if (!string.IsNullOrWhiteSpace(direction)) {
            sb.AppendLine($"// Transition direction: {CsString(direction!)}");
        }

        var duration = GetTransitionAttribute(resolvedTransition, "duration");
        if (!string.IsNullOrWhiteSpace(duration)) {
            sb.AppendLine($"// Transition duration: {CsString(duration!)}");
        }

        var speed = GetTransitionAttribute(resolvedTransition, "speed", "spd");
        if (!string.IsNullOrWhiteSpace(speed)) {
            sb.AppendLine($"// Transition speed: {CsString(speed!)}");
        }

        var advanceOnClick = GetTransitionAttribute(resolvedTransition, "advance-on-click", "advanceonclick", "advance-click", "onclick", "click");
        if (!string.IsNullOrWhiteSpace(advanceOnClick)) {
            sb.AppendLine($"// Transition advance-on-click: {CsString(advanceOnClick!)}");
        }

        var advanceAfter = GetTransitionAttribute(resolvedTransition, "advance-after", "advanceafter", "after", "delay");
        if (!string.IsNullOrWhiteSpace(advanceAfter)) {
            sb.AppendLine($"// Transition advance-after: {CsString(advanceAfter!)}");
        }
    }

    private static void EmitTransitionAssignments(StringBuilder sb, string slideVariable, OfficeMarkupResolvedTransition resolvedTransition) {
        if (TryGetTransitionSpeed(resolvedTransition, out var speedIdentifier)) {
            sb.AppendLine($"{slideVariable}.TransitionSpeed = SlideTransitionSpeed.{speedIdentifier};");
        }

        if (TryGetTransitionSeconds(resolvedTransition, out var durationSeconds, "duration", "dur")) {
            sb.AppendLine($"{slideVariable}.TransitionDurationSeconds = {FormatDoubleLiteral(durationSeconds)};");
        }

        if (TryGetTransitionBoolean(resolvedTransition, out var advanceOnClick, "advance-on-click", "advanceonclick", "advance-click", "onclick", "click")) {
            sb.AppendLine($"{slideVariable}.TransitionAdvanceOnClick = {(advanceOnClick ? "true" : "false")};");
        }

        if (TryGetTransitionSeconds(resolvedTransition, out var advanceAfterSeconds, "advance-after", "advanceafter", "after", "delay")) {
            sb.AppendLine($"{slideVariable}.TransitionAdvanceAfterSeconds = {FormatDoubleLiteral(advanceAfterSeconds)};");
        }
    }

    private static string? GetTransitionAttribute(OfficeMarkupResolvedTransition resolvedTransition, params string[] names) {
        foreach (var name in names) {
            if (resolvedTransition.Attributes.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)) {
                return value.Trim();
            }
        }

        return null;
    }

    private static bool TryGetTransitionSpeed(OfficeMarkupResolvedTransition resolvedTransition, out string identifier) {
        identifier = string.Empty;
        var value = GetTransitionAttribute(resolvedTransition, "speed", "spd");
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        switch (NormalizeTransitionToken(value)) {
            case "slow":
                identifier = "Slow";
                return true;
            case "medium":
            case "med":
                identifier = "Medium";
                return true;
            case "fast":
                identifier = "Fast";
                return true;
            default:
                return false;
        }
    }

    private static bool TryGetTransitionSeconds(OfficeMarkupResolvedTransition resolvedTransition, out double seconds, params string[] names) {
        seconds = default;
        var value = GetTransitionAttribute(resolvedTransition, names);
        return !string.IsNullOrWhiteSpace(value)
               && double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out seconds);
    }

    private static bool TryGetTransitionBoolean(OfficeMarkupResolvedTransition resolvedTransition, out bool enabled, params string[] names) {
        enabled = default;
        var value = GetTransitionAttribute(resolvedTransition, names);
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        switch (NormalizeTransitionToken(value)) {
            case "true":
            case "yes":
            case "on":
            case "1":
                enabled = true;
                return true;
            case "false":
            case "no":
            case "off":
            case "0":
                enabled = false;
                return true;
            default:
                return false;
        }
    }

    private static string NormalizeTransitionToken(string? value) =>
        new string((value ?? string.Empty).Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());

    private static string FormatDoubleLiteral(double value) =>
        value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);

    private static void EmitWordDocument(OfficeMarkupDocument document, OfficeMarkupEmitterOptions options, StringBuilder sb) {
        sb.AppendLine("using OfficeIMO.Word;");
        sb.AppendLine();
        sb.AppendLine($"using WordDocument document = WordDocument.Create({options.FilePathVariable});");
        foreach (var block in document.Blocks) {
            EmitWordBlock(block, "document", sb);
        }

        sb.AppendLine("document.Save();");
    }

    private static void EmitWordBlock(OfficeMarkupBlock block, string documentVariable, StringBuilder sb) {
        switch (block) {
            case OfficeMarkupHeadingBlock heading:
                sb.AppendLine($"{documentVariable}.AddParagraph({CsString(heading.Text)}).SetStyle(WordParagraphStyles.Heading{Math.Max(1, Math.Min(6, heading.Level))});");
                break;
            case OfficeMarkupParagraphBlock paragraph:
                sb.AppendLine($"{documentVariable}.AddParagraph({CsString(paragraph.Text)});");
                break;
            case OfficeMarkupListBlock list:
                foreach (var item in list.Items) {
                    sb.AppendLine($"{documentVariable}.AddParagraph({CsString(item.Text)});");
                }

                break;
            case OfficeMarkupPageBreakBlock:
                sb.AppendLine($"{documentVariable}.AddPageBreak();");
                break;
            case OfficeMarkupHeaderFooterBlock headerFooter when string.Equals(headerFooter.HeaderFooterKind, "header", StringComparison.OrdinalIgnoreCase):
                sb.AppendLine($"{documentVariable}.HeaderDefaultOrCreate.AddParagraph({CsString(headerFooter.Text)});");
                break;
            case OfficeMarkupHeaderFooterBlock headerFooter:
                sb.AppendLine($"{documentVariable}.FooterDefaultOrCreate.AddParagraph({CsString(headerFooter.Text)});");
                break;
            case OfficeMarkupTableOfContentsBlock toc:
                if (!string.IsNullOrWhiteSpace(toc.Title)) {
                    sb.AppendLine($"{documentVariable}.AddParagraph({CsString(toc.Title!)});");
                }

                sb.AppendLine($"{documentVariable}.AddTableOfContent(TableOfContentStyle.Template1, {toc.MinLevel ?? 1}, {toc.MaxLevel ?? 3});");
                break;
            case OfficeMarkupSectionBlock section:
                sb.AppendLine($"// Section: {CsString(section.Name ?? "section")}");
                foreach (var child in section.Blocks) {
                    EmitWordBlock(child, documentVariable, sb);
                }
                break;
            case OfficeMarkupImageBlock image:
                sb.AppendLine($"{documentVariable}.AddParagraph().AddImage({CsString(image.Source)});");
                break;
            case OfficeMarkupDiagramBlock diagram:
                sb.AppendLine($"// Render {diagram.Language} diagram to an image, then add it with document.AddParagraph().AddImage(...).");
                break;
            case OfficeMarkupChartBlock chart:
                EmitWordChart(chart, documentVariable, sb);
                break;
            case OfficeMarkupTableBlock table:
                sb.AppendLine($"var wordTable = {documentVariable}.AddTable({Math.Max(1, table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0))}, {Math.Max(1, table.Headers.Count > 0 ? table.Headers.Count : table.Rows.Select(row => row.Count).DefaultIfEmpty(1).Max())});");
                sb.AppendLine("// Fill wordTable cells from the semantic table AST.");
                break;
            default:
                sb.AppendLine($"// {block.Kind}: {CsString(Describe(block))}");
                break;
        }
    }

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

    private static void EmitPowerPointChart(OfficeMarkupChartBlock chart, string slideVariable, StringBuilder sb, int chartIndex) {
        EmitPlacementComment(chart.Placement, sb);
        if (chart.Data.Count < 2 || chart.Data[0].Count < 2) {
            sb.AppendLine($"// Add {chart.ChartType} chart from source {CsString(chart.Source ?? string.Empty)}.");
            return;
        }

        var dataVariable = $"chartData{chartIndex}";
        EmitPowerPointChartData(chart, dataVariable, sb);
        var method = ToPowerPointChartMethod(chart.ChartType);
        var chartVariable = $"chart{chartIndex}";
        sb.AppendLine($"var {chartVariable} = {slideVariable}.{method}({dataVariable});");
        if (!string.IsNullOrWhiteSpace(chart.Title)) {
            sb.AppendLine($"{chartVariable}.SetTitle({CsString(chart.Title!)});");
        }

        EmitPowerPointChartSemanticOptions(chart, chartVariable, sb);
    }

    private static void EmitPowerPointChartSemanticOptions(OfficeMarkupChartBlock chart, string chartVariable, StringBuilder sb) {
        if (GetAttribute(chart.Attributes, "category-title", "categoryTitle", "x-title", "xTitle", "x-axis-title", "xAxisTitle") is { Length: > 0 } categoryTitle) {
            sb.AppendLine($"{chartVariable}.SetCategoryAxisTitle({CsString(categoryTitle)});");
        }

        if (GetAttribute(chart.Attributes, "value-title", "valueTitle", "y-title", "yTitle", "y-axis-title", "yAxisTitle") is { Length: > 0 } valueTitle) {
            sb.AppendLine($"{chartVariable}.SetValueAxisTitle({CsString(valueTitle)});");
        }

        if (GetAttribute(chart.Attributes, "category-format", "categoryFormat", "x-format", "xFormat", "category-number-format", "categoryNumberFormat") is { Length: > 0 } categoryFormat) {
            sb.AppendLine($"{chartVariable}.SetCategoryAxisNumberFormat({CsString(categoryFormat)});");
        }

        if (GetAttribute(chart.Attributes, "value-format", "valueFormat", "y-format", "yFormat", "value-number-format", "valueNumberFormat") is { Length: > 0 } valueFormat) {
            sb.AppendLine($"{chartVariable}.SetValueAxisNumberFormat({CsString(valueFormat)});");
        }

        var legend = GetAttribute(chart.Attributes, "legend", "legend-position", "legendPosition");
        if (!string.IsNullOrWhiteSpace(legend)) {
            var legendValue = legend!;
            var normalized = NormalizeToken(legendValue);
            if (normalized is "false" or "none" or "hidden" or "off") {
                sb.AppendLine($"{chartVariable}.HideLegend();");
            } else if (TryGetLegendPositionIdentifier(legendValue, out var legendPosition)) {
                sb.AppendLine($"{chartVariable}.SetLegend(C.LegendPositionValues.{legendPosition});");
            }
        }

        var labels = GetAttribute(chart.Attributes, "labels", "data-labels", "dataLabels");
        if (!string.IsNullOrWhiteSpace(labels)) {
            if (IsTruthy(labels!)) {
                sb.AppendLine($"{chartVariable}.SetDataLabels(showValue: true, showCategoryName: false, showSeriesName: false, showLegendKey: false, showPercent: false);");
                var labelPosition = GetAttribute(chart.Attributes, "label-position", "labelPosition", "data-label-position", "dataLabelPosition");
                if (TryGetDataLabelPositionIdentifier(labelPosition, out var dataLabelPosition)) {
                    sb.AppendLine($"{chartVariable}.SetDataLabelPosition(C.DataLabelPositionValues.{dataLabelPosition});");
                }

                var labelFormat = GetAttribute(chart.Attributes, "label-format", "labelFormat", "data-label-format", "dataLabelFormat");
                if (!string.IsNullOrWhiteSpace(labelFormat)) {
                    sb.AppendLine($"{chartVariable}.SetDataLabelNumberFormat({CsString(labelFormat!)});");
                }
            } else {
                sb.AppendLine($"{chartVariable}.ClearDataLabels();");
            }
        }

        var gridlines = GetAttribute(chart.Attributes, "gridlines");
        var valueGridlines = GetAttribute(chart.Attributes, "value-gridlines", "valueGridlines", "y-gridlines", "yGridlines") ?? gridlines;
        var categoryGridlines = GetAttribute(chart.Attributes, "category-gridlines", "categoryGridlines", "x-gridlines", "xGridlines");
        if (!string.IsNullOrWhiteSpace(valueGridlines)) {
            sb.AppendLine(IsTruthy(valueGridlines!)
                ? $"{chartVariable}.SetValueAxisGridlines(showMajor: true, showMinor: false);"
                : $"{chartVariable}.ClearValueAxisGridlines();");
        }

        if (!string.IsNullOrWhiteSpace(categoryGridlines)) {
            sb.AppendLine(IsTruthy(categoryGridlines!)
                ? $"{chartVariable}.SetCategoryAxisGridlines(showMajor: true, showMinor: false);"
                : $"{chartVariable}.ClearCategoryAxisGridlines();");
        }
    }

    private static void EmitPowerPointChartData(OfficeMarkupChartBlock chart, string variableName, StringBuilder sb) {
        var headers = chart.Data[0];
        var categories = chart.Data.Skip(1).Select(row => row.Count > 0 ? row[0] : string.Empty).ToList();
        sb.AppendLine($"var {variableName} = new PowerPointChartData(");
        sb.AppendLine($"    new[] {{ {string.Join(", ", categories.Select(CsString))} }},");
        sb.AppendLine("    new[] {");
        for (int seriesIndex = 1; seriesIndex < headers.Count; seriesIndex++) {
            var values = chart.Data.Skip(1).Select(row => NumericLiteral(row.Count > seriesIndex ? row[seriesIndex] : "0"));
            var comma = seriesIndex == headers.Count - 1 ? string.Empty : ",";
            sb.AppendLine($"        new PowerPointChartSeries({CsString(headers[seriesIndex])}, new[] {{ {string.Join(", ", values)} }}){comma}");
        }

        sb.AppendLine("    });");
    }

    private static void EmitWordChart(OfficeMarkupChartBlock chart, string documentVariable, StringBuilder sb) {
        if (chart.Data.Count < 2 || chart.Data[0].Count < 2) {
            sb.AppendLine($"// Add {chart.ChartType} chart from source {CsString(chart.Source ?? string.Empty)}.");
            return;
        }

        sb.AppendLine($"// Add {chart.ChartType} chart {CsString(chart.Title ?? string.Empty)} from inline data.");
        sb.AppendLine($"// Word chart APIs can consume the same categories and series represented by this AST node.");
    }

    private static void EmitExcelChart(OfficeMarkupChartBlock chart, StringBuilder sb, int chartIndex) {
        var row = 1;
        var column = 6;
        chart.Attributes.TryGetValue("cell", out var cell);
        var (chartSheetExpression, placementCell) = ResolveWorkbookTarget(chart.Sheet, cell ?? string.Empty);
        if (TryParseCellAddress(placementCell, out var parsedRow, out var parsedColumn)) {
            row = parsedRow;
            column = parsedColumn;
        }

        var width = TryParseInt(chart.Placement?.Width, out var parsedWidth) ? parsedWidth : 640;
        var height = TryParseInt(chart.Placement?.Height, out var parsedHeight) ? parsedHeight : 360;
        var chartType = ToExcelChartType(chart.ChartType);
        var chartVariable = $"chart{chartIndex}";
        if (!string.IsNullOrWhiteSpace(chart.Source)) {
            EmitExcelChartFromSource(chart, sb, chartIndex, chartVariable, chartSheetExpression, row, column, width, height, chartType);
            EmitExcelChartSemanticOptions(chart, chartVariable, sb);
            return;
        }

        if (chart.Data.Count < 2 || chart.Data[0].Count < 2) {
            sb.AppendLine($"// Add {chart.ChartType} chart from source {CsString(chart.Source ?? string.Empty)}.");
            return;
        }

        var dataVariable = $"chartData{chartIndex}";
        EmitExcelChartData(chart, dataVariable, sb);
        sb.AppendLine($"var {chartVariable} = {chartSheetExpression}.AddChart({dataVariable}, row: {row}, column: {column}, widthPixels: {width}, heightPixels: {height}, type: ExcelChartType.{chartType}, title: {CsString(chart.Title ?? string.Empty)});");
        EmitExcelChartSemanticOptions(chart, chartVariable, sb);
    }

    private static void EmitExcelChartFromSource(
        OfficeMarkupChartBlock chart,
        StringBuilder sb,
        int chartIndex,
        string chartVariable,
        string chartSheetExpression,
        int row,
        int column,
        int width,
        int height,
        string chartType) {
        var source = chart.Source ?? string.Empty;
        if (TrySplitSheetQualifiedReference(source, out var sourceSheetName, out var localSource)) {
            var sourceSheetVariable = $"chartSourceSheet{chartIndex}";
            sb.AppendLine($"var {sourceSheetVariable} = GetOrAddSheet({CsString(sourceSheetName)});");
            EmitExcelChartDataRangeFromSource(sb, chartIndex, sourceSheetVariable, localSource);
            sb.AppendLine($"var {chartVariable} = {chartSheetExpression}.AddChart(chartDataRange{chartIndex}, row: {row}, column: {column}, widthPixels: {width}, heightPixels: {height}, type: ExcelChartType.{chartType}, title: {CsString(chart.Title ?? string.Empty)});");
            return;
        }

        if (source.IndexOf(':') >= 0) {
            sb.AppendLine($"var {chartVariable} = {chartSheetExpression}.AddChartFromRange({CsString(source)}, row: {row}, column: {column}, widthPixels: {width}, heightPixels: {height}, type: ExcelChartType.{chartType}, title: {CsString(chart.Title ?? string.Empty)});");
        } else {
            sb.AppendLine($"var {chartVariable} = {chartSheetExpression}.AddChartFromTable({CsString(source)}, row: {row}, column: {column}, widthPixels: {width}, heightPixels: {height}, type: ExcelChartType.{chartType}, title: {CsString(chart.Title ?? string.Empty)});");
        }
    }

    private static void EmitExcelChartDataRangeFromSource(StringBuilder sb, int chartIndex, string sourceSheetVariable, string localSource) {
        if (localSource.IndexOf(':') >= 0) {
            sb.AppendLine($"var (chartR1_{chartIndex}, chartC1_{chartIndex}, chartR2_{chartIndex}, chartC2_{chartIndex}) = A1.ParseRange({CsString(localSource)});");
        } else {
            sb.AppendLine($"var chartSourceRange{chartIndex} = {sourceSheetVariable}.GetTableRange({CsString(localSource)}) ?? throw new global::System.InvalidOperationException({CsString("Chart source table was not found.")});");
            sb.AppendLine($"var (chartR1_{chartIndex}, chartC1_{chartIndex}, chartR2_{chartIndex}, chartC2_{chartIndex}) = A1.ParseRange(chartSourceRange{chartIndex});");
        }

        sb.AppendLine($"var chartDataRange{chartIndex} = new ExcelChartDataRange({sourceSheetVariable}.Name, chartR1_{chartIndex}, chartC1_{chartIndex}, chartR2_{chartIndex} - chartR1_{chartIndex}, chartC2_{chartIndex} - chartC1_{chartIndex}, hasHeaderRow: true);");
    }

    private static void EmitExcelChartSemanticOptions(OfficeMarkupChartBlock chart, string chartVariable, StringBuilder sb) {
        if (GetAttribute(chart.Attributes, "category-title", "categoryTitle", "x-title", "xTitle", "x-axis-title", "xAxisTitle") is { Length: > 0 } categoryTitle) {
            sb.AppendLine($"{chartVariable}.SetCategoryAxisTitle({CsString(categoryTitle)});");
        }

        if (GetAttribute(chart.Attributes, "value-title", "valueTitle", "y-title", "yTitle", "y-axis-title", "yAxisTitle") is { Length: > 0 } valueTitle) {
            sb.AppendLine($"{chartVariable}.SetValueAxisTitle({CsString(valueTitle)});");
        }

        if (GetAttribute(chart.Attributes, "category-format", "categoryFormat", "x-format", "xFormat", "category-number-format", "categoryNumberFormat") is { Length: > 0 } categoryFormat) {
            sb.AppendLine($"{chartVariable}.SetCategoryAxisNumberFormat({CsString(categoryFormat)});");
        }

        if (GetAttribute(chart.Attributes, "value-format", "valueFormat", "y-format", "yFormat", "value-number-format", "valueNumberFormat") is { Length: > 0 } valueFormat) {
            sb.AppendLine($"{chartVariable}.SetValueAxisNumberFormat({CsString(valueFormat)});");
        }

        var legend = GetAttribute(chart.Attributes, "legend", "legend-position", "legendPosition");
        if (!string.IsNullOrWhiteSpace(legend)) {
            var legendValue = legend!;
            var normalized = NormalizeToken(legendValue);
            if (normalized is "false" or "none" or "hidden" or "off") {
                sb.AppendLine($"{chartVariable}.HideLegend();");
            } else if (TryGetLegendPositionIdentifier(legendValue, out var legendPosition)) {
                sb.AppendLine($"{chartVariable}.SetLegend(C.LegendPositionValues.{legendPosition});");
            }
        }

        var labels = GetAttribute(chart.Attributes, "labels", "data-labels", "dataLabels");
        if (!string.IsNullOrWhiteSpace(labels) && IsTruthy(labels!)) {
            var labelPosition = GetAttribute(chart.Attributes, "label-position", "labelPosition", "data-label-position", "dataLabelPosition");
            var labelFormat = GetAttribute(chart.Attributes, "label-format", "labelFormat", "data-label-format", "dataLabelFormat");
            var positionExpression = TryGetDataLabelPositionIdentifier(labelPosition, out var dataLabelPosition)
                ? $"C.DataLabelPositionValues.{dataLabelPosition}"
                : "null";
            var numberFormatExpression = !string.IsNullOrWhiteSpace(labelFormat) ? CsString(labelFormat!) : "null";
            sb.AppendLine($"{chartVariable}.SetDataLabels(showValue: true, showCategoryName: false, showSeriesName: false, showLegendKey: false, showPercent: false, position: {positionExpression}, numberFormat: {numberFormatExpression});");
        }

        var gridlines = GetAttribute(chart.Attributes, "gridlines");
        var valueGridlines = GetAttribute(chart.Attributes, "value-gridlines", "valueGridlines", "y-gridlines", "yGridlines") ?? gridlines;
        var categoryGridlines = GetAttribute(chart.Attributes, "category-gridlines", "categoryGridlines", "x-gridlines", "xGridlines");
        if (!string.IsNullOrWhiteSpace(valueGridlines)) {
            sb.AppendLine($"{chartVariable}.SetValueAxisGridlines(showMajor: {BoolLiteral(IsTruthy(valueGridlines!))}, showMinor: false);");
        }

        if (!string.IsNullOrWhiteSpace(categoryGridlines)) {
            sb.AppendLine($"{chartVariable}.SetCategoryAxisGridlines(showMajor: {BoolLiteral(IsTruthy(categoryGridlines!))}, showMinor: false);");
        }
    }

    private static void EmitExcelChartData(OfficeMarkupChartBlock chart, string variableName, StringBuilder sb) {
        var headers = chart.Data[0];
        var categories = chart.Data.Skip(1).Select(row => row.Count > 0 ? row[0] : string.Empty).ToList();
        sb.AppendLine($"var {variableName} = new ExcelChartData(");
        sb.AppendLine($"    new[] {{ {string.Join(", ", categories.Select(CsString))} }},");
        sb.AppendLine("    new[] {");
        for (int seriesIndex = 1; seriesIndex < headers.Count; seriesIndex++) {
            var values = chart.Data.Skip(1).Select(row => NumericLiteral(row.Count > seriesIndex ? row[seriesIndex] : "0"));
            var comma = seriesIndex == headers.Count - 1 ? string.Empty : ",";
            sb.AppendLine($"        new ExcelChartSeries({CsString(headers[seriesIndex])}, new[] {{ {string.Join(", ", values)} }}){comma}");
        }

        sb.AppendLine("    });");
    }

    private static string Describe(OfficeMarkupBlock block) {
        switch (block) {
            case OfficeMarkupHeadingBlock heading:
                return heading.Text;
            case OfficeMarkupParagraphBlock paragraph:
                return paragraph.Text;
            case OfficeMarkupImageBlock image:
                return image.Source;
            case OfficeMarkupCodeBlock code:
                return code.Language;
            case OfficeMarkupChartBlock chart:
                return chart.Title ?? chart.Source ?? chart.ChartType;
            case OfficeMarkupTextBoxBlock textBox:
                return textBox.Text;
            case OfficeMarkupColumnBlock column:
                return column.ColumnKind;
            case OfficeMarkupCardBlock card:
                return card.Title ?? card.Body;
            case OfficeMarkupExtensionBlock extension:
                return extension.Command;
            default:
                return block.Kind.ToString();
        }
    }

    private static string ToPascalIdentifier(string value) {
        var parts = (value ?? string.Empty).Split(new[] { '-', '_', ' ', '=' }, StringSplitOptions.RemoveEmptyEntries);
        var sb = new StringBuilder();
        foreach (var part in parts) {
            if (part.Length == 0) {
                continue;
            }

            sb.Append(char.ToUpperInvariant(part[0]));
            if (part.Length > 1) {
                sb.Append(part.Substring(1));
            }
        }

        return sb.Length == 0 ? "None" : sb.ToString();
    }

    private static string ToPowerPointChartMethod(string chartType) {
        var normalized = NormalizeToken(chartType);
        return normalized switch {
            "line" => "AddLineChart",
            "pie" => "AddPieChart",
            "doughnut" or "donut" => "AddDoughnutChart",
            _ => "AddChart"
        };
    }

    private static string ToExcelChartType(string chartType) {
        var normalized = NormalizeToken(chartType);
        return normalized switch {
            "line" => "Line",
            "bar" => "BarClustered",
            "stackedbar" => "BarStacked",
            "stackedcolumn" => "ColumnStacked",
            "pie" => "Pie",
            "doughnut" or "donut" => "Doughnut",
            "scatter" => "Scatter",
            "area" => "Area",
            _ => "ColumnClustered"
        };
    }

    private static string NormalizeToken(string value) =>
        new string((value ?? string.Empty).Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());

    private static void EmitPlacementComment(OfficeMarkupPlacement? placement, StringBuilder sb) {
        if (placement == null || !placement.HasValue) {
            return;
        }

        sb.AppendLine($"// Placement: x={CsString(placement.X ?? string.Empty)}, y={CsString(placement.Y ?? string.Empty)}, w={CsString(placement.Width ?? string.Empty)}, h={CsString(placement.Height ?? string.Empty)}");
    }

    private static void EmitComment(StringBuilder sb, string text) {
        foreach (var line in (text ?? string.Empty).Replace("\r\n", "\n").Split('\n')) {
            sb.AppendLine($"// {line}");
        }
    }

    private static (string SheetExpression, string LocalReference) ResolveWorkbookTarget(string? explicitSheet, string? reference) {
        if (TrySplitSheetQualifiedReference(reference, out var sheetName, out var localReference)) {
            return ($"GetOrAddSheet({CsString(sheetName)})", localReference);
        }

        if (!string.IsNullOrWhiteSpace(explicitSheet)) {
            return ($"GetOrAddSheet({CsString(explicitSheet!.Trim())})", (reference ?? string.Empty).Trim());
        }

        return ($"GetOrAddSheet({CsString("Sheet1")})", (reference ?? string.Empty).Trim());
    }

    private static bool TrySplitSheetQualifiedReference(string? reference, out string sheetName, out string localReference) {
        sheetName = string.Empty;
        localReference = string.Empty;
        if (string.IsNullOrWhiteSpace(reference)) {
            return false;
        }

        var value = reference!.Trim();
        var bangIndex = value.LastIndexOf('!');
        if (bangIndex <= 0 || bangIndex >= value.Length - 1) {
            return false;
        }

        sheetName = value.Substring(0, bangIndex).Trim().Trim('\'').Replace("''", "'");
        localReference = value.Substring(bangIndex + 1).Trim();
        return !string.IsNullOrWhiteSpace(sheetName) && !string.IsNullOrWhiteSpace(localReference);
    }

    private static bool TryParseCellAddress(string? address, out int row, out int column) {
        row = 1;
        column = 1;
        if (string.IsNullOrWhiteSpace(address)) {
            return false;
        }

        var match = System.Text.RegularExpressions.Regex.Match(address!.Trim(), @"^\$?([A-Za-z]+)\$?(\d+)");
        if (!match.Success) {
            return false;
        }

        column = 0;
        foreach (var character in match.Groups[1]!.Value.ToUpperInvariant()) {
            column = (column * 26) + character - 'A' + 1;
        }

        row = int.Parse(match.Groups[2]!.Value, System.Globalization.CultureInfo.InvariantCulture);
        return row > 0 && column > 0;
    }

    private static bool TryParseInt(string? value, out int result) =>
        int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out result);

    private static string NumericLiteral(string? value) {
        if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var number)) {
            return number.ToString("0.###############", System.Globalization.CultureInfo.InvariantCulture) + "d";
        }

        return "0d";
    }

    private static string FormatList(OfficeMarkupListBlock list) {
        var lines = list.Items.Select((item, index) => list.Ordered
            ? $"{list.Start + index}. {item.Text}"
            : $"- {item.Text}");
        return string.Join(Environment.NewLine, lines);
    }

    private static string CsString(string value) {
        return "@\"" + (value ?? string.Empty).Replace("\"", "\"\"") + "\"";
    }

    private static string? GetAttribute(IDictionary<string, string> attributes, params string[] names) {
        foreach (var name in names) {
            if (attributes.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)) {
                return value.Trim();
            }
        }

        return null;
    }

    private static bool IsTruthy(string value) {
        var normalized = NormalizeToken(value);
        return normalized is not ("false" or "no" or "off" or "none" or "hidden" or "0");
    }

    private static bool TryGetLegendPositionIdentifier(string value, out string identifier) {
        identifier = NormalizeToken(value) switch {
            "left" => "Left",
            "right" => "Right",
            "top" => "Top",
            "bottom" => "Bottom",
            "corner" or "topright" => "TopRight",
            _ => string.Empty
        };
        return identifier.Length > 0;
    }

    private static bool TryGetDataLabelPositionIdentifier(string? value, out string identifier) {
        identifier = NormalizeToken(value ?? string.Empty) switch {
            "center" => "Center",
            "insideend" => "InsideEnd",
            "insidebase" => "InsideBase",
            "outsideend" => "OutsideEnd",
            "bestfit" => "BestFit",
            "left" => "Left",
            "right" => "Right",
            "top" => "Top",
            "bottom" => "Bottom",
            _ => string.Empty
        };
        return identifier.Length > 0;
    }

    private static bool TryGetHorizontalAlignmentIdentifier(string? value, out string identifier) {
        identifier = NormalizeToken(value ?? string.Empty) switch {
            "general" => "General",
            "left" => "Left",
            "center" or "centre" => "Center",
            "centercontinuous" or "centeracross" or "centeracrossselection" => "CenterContinuous",
            "right" => "Right",
            "fill" => "Fill",
            "justify" => "Justify",
            "distributed" => "Distributed",
            _ => string.Empty
        };
        return identifier.Length > 0;
    }

    private static bool TryGetVerticalAlignmentIdentifier(string? value, out string identifier) {
        identifier = NormalizeToken(value ?? string.Empty) switch {
            "top" => "Top",
            "middle" or "center" or "centre" => "Center",
            "bottom" => "Bottom",
            "justify" => "Justify",
            "distributed" => "Distributed",
            _ => string.Empty
        };
        return identifier.Length > 0;
    }

    private static bool TryGetBorderStyleIdentifier(string? value, out string identifier) {
        identifier = NormalizeToken(value ?? string.Empty) switch {
            "true" or "yes" or "on" or "1" or "thin" => "Thin",
            "medium" => "Medium",
            "thick" => "Thick",
            "dashed" => "Dashed",
            "dotted" => "Dotted",
            "double" => "Double",
            "hair" => "Hair",
            "dashdot" => "DashDot",
            "dashdotdot" => "DashDotDot",
            "mediumdashed" => "MediumDashed",
            "mediumdashdot" => "MediumDashDot",
            "mediumdashdotdot" => "MediumDashDotDot",
            "slantdashdot" => "SlantDashDot",
            _ => string.Empty
        };
        return identifier.Length > 0;
    }

    private static IEnumerable<(int Row, int Column)> EnumerateTargetCells(string target) {
        if (string.IsNullOrWhiteSpace(target)) {
            yield break;
        }

        var parts = target.Split(new[] { ':' }, 2, StringSplitOptions.None)
            .Select(part => part.Trim())
            .ToArray();
        if (parts.Length == 1) {
            if (TryParseCellAddress(parts[0], out var singleRow, out var singleColumn)) {
                yield return (singleRow, singleColumn);
            }

            yield break;
        }

        if (!TryParseCellAddress(parts[0], out var startRow, out var startColumn) ||
            !TryParseCellAddress(parts[1], out var endRow, out var endColumn)) {
            yield break;
        }

        if (endRow < startRow) {
            (startRow, endRow) = (endRow, startRow);
        }

        if (endColumn < startColumn) {
            (startColumn, endColumn) = (endColumn, startColumn);
        }

        for (var row = startRow; row <= endRow; row++) {
            for (var column = startColumn; column <= endColumn; column++) {
                yield return (row, column);
            }
        }
    }

    private static string BoolLiteral(bool value) => value ? "true" : "false";
}
