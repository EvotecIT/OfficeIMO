namespace OfficeIMO.Markup;

/// <summary>
/// Emits starter PowerShell code from the semantic OfficeIMO markup AST.
/// </summary>
public sealed class OfficeMarkupPowerShellEmitter {
    public string Emit(OfficeMarkupDocument document, OfficeMarkupEmitterOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options ??= new OfficeMarkupEmitterOptions { FilePathVariable = "$FilePath" };
        var sb = new StringBuilder();
        if (options.IncludeHeader) {
            sb.AppendLine("# Generated from OfficeIMO.Markup semantic AST.");
            sb.AppendLine("# Treat this as the handoff point from Markdown-like authoring to scriptable automation.");
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
                EmitDocument(document, options, sb);
                break;
        }

        return sb.ToString().TrimEnd();
    }

    private static void EmitPresentation(OfficeMarkupDocument document, OfficeMarkupEmitterOptions options, StringBuilder sb) {
        sb.AppendLine("$presentation = New-OfficePowerPoint -FilePath " + options.FilePathVariable);
        var index = 0;
        var chartIndex = 0;
        string? activeSection = null;
        foreach (var block in document.Blocks) {
            if (block is OfficeMarkupSlideBlock slide) {
                index++;
                sb.AppendLine($"$slide{index} = Add-OfficePowerPointSlide -Presentation $presentation");
                if (!string.IsNullOrWhiteSpace(slide.Section)) {
                    var section = slide.Section!.Trim();
                    sb.AppendLine($"# Section: {section}");
                    if (!string.Equals(activeSection, section, StringComparison.Ordinal)) {
                        sb.AppendLine($"$null = $presentation.AddSection({PsString(section)}, {index - 1})");
                        activeSection = section;
                    }
                }

                if (!string.IsNullOrWhiteSpace(slide.Layout)) {
                    sb.AppendLine($"# Layout: {slide.Layout}");
                }

                if (!string.IsNullOrWhiteSpace(slide.Transition)) {
                    var resolvedTransition = OfficeMarkupTransitionResolver.Parse(slide.Transition);
                    if (!string.IsNullOrWhiteSpace(resolvedTransition.ResolvedIdentifier)) {
                        sb.AppendLine($"$slide{index}.Transition = [OfficeIMO.PowerPoint.SlideTransition]::{resolvedTransition.ResolvedIdentifier}");
                        EmitTransitionAssignments(sb, $"$slide{index}", resolvedTransition);
                    }

                    if (resolvedTransition.HasArguments || string.IsNullOrWhiteSpace(resolvedTransition.ResolvedIdentifier)) {
                        EmitTransitionComments(sb, resolvedTransition);
                    }
                }

                if (!string.IsNullOrWhiteSpace(slide.Background)) {
                    sb.AppendLine($"# Background: {slide.Background}");
                }

                if (!string.IsNullOrWhiteSpace(slide.Title)) {
                    sb.AppendLine($"Add-OfficePowerPointText -Slide $slide{index} -Text {PsString(slide.Title!)}");
                }

                foreach (var child in slide.Blocks) {
                    EmitSlideChild(child, $"$slide{index}", sb, ref chartIndex);
                }

                if (!string.IsNullOrWhiteSpace(slide.Notes)) {
                    sb.AppendLine($"Set-OfficePowerPointNotes -Slide $slide{index} -Text {PsString(slide.Notes!)}");
                }
            } else {
                sb.AppendLine($"# {block.Kind}: {Describe(block)}");
            }
        }

        sb.AppendLine("Save-OfficePowerPoint -Presentation $presentation");
    }

    private static void EmitSlideChild(OfficeMarkupBlock block, string slideVariable, StringBuilder sb, ref int chartIndex) {
        switch (block) {
            case OfficeMarkupHeadingBlock heading:
                sb.AppendLine($"Add-OfficePowerPointText -Slide {slideVariable} -Text {PsString(heading.Text)}");
                break;
            case OfficeMarkupParagraphBlock paragraph:
                sb.AppendLine($"Add-OfficePowerPointText -Slide {slideVariable} -Text {PsString(paragraph.Text)}");
                break;
            case OfficeMarkupListBlock list:
                sb.AppendLine($"Add-OfficePowerPointText -Slide {slideVariable} -Text {PsString(FormatList(list))}");
                break;
            case OfficeMarkupImageBlock image:
                EmitPlacementComment(image.Placement, sb);
                sb.AppendLine($"Add-OfficePowerPointImage -Slide {slideVariable} -Path {PsString(image.Source)}");
                break;
            case OfficeMarkupDiagramBlock diagram:
                EmitPlacementComment(diagram.Placement, sb);
                sb.AppendLine($"# Render {diagram.Language} to an image and add it to {slideVariable}.");
                break;
            case OfficeMarkupChartBlock chart:
                EmitChartComment(chart, sb);
                if (!string.IsNullOrWhiteSpace(chart.Source)) {
                    sb.AppendLine($"Add-OfficePowerPointChart -Slide {slideVariable} -Type {PsString(chart.ChartType)} -Source {PsString(chart.Source!)} -Title {PsString(chart.Title ?? string.Empty)}");
                } else if (EmitChartData(chart, $"$chartData{++chartIndex}", sb)) {
                    sb.AppendLine($"Add-OfficePowerPointChart -Slide {slideVariable} -Type {PsString(chart.ChartType)} -Title {PsString(chart.Title ?? string.Empty)} -Data $chartData{chartIndex}");
                }

                break;
            case OfficeMarkupTextBoxBlock textBox:
                EmitPlacementComment(textBox.Placement, sb);
                if (!string.IsNullOrWhiteSpace(textBox.Style)) {
                    sb.AppendLine($"# Style: {textBox.Style}");
                }

                sb.AppendLine($"Add-OfficePowerPointText -Slide {slideVariable} -Text {PsString(textBox.Text)}");
                break;
            case OfficeMarkupCardBlock card:
                EmitPlacementComment(card.Placement, sb);
                if (!string.IsNullOrWhiteSpace(card.Style)) {
                    sb.AppendLine($"# Style: {card.Style}");
                }

                sb.AppendLine($"Add-OfficePowerPointText -Slide {slideVariable} -Text {PsString((card.Title ?? string.Empty) + Environment.NewLine + card.Body)}");
                break;
            case OfficeMarkupColumnsBlock columns:
                EmitPlacementComment(columns.Placement, sb);
                sb.AppendLine($"# Start a semantic columns region; gap={columns.Gap}. Place following Column blocks into separate slide regions.");
                break;
            case OfficeMarkupColumnBlock column:
                EmitComment(sb, $"Column {column.ColumnKind} width={column.Width}");
                if (!string.IsNullOrWhiteSpace(column.Body)) {
                    EmitComment(sb, column.Body);
                }

                break;
            default:
                sb.AppendLine($"# {block.Kind}: {Describe(block)}");
                break;
        }
    }

    private static void EmitTransitionComments(StringBuilder sb, OfficeMarkupResolvedTransition resolvedTransition) {
        sb.AppendLine($"# Transition details: {resolvedTransition.RawText ?? string.Empty}");

        if (!string.IsNullOrWhiteSpace(resolvedTransition.Effect)) {
            sb.AppendLine($"# Transition effect: {resolvedTransition.Effect}");
        }

        if (!string.IsNullOrWhiteSpace(resolvedTransition.ResolvedIdentifier)) {
            sb.AppendLine($"# Transition native enum: {resolvedTransition.ResolvedIdentifier}");
        }

        var direction = GetTransitionAttribute(resolvedTransition, "direction", "dir", "orientation", "axis", "mode");
        if (!string.IsNullOrWhiteSpace(direction)) {
            sb.AppendLine($"# Transition direction: {direction}");
        }

        var duration = GetTransitionAttribute(resolvedTransition, "duration");
        if (!string.IsNullOrWhiteSpace(duration)) {
            sb.AppendLine($"# Transition duration: {duration}");
        }

        var speed = GetTransitionAttribute(resolvedTransition, "speed", "spd");
        if (!string.IsNullOrWhiteSpace(speed)) {
            sb.AppendLine($"# Transition speed: {speed}");
        }

        var advanceOnClick = GetTransitionAttribute(resolvedTransition, "advance-on-click", "advanceonclick", "advance-click", "onclick", "click");
        if (!string.IsNullOrWhiteSpace(advanceOnClick)) {
            sb.AppendLine($"# Transition advance-on-click: {advanceOnClick}");
        }

        var advanceAfter = GetTransitionAttribute(resolvedTransition, "advance-after", "advanceafter", "after", "delay");
        if (!string.IsNullOrWhiteSpace(advanceAfter)) {
            sb.AppendLine($"# Transition advance-after: {advanceAfter}");
        }
    }

    private static void EmitTransitionAssignments(StringBuilder sb, string slideVariable, OfficeMarkupResolvedTransition resolvedTransition) {
        if (TryGetTransitionSpeed(resolvedTransition, out var speedIdentifier)) {
            sb.AppendLine($"{slideVariable}.TransitionSpeed = [OfficeIMO.PowerPoint.SlideTransitionSpeed]::{speedIdentifier}");
        }

        if (TryGetTransitionSeconds(resolvedTransition, out var durationSeconds, "duration", "dur")) {
            sb.AppendLine($"{slideVariable}.TransitionDurationSeconds = {FormatDoubleLiteral(durationSeconds)}");
        }

        if (TryGetTransitionBoolean(resolvedTransition, out var advanceOnClick, "advance-on-click", "advanceonclick", "advance-click", "onclick", "click")) {
            sb.AppendLine($"{slideVariable}.TransitionAdvanceOnClick = ${advanceOnClick.ToString().ToLowerInvariant()}");
        }

        if (TryGetTransitionSeconds(resolvedTransition, out var advanceAfterSeconds, "advance-after", "advanceafter", "after", "delay")) {
            sb.AppendLine($"{slideVariable}.TransitionAdvanceAfterSeconds = {FormatDoubleLiteral(advanceAfterSeconds)}");
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

    private static void EmitDocument(OfficeMarkupDocument document, OfficeMarkupEmitterOptions options, StringBuilder sb) {
        sb.AppendLine("$document = New-OfficeWordDocument -FilePath " + options.FilePathVariable);
        var chartIndex = 0;
        foreach (var block in document.Blocks) {
            EmitDocumentBlock(block, "$document", sb, ref chartIndex);
        }

        sb.AppendLine("Save-OfficeWordDocument -Document $document");
    }

    private static void EmitDocumentBlock(OfficeMarkupBlock block, string documentVariable, StringBuilder sb, ref int chartIndex) {
        switch (block) {
            case OfficeMarkupHeadingBlock heading:
                sb.AppendLine($"Add-OfficeWordParagraph -Document {documentVariable} -Text {PsString(heading.Text)} -Style Heading{heading.Level}");
                break;
            case OfficeMarkupParagraphBlock paragraph:
                sb.AppendLine($"Add-OfficeWordParagraph -Document {documentVariable} -Text {PsString(paragraph.Text)}");
                break;
            case OfficeMarkupListBlock list:
                foreach (var item in list.Items) {
                    sb.AppendLine($"Add-OfficeWordParagraph -Document {documentVariable} -Text {PsString(item.Text)}");
                }

                break;
            case OfficeMarkupPageBreakBlock:
                sb.AppendLine($"Add-OfficeWordPageBreak -Document {documentVariable}");
                break;
            case OfficeMarkupHeaderFooterBlock headerFooter:
                sb.AppendLine($"Set-OfficeWord{NormalizeHeaderFooterKind(headerFooter.HeaderFooterKind)} -Document {documentVariable} -Text {PsString(headerFooter.Text)}");
                break;
            case OfficeMarkupTableOfContentsBlock:
                sb.AppendLine($"Add-OfficeWordTableOfContents -Document {documentVariable}");
                break;
            case OfficeMarkupSectionBlock section:
                sb.AppendLine($"# Section: {section.Name}");
                foreach (var child in section.Blocks) {
                    EmitDocumentBlock(child, documentVariable, sb, ref chartIndex);
                }

                break;
            case OfficeMarkupDiagramBlock diagram:
                sb.AppendLine($"# Render {diagram.Language} to an image and add it to the Word document.");
                break;
            case OfficeMarkupTableBlock table:
                sb.AppendLine($"Add-OfficeWordTable -Document {documentVariable} -Rows {Math.Max(1, table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0))} -Columns {Math.Max(1, table.Headers.Count > 0 ? table.Headers.Count : table.Rows.Select(row => row.Count).DefaultIfEmpty(1).Max())}");
                sb.AppendLine("# Fill table cells from the semantic table AST.");
                break;
            case OfficeMarkupChartBlock chart:
                EmitChartComment(chart, sb);
                if (EmitChartData(chart, $"$chartData{++chartIndex}", sb)) {
                    sb.AppendLine($"Add-OfficeWordChart -Document {documentVariable} -Type {PsString(chart.ChartType)} -Title {PsString(chart.Title ?? string.Empty)} -Data $chartData{chartIndex}");
                }

                break;
            default:
                sb.AppendLine($"# {block.Kind}: {Describe(block)}");
                break;
        }
    }

    private static void EmitWorkbook(OfficeMarkupDocument document, OfficeMarkupEmitterOptions options, StringBuilder sb) {
        sb.AppendLine("$workbook = New-OfficeExcelWorkbook -FilePath " + options.FilePathVariable);
        sb.AppendLine("$sheet = $null");
        sb.AppendLine("function Get-OrAddOfficeExcelWorksheet {");
        sb.AppendLine("    param($Workbook, [string] $Name)");
        sb.AppendLine("    $existing = Get-OfficeExcelWorksheet -Workbook $Workbook -Name $Name -ErrorAction SilentlyContinue");
        sb.AppendLine("    if ($null -ne $existing) { return $existing }");
        sb.AppendLine("    return Add-OfficeExcelWorksheet -Workbook $Workbook -Name $Name");
        sb.AppendLine("}");
        var chartIndex = 0;
        foreach (var block in document.Blocks) {
            switch (block) {
                case OfficeMarkupSheetBlock sheet:
                    sb.AppendLine($"$sheet = Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name {PsString(sheet.Name)}");
                    break;
                case OfficeMarkupRangeBlock range:
                    var (rangeWorksheetExpression, rangeAddress) = ResolveWorkbookTarget(range.Sheet, range.Address);
                    sb.AppendLine($"# Range {range.Address}");
                    var startRow = 1;
                    var startColumn = 1;
                    if (!TryParseCellAddress(rangeAddress, out startRow, out startColumn)) {
                        sb.AppendLine("# Could not parse range start. Values are emitted from row 1, column 1.");
                    }

                    for (int row = 0; row < range.Values.Count; row++) {
                        var values = range.Values[row];
                        for (int column = 0; column < values.Count; column++) {
                            sb.AppendLine($"Set-OfficeExcelCell -Worksheet {rangeWorksheetExpression} -Row {startRow + row} -Column {startColumn + column} -Value {PsString(values[column])}");
                        }
                    }

                    break;
                case OfficeMarkupFormulaBlock formula:
                    var (formulaWorksheetExpression, formulaCell) = ResolveWorkbookTarget(formula.Sheet, formula.Cell);
                    sb.AppendLine($"Set-OfficeExcelFormula -Worksheet {formulaWorksheetExpression} -Cell {PsString(formulaCell)} -Formula {PsString(formula.Expression)}");
                    break;
                case OfficeMarkupNamedTableBlock table:
                    table.Attributes.TryGetValue("sheet", out var tableSheet);
                    var (tableWorksheetExpression, tableRange) = ResolveWorkbookTarget(tableSheet, table.Range);
                    sb.AppendLine($"Add-OfficeExcelTable -Worksheet {tableWorksheetExpression} -Range {PsString(tableRange)} -Name {PsString(table.Name)} -HasHeader ${table.HasHeader.ToString().ToLowerInvariant()}");
                    break;
                case OfficeMarkupChartBlock chart:
                    EmitChartComment(chart, sb);
                    chart.Attributes.TryGetValue("cell", out var cell);
                    var (chartWorksheetExpression, chartCell) = ResolveWorkbookTarget(chart.Sheet, cell);
                    var chartRow = 1;
                    var chartColumn = 6;
                    if (!string.IsNullOrWhiteSpace(chartCell) && TryParseCellAddress(chartCell, out var parsedRow, out var parsedColumn)) {
                        chartRow = parsedRow;
                        chartColumn = parsedColumn;
                    }

                    if (!string.IsNullOrWhiteSpace(chart.Source)) {
                        sb.AppendLine($"Add-OfficeExcelChart -Worksheet {chartWorksheetExpression} -Type {PsString(chart.ChartType)} -Source {PsString(chart.Source!)} -Row {chartRow} -Column {chartColumn} -Title {PsString(chart.Title ?? string.Empty)}");
                    } else if (EmitChartData(chart, $"$chartData{++chartIndex}", sb)) {
                        sb.AppendLine($"Add-OfficeExcelChart -Worksheet {chartWorksheetExpression} -Type {PsString(chart.ChartType)} -Data $chartData{chartIndex} -Row {chartRow} -Column {chartColumn} -Title {PsString(chart.Title ?? string.Empty)}");
                    }

                    break;
                case OfficeMarkupFormattingBlock formatting:
                    formatting.Attributes.TryGetValue("sheet", out var formattingSheet);
                    var (formattingWorksheetExpression, formattingTarget) = ResolveWorkbookTarget(formattingSheet, formatting.Target);
                    EmitWorkbookFormatting(formatting, formattingWorksheetExpression, formattingTarget, sb);
                    break;
                default:
                    sb.AppendLine($"# {block.Kind}: {Describe(block)}");
                    break;
            }
        }

        sb.AppendLine("Save-OfficeExcelWorkbook -Workbook $workbook");
    }

    private static void EmitWorkbookFormatting(OfficeMarkupFormattingBlock formatting, string worksheetExpression, string target, StringBuilder sb) {
        var cells = EnumerateTargetCells(target).ToList();
        if (cells.Count == 0) {
            sb.AppendLine($"# Could not parse formatting target {target}. Style={formatting.Style ?? string.Empty} NumberFormat={formatting.NumberFormat ?? string.Empty}");
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
                sb.AppendLine($"{worksheetExpression}.FormatCell({row}, {column}, {PsString(formatting.NumberFormat!)})");
            }

            if (!string.IsNullOrWhiteSpace(fill)) {
                sb.AppendLine($"{worksheetExpression}.CellBackground({row}, {column}, {PsString(fill!)})");
            }

            if (!string.IsNullOrWhiteSpace(fontColor)) {
                sb.AppendLine($"{worksheetExpression}.CellFontColor({row}, {column}, {PsString(fontColor!)})");
            }

            if (!string.IsNullOrWhiteSpace(bold) && IsTruthy(bold!)) {
                sb.AppendLine($"{worksheetExpression}.CellBold({row}, {column}, $true)");
            }

            if (!string.IsNullOrWhiteSpace(italic) && IsTruthy(italic!)) {
                sb.AppendLine($"{worksheetExpression}.CellItalic({row}, {column}, $true)");
            }

            if (!string.IsNullOrWhiteSpace(underline) && IsTruthy(underline!)) {
                sb.AppendLine($"{worksheetExpression}.CellUnderline({row}, {column}, $true)");
            }

            if (TryGetHorizontalAlignmentIdentifier(alignment, out var alignmentIdentifier)) {
                sb.AppendLine($"{worksheetExpression}.CellAlign({row}, {column}, [DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues]::{alignmentIdentifier})");
            }

            if (TryGetVerticalAlignmentIdentifier(verticalAlignment, out var verticalAlignmentIdentifier)) {
                sb.AppendLine($"{worksheetExpression}.CellVerticalAlign({row}, {column}, [DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues]::{verticalAlignmentIdentifier})");
            }

            if (!string.IsNullOrWhiteSpace(wrap) && IsTruthy(wrap!)) {
                sb.AppendLine($"{worksheetExpression}.WrapCells({row}, {row}, {column})");
            }

            if (TryGetBorderStyleIdentifier(border, out var borderStyleIdentifier)) {
                var borderColorArgument = !string.IsNullOrWhiteSpace(borderColor)
                    ? $", {PsString(borderColor!)}"
                    : string.Empty;
                sb.AppendLine($"{worksheetExpression}.CellBorder({row}, {column}, [DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues]::{borderStyleIdentifier}{borderColorArgument})");
            }
        }
    }

    private static (string WorksheetExpression, string LocalReference) ResolveWorkbookTarget(string? explicitSheet, string? reference) {
        if (TrySplitSheetQualifiedReference(reference, out var sheetName, out var localReference)) {
            return (GetWorksheetExpression(sheetName), localReference);
        }

        if (!string.IsNullOrWhiteSpace(explicitSheet)) {
            return (GetWorksheetExpression(explicitSheet!), reference ?? string.Empty);
        }

        return (GetWorksheetExpression("Sheet1"), reference ?? string.Empty);
    }

    private static string GetWorksheetExpression(string sheetName) =>
        $"(Get-OrAddOfficeExcelWorksheet -Workbook $workbook -Name {PsString(sheetName)})";

    private static string NormalizeHeaderFooterKind(string kind) =>
        string.Equals(kind, "footer", StringComparison.OrdinalIgnoreCase) ? "Footer" : "Header";

    private static string Describe(OfficeMarkupBlock block) {
        switch (block) {
            case OfficeMarkupHeadingBlock heading:
                return heading.Text;
            case OfficeMarkupParagraphBlock paragraph:
                return paragraph.Text;
            case OfficeMarkupImageBlock image:
                return image.Source;
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

    private static string PsString(string value) {
        return "'" + (value ?? string.Empty).Replace("'", "''") + "'";
    }

    private static void EmitPlacementComment(OfficeMarkupPlacement? placement, StringBuilder sb) {
        if (placement == null || !placement.HasValue) {
            return;
        }

        sb.AppendLine($"# Placement: x={placement.X}, y={placement.Y}, w={placement.Width}, h={placement.Height}");
    }

    private static void EmitComment(StringBuilder sb, string text) {
        foreach (var line in (text ?? string.Empty).Replace("\r\n", "\n").Split('\n')) {
            sb.AppendLine($"# {line}");
        }
    }

    private static void EmitChartComment(OfficeMarkupChartBlock chart, StringBuilder sb) {
        if (chart.Attributes.Count > 0) {
            sb.AppendLine("# Chart options:");
            foreach (var attribute in chart.Attributes.OrderBy(attribute => attribute.Key, StringComparer.OrdinalIgnoreCase)) {
                sb.AppendLine($"#   {attribute.Key}: {attribute.Value}");
            }
        }

        if (chart.Data.Count > 0) {
            sb.AppendLine("# Inline chart data:");
            foreach (var row in chart.Data) {
                sb.AppendLine("#   " + string.Join(", ", row));
            }
        }
    }

    private static bool EmitChartData(OfficeMarkupChartBlock chart, string variableName, StringBuilder sb) {
        if (chart.Data.Count < 2 || chart.Data[0].Count < 2) {
            sb.AppendLine($"# Add {chart.ChartType} chart from source {chart.Source ?? string.Empty}.");
            return false;
        }

        var headers = chart.Data[0];
        var categories = chart.Data.Skip(1).Select(row => row.Count > 0 ? row[0] : string.Empty).ToList();
        sb.AppendLine($"{variableName} = @{{");
        sb.AppendLine($"    Categories = @({string.Join(", ", categories.Select(PsString))})");
        sb.AppendLine("    Series = @(");
        for (int seriesIndex = 1; seriesIndex < headers.Count; seriesIndex++) {
            var values = chart.Data.Skip(1).Select(row => PsValue(row.Count > seriesIndex ? row[seriesIndex] : "0"));
            var comma = seriesIndex == headers.Count - 1 ? string.Empty : ",";
            sb.AppendLine("        @{");
            sb.AppendLine($"            Name = {PsString(headers[seriesIndex])}");
            sb.AppendLine($"            Values = @({string.Join(", ", values)})");
            sb.AppendLine($"        }}{comma}");
        }

        sb.AppendLine("    )");
        sb.AppendLine("}");
        return true;
    }

    private static string PsValue(string value) =>
        double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var number)
            ? number.ToString("G", System.Globalization.CultureInfo.InvariantCulture)
            : PsString(value);

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

    private static string NormalizeToken(string value) =>
        new string((value ?? string.Empty).Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());

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

    private static bool TrySplitSheetQualifiedReference(string? reference, out string sheetName, out string localReference) {
        sheetName = string.Empty;
        localReference = string.Empty;
        if (string.IsNullOrWhiteSpace(reference)) {
            return false;
        }

        var value = reference!.Trim();
        var bangIndex = value.LastIndexOf('!');
        if (bangIndex <= 0 || bangIndex == value.Length - 1) {
            return false;
        }

        sheetName = value.Substring(0, bangIndex).Trim();
        if (sheetName.Length >= 2 && sheetName[0] == '\'' && sheetName[sheetName.Length - 1] == '\'') {
            sheetName = sheetName.Substring(1, sheetName.Length - 2).Replace("''", "'");
        }

        localReference = value.Substring(bangIndex + 1).Trim();
        return sheetName.Length > 0 && localReference.Length > 0;
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

    private static string FormatList(OfficeMarkupListBlock list) {
        var lines = list.Items.Select((item, index) => list.Ordered
            ? $"{list.Start + index}. {item.Text}"
            : $"- {item.Text}");
        return string.Join(Environment.NewLine, lines);
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
}
