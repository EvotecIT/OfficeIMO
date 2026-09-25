using OfficeIMO.Html;
using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.OneNote.Html;

/// <summary>Imports prepared ordinary HTML into typed offline OneNote models.</summary>
public static class HtmlOneNoteConverterExtensions {
    private const string ComponentName = "OfficeIMO.OneNote.Html";

    /// <summary>Imports HTML as a OneNote section or throws when an error diagnostic is produced.</summary>
    public static OneNoteSection ToOneNoteSection(this HtmlConversionDocument document, HtmlToOneNoteOptions? options = null) =>
        Require(document.ToOneNoteSectionResult(options));

    /// <summary>Imports HTML as a OneNote section with structured diagnostics and counters.</summary>
    public static HtmlToOneNoteSectionResult ToOneNoteSectionResult(this HtmlConversionDocument document, HtmlToOneNoteOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlToOneNoteOptions resolved = options?.Clone() ?? new HtmlToOneNoteOptions();
        resolved.Limits.Validate();
        var section = new OneNoteSection { Name = CleanName(resolved.SectionName, "Imported") };
        var result = new HtmlToOneNoteSectionResult(section);
        HtmlSemanticDocument semanticDocument = document.SemanticDocument;
        AngleSharp.Html.Dom.IHtmlDocument sourceDocument = document.CreateNativeDocumentForConversion();
        foreach (HtmlDiagnostic diagnostic in document.Diagnostics) result.AddImportDiagnostic(diagnostic);
        ReportSkippedStylesheetLinks(sourceDocument, document.MediaContext, result);
        ReportUnpreservedColumnWidths(sourceDocument, result);
        ImportPages(semanticDocument, section, resolved, result);
        return result;
    }

    private static void ReportUnpreservedColumnWidths(
        AngleSharp.Html.Dom.IHtmlDocument sourceDocument, HtmlToOneNoteSectionResult result) {
        if (sourceDocument.QuerySelector("colgroup, col") == null) return;
        Add(result, HtmlConversionDiagnosticCodes.ContentApproximated,
            "HTML colgroup and col column definitions are not projected to OneNote table columns.",
            HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Approximation);
    }

    private static void ReportSkippedStylesheetLinks(
        AngleSharp.Html.Dom.IHtmlDocument sourceDocument,
        HtmlCssMediaContext mediaContext,
        HtmlToOneNoteSectionResult result) {
        var stylesheetOptions = new HtmlRenderOptions {
            Mode = mediaContext == HtmlCssMediaContext.Print ? HtmlRenderMode.Paged : HtmlRenderMode.Continuous
        };
        foreach (AngleSharp.Dom.IElement link in sourceDocument.QuerySelectorAll("link[href]")) {
            if (!HtmlRenderStylesheetApplier.IsApplicableStylesheetLink(link, stylesheetOptions)) continue;
            string href = link.GetAttribute("href") ?? string.Empty;
            if (string.IsNullOrWhiteSpace(href)) continue;
            result.AddImportDiagnostic(new HtmlDiagnostic(
                ComponentName,
                "HtmlStylesheetLinkSkipped",
                "An applicable HTML stylesheet link was not loaded by the semantic OneNote importer.",
                HtmlDiagnosticSeverity.Warning,
                source: href,
                lossKind: OfficeConversionLossKind.Omission));
        }
    }

    /// <summary>Imports HTML as a single-section OneNote notebook or throws on conversion errors.</summary>
    public static OneNoteNotebook ToOneNoteNotebook(this HtmlConversionDocument document, HtmlToOneNoteOptions? options = null) =>
        Require(document.ToOneNoteNotebookResult(options));

    /// <summary>Imports HTML as a single-section OneNote notebook with structured evidence.</summary>
    public static HtmlToOneNoteNotebookResult ToOneNoteNotebookResult(this HtmlConversionDocument document, HtmlToOneNoteOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlToOneNoteOptions resolved = options?.Clone() ?? new HtmlToOneNoteOptions();
        HtmlToOneNoteSectionResult sectionResult = document.ToOneNoteSectionResult(resolved);
        var notebook = new OneNoteNotebook { Name = CleanName(resolved.NotebookName, "Imported") };
        notebook.Sections.Add(sectionResult.Value);
        var result = new HtmlToOneNoteNotebookResult(notebook) {
            Sections = 1,
            Pages = sectionResult.Pages,
            Elements = sectionResult.Elements,
            Tables = sectionResult.Tables,
            Images = sectionResult.Images
        };
        foreach (HtmlDiagnostic diagnostic in sectionResult.Report.Diagnostics) result.AddImportDiagnostic(diagnostic);
        return result;
    }

    private static void ImportPages(
        HtmlSemanticDocument document,
        OneNoteSection target,
        HtmlToOneNoteOptions options,
        HtmlToOneNoteSectionResult result) {
        var budget = new HtmlImportBudget(options.Limits);
        foreach (HtmlSemanticSection projection in document.Sections) {
            if (!budget.TryReserveSemanticContainer(out string containerLimit)) {
                Add(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "Additional HTML sections were omitted because the shared page limit was reached.",
                    HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, containerLimit);
                break;
            }

            var page = new OneNotePage { Title = CleanName(projection.Title, "Imported") };
            var outline = new OneNoteOutline();
            foreach (HtmlSemanticBlock block in projection.Blocks) {
                ImportBlock(block, outline.Children, options, result, budget);
            }
            if (outline.Children.Count > 0) page.Outlines.Add(outline);
            target.Pages.Add(page);
            result.Pages++;
        }
    }

    private static void ImportBlock(
        HtmlSemanticBlock block,
        IList<OneNoteElement> target,
        HtmlToOneNoteOptions options,
        HtmlToOneNoteSectionResult result,
        HtmlImportBudget budget) {
        if (block.Kind == HtmlSemanticBlockKind.Table) {
            ImportTable(block, target, options, result, budget);
            return;
        }
        if (block.Kind == HtmlSemanticBlockKind.Image) {
            if (options.ImportImages && block.Resource != null) ImportImage(block.Resource, target, result, budget);
            return;
        }
        if (block.Kind == HtmlSemanticBlockKind.List) {
            ImportList(block, target, options, result, budget, Math.Max(0, block.Level - 1));
            return;
        }
        OneNoteParagraph? text = CreateParagraph(block, result, budget);
        if (text != null) target.Add(text);
        if (options.ImportImages) {
            foreach (HtmlSemanticResource resource in block.InlineResources.Where(item => item.Kind == HtmlResourceKind.Image)) {
                ImportImage(resource, target, result, budget);
            }
        }
    }

    private static void ImportList(
        HtmlSemanticBlock list,
        IList<OneNoteElement> target,
        HtmlToOneNoteOptions options,
        HtmlToOneNoteSectionResult result,
        HtmlImportBudget budget,
        int level) {
        int? previousOrdinal = null;
        foreach (HtmlSemanticBlock item in list.Children) {
            OneNoteParagraph? paragraph = CreateParagraph(item, result, budget);
            if (paragraph == null && item.Text.Length > 0) break;
            if (paragraph != null) {
                int? displayIndex = null;
                if (list.Ordered) {
                    int ordinal = NormalizeOneNoteListOrdinal(item.ListItem?.Ordinal ?? 1, result);
                    bool shouldRestart = !previousOrdinal.HasValue
                        || item.ListItem?.ExplicitOrdinal.HasValue == true
                        || list.List?.IsReversed == true
                        || (long)ordinal != (long)previousOrdinal.Value + 1L;
                    if (shouldRestart) displayIndex = ordinal;
                    previousOrdinal = ordinal;
                }
                paragraph.List = new OneNoteListInfo {
                    Ordered = list.Ordered,
                    Level = level,
                    DisplayIndex = displayIndex,
                    Restart = displayIndex.HasValue
                };
                target.Add(paragraph);
            }
            if (options.ImportImages) {
                foreach (HtmlSemanticResource resource in item.InlineResources.Where(candidate => candidate.Kind == HtmlResourceKind.Image)) {
                    ImportImage(resource, target, result, budget);
                }
            }
            foreach (HtmlSemanticBlock nested in item.Children.Where(child => child.Kind == HtmlSemanticBlockKind.List)) {
                ImportList(nested, target, options, result, budget, level + 1);
            }
        }
    }

    private static int NormalizeOneNoteListOrdinal(int ordinal, HtmlToOneNoteSectionResult result) {
        if (ordinal > 0) return ordinal;
        Add(result, HtmlConversionDiagnosticCodes.ContentApproximated,
            "A nonpositive HTML list ordinal was clamped to OneNote's minimum display index of 1.",
            HtmlDiagnosticSeverity.Warning,
            OfficeConversionLossKind.Approximation,
            "Ordinal=" + ordinal.ToString(CultureInfo.InvariantCulture) + "; Supported=1..2147483647");
        return 1;
    }

    private static OneNoteParagraph? CreateParagraph(
        HtmlSemanticBlock source,
        HtmlToOneNoteSectionResult result,
        HtmlImportBudget budget) {
        return CreateParagraph(source.Text, source.Runs, source.Kind == HtmlSemanticBlockKind.Heading ? source.Level : 0, source.Style, result, budget);
    }

    private static OneNoteParagraph? CreateParagraph(
        string plainText,
        IReadOnlyList<HtmlSemanticRun> runs,
        int headingLevel,
        HtmlComputedStyle? containerStyle,
        HtmlToOneNoteSectionResult result,
        HtmlImportBudget budget) {
        if (plainText.Length == 0) return null;
        if (!budget.IsMetadataWithinLimit(plainText, out string metadataLimit)) {
            Add(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                "An HTML text block was omitted because it exceeded the shared field limit.",
                HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Omission, metadataLimit);
            return null;
        }
        if (!budget.TryReserveShape(out string shapeLimit)) {
            Add(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "Additional HTML blocks were omitted because the shared element limit was reached.",
                HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Omission, shapeLimit);
            return null;
        }

        var paragraph = new OneNoteParagraph();
        foreach (HtmlSemanticRun sourceRun in runs) {
            var run = new OneNoteTextRun { Text = sourceRun.Text, Hyperlink = sourceRun.Hyperlink };
            if (sourceRun.DataAttributes.TryGetValue("data-officeimo-math-format", out string? mathFormat)
                && string.Equals(mathFormat, "latex", StringComparison.OrdinalIgnoreCase)) {
                try {
                    run.SetMathExpression(OfficeMathMarkup.FromLatex(sourceRun.Text));
                } catch (Exception exception) when (exception is FormatException || exception is ArgumentException) {
                    Add(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                        "Inline mathematical markup was retained as plain text because its LaTeX payload could not be parsed.",
                        HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Approximation,
                        exception.Message);
                }
            }
            run.Style.Bold = sourceRun.Bold ? true : null;
            run.Style.Italic = sourceRun.Italic ? true : null;
            run.Style.Underline = sourceRun.Underline ? true : null;
            run.Style.Strikethrough = sourceRun.Strikethrough ? true : null;
            run.Style.Superscript = sourceRun.Superscript ? true : null;
            run.Style.Subscript = sourceRun.Subscript ? true : null;
            run.Style.FontFamily = HtmlRenderCssValues.FirstFontFamily(sourceRun.Style?.GetValue("font-family"));
            double? resolvedFontSize = sourceRun.Style?.ResolvedFontSizePoints ?? containerStyle?.ResolvedFontSizePoints;
            if (resolvedFontSize.HasValue && resolvedFontSize.Value > 0D) {
                run.Style.FontSize = resolvedFontSize.Value;
            } else if (TryParseCssPoints(sourceRun.Style?.GetValue("font-size"), out double fontSize)) {
                run.Style.FontSize = fontSize;
            }
            if (TryParseArgb(sourceRun.Style?.GetValue("color"), out uint foreground)) {
                run.Style.ColorArgb = foreground;
            }
            if (TryParseArgb(sourceRun.BackgroundColor, out uint highlight)) {
                run.Style.HighlightColorArgb = highlight;
            }
            paragraph.Runs.Add(run);
        }
        if (paragraph.Runs.Count == 0) paragraph.Runs.Add(new OneNoteTextRun { Text = plainText });
        TrimRuns(paragraph);
        if (headingLevel > 0) paragraph.Style.StyleId = "Heading" + Math.Min(6, headingLevel);
        result.Elements++;
        return paragraph;
    }

    private static bool TryParseCssPoints(string? value, out double points) {
        points = 0D;
        string text = (value ?? string.Empty).Trim().ToLowerInvariant();
        double multiplier;
        if (text.EndsWith("pt", StringComparison.Ordinal)) {
            multiplier = 1D;
            text = text.Substring(0, text.Length - 2).Trim();
        } else if (text.EndsWith("px", StringComparison.Ordinal)) {
            multiplier = 0.75D;
            text = text.Substring(0, text.Length - 2).Trim();
        } else {
            return false;
        }
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            && parsed > 0D
            && (points = parsed * multiplier) > 0D;
    }

    private static bool TryParseArgb(string? value, out uint argb) {
        argb = 0U;
        if (!OfficeColor.TryParseCss(value, out OfficeColor color)) return false;
        argb = ((uint)color.A << 24) | ((uint)color.R << 16) | ((uint)color.G << 8) | color.B;
        return true;
    }

    private static void ImportTable(
        HtmlSemanticBlock source,
        IList<OneNoteElement> target,
        HtmlToOneNoteOptions options,
        HtmlToOneNoteSectionResult result,
        HtmlImportBudget budget) {
        HtmlSemanticTable? sourceTable = source.Table;
        if (sourceTable == null || sourceTable.Rows.Count == 0) return;

        if (!budget.TryReserveTableWithShape(out string tableLimit)) {
            Add(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "An HTML table was omitted because the shared import limit was reached.",
                HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Omission,
                tableLimit);
            return;
        }
        var table = new OneNoteTable { BordersVisible = true };
        int cells = 0;
        int maxTableCells = budget.Limits.MaxTableCells;
        foreach (HtmlSemanticTableRow rowElement in sourceTable.Rows) {
            var row = new OneNoteTableRow();
            foreach (HtmlSemanticTableCell cellElement in rowElement.Cells) {
                if (++cells > maxTableCells) {
                    Add(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "Remaining HTML table cells were omitted because the configured table limit was reached.",
                        HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Omission,
                        "limit=" + maxTableCells);
                    break;
                }
                var cell = new OneNoteTableCell();
                if (TryParseArgb(cellElement.Style?.GetValue("background-color"), out uint shading)) {
                    cell.ShadingColorArgb = shading;
                }
                OneNoteParagraph? paragraph = CreateParagraph(cellElement.Text, cellElement.Runs, 0, cellElement.Style, result, budget);
                if (paragraph != null) cell.Content.Add(paragraph);
                if (options.ImportImages) {
                    foreach (HtmlSemanticResource resource in cellElement.Resources.Where(item => item.Kind == HtmlResourceKind.Image)) {
                        ImportImage(resource, cell.Content, result, budget);
                    }
                }
                row.Cells.Add(cell);
            }
            if (row.Cells.Count > 0) table.Rows.Add(row);
            if (cells >= maxTableCells) break;
        }
        if (table.Rows.Count == 0) return;
        SetImportedTableColumnWidths(source, table, result);
        target.Add(table);
        result.Elements++;
        result.Tables++;
    }

    private static void SetImportedTableColumnWidths(
        HtmlSemanticBlock source, OneNoteTable table, HtmlToOneNoteSectionResult result) {
        int columns = table.Rows.Max(row => row.Cells.Count);
        if (columns == 0) return;

        // Native writing otherwise assigns one half-inch unit to every column,
        // making ordinary imported tables unreadably narrow after reopen.
        const double defaultWidthHalfInches = 15D;
        double totalWidth = defaultWidthHalfInches;
        string? authoredWidth = source.Style?.GetValue("width");
        bool hasAuthoredTableWidth = TryParseTableWidth(authoredWidth, defaultWidthHalfInches, out double requestedWidth);
        if (hasAuthoredTableWidth) {
            totalWidth = Math.Min(defaultWidthHalfInches, Math.Max(columns, requestedWidth));
            if (totalWidth != requestedWidth) ReportTableWidthApproximation(result, authoredWidth);
        } else if (IsAuthoredWidth(authoredWidth)) {
            ReportTableWidthApproximation(result, authoredWidth);
        }
        totalWidth = Math.Max(columns, totalWidth);

        var weights = new double[columns];
        foreach (OneNoteTableRow row in table.Rows) {
            for (int column = 0; column < row.Cells.Count; column++) {
                int characters = 0;
                foreach (OneNoteTextRun run in row.Cells[column].Content.OfType<OneNoteParagraph>()
                    .SelectMany(paragraph => paragraph.Runs)) {
                    characters = Math.Min(256, characters + Math.Min(256, run.Text.Length));
                    if (characters == 256) break;
                }
                weights[column] = Math.Max(weights[column], Math.Sqrt(Math.Min(256, characters) + 1D));
            }
        }

        var widths = new double[columns];
        var reportedConflicts = new bool[columns];
        double explicitTotal = 0D;
        for (int row = 0; row < table.Rows.Count; row++) {
            HtmlSemanticTableRow sourceRow = source.Table!.Rows[row];
            for (int column = 0; column < Math.Min(columns, sourceRow.Cells.Count); column++) {
                HtmlSemanticTableCell cell = sourceRow.Cells[column];
                string? cellWidth = cell.Style?.GetValue("width");
                if (!IsAuthoredWidth(cellWidth)) continue;
                if (cell.ColumnSpan != 1 || !TryParseTableWidth(cellWidth, totalWidth, out double requestedColumnWidth)) {
                    ReportTableWidthApproximation(result, cellWidth);
                    continue;
                }
                if (widths[column] == 0D) {
                    widths[column] = Math.Max(1D, requestedColumnWidth);
                    explicitTotal += widths[column];
                    if (widths[column] != requestedColumnWidth) ReportTableWidthApproximation(result, cellWidth);
                } else if (!reportedConflicts[column]
                    && Math.Abs(widths[column] - requestedColumnWidth) > 0.01D) {
                    ReportTableWidthApproximation(result, cellWidth);
                    reportedConflicts[column] = true;
                }
            }
        }

        int inferredColumns = widths.Count(width => width == 0D);
        if (explicitTotal > totalWidth - inferredColumns) {
            Add(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                "Authored HTML table columns exceeded OneNote's available page width and were scaled.",
                HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Approximation);
            double available = Math.Max(0D, totalWidth - columns);
            double excess = explicitTotal - (columns - inferredColumns);
            for (int column = 0; column < columns; column++) {
                if (widths[column] > 0D) widths[column] = 1D + (excess > 0D ? available * (widths[column] - 1D) / excess : 0D);
            }
            explicitTotal = widths.Sum();
        } else if (hasAuthoredTableWidth && inferredColumns == 0 && explicitTotal < totalWidth) {
            Add(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                "Authored HTML table columns left unused table width, which was distributed proportionally.",
                HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Approximation);
            for (int column = 0; column < columns; column++) widths[column] *= totalWidth / explicitTotal;
            explicitTotal = totalWidth;
        }

        double distributable = Math.Max(0D, totalWidth - explicitTotal - inferredColumns);
        double weightSum = 0D;
        for (int column = 0; column < columns; column++) {
            if (widths[column] == 0D) weightSum += weights[column];
        }
        for (int column = 0; column < columns; column++) {
            if (widths[column] == 0D) {
                widths[column] = 1D + (weightSum > 0D
                    ? distributable * weights[column] / weightSum
                    : distributable / inferredColumns);
            }
            table.ColumnWidths.Add(widths[column]);
        }
    }

    private static bool TryParseTableWidth(string? value, double referenceHalfInches, out double widthHalfInches) {
        widthHalfInches = 0D;
        if (TryParseCssPoints(value, out double points)) {
            widthHalfInches = points / 36D;
        } else {
            string text = (value ?? string.Empty).Trim();
            if (!text.EndsWith("%", StringComparison.Ordinal)
                || !double.TryParse(text.Substring(0, text.Length - 1), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out double percent)) return false;
            widthHalfInches = referenceHalfInches * percent / 100D;
        }
        return widthHalfInches > 0D && !double.IsInfinity(widthHalfInches) && !double.IsNaN(widthHalfInches);
    }

    private static bool IsAuthoredWidth(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return false;
        return !string.Equals(value!.Trim(), "auto", StringComparison.OrdinalIgnoreCase);
    }

    private static void ReportTableWidthApproximation(HtmlToOneNoteSectionResult result, string? authoredWidth) =>
        Add(result, HtmlConversionDiagnosticCodes.ContentApproximated,
            "An authored HTML table width could not be retained in OneNote's bounded page layout.",
            HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Approximation, authoredWidth);

    private static void ImportImage(
        HtmlSemanticResource resource,
        IList<OneNoteElement> target,
        HtmlToOneNoteSectionResult result,
        HtmlImportBudget budget) {
        if (!HtmlImageDataUri.TryParse(resource.Source, out HtmlImageDataUri dataUri)) {
            Add(result, HtmlConversionDiagnosticCodes.ResourceTypeUnsupported,
                "An HTML image was omitted because native import requires a bounded image data URI.",
                HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Omission, resource.Source);
            return;
        }
        if (!budget.TryReserveImageWithShape(dataUri, out HtmlImportBudgetReservation imageReservation, out string imageLimit)) {
            Add(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "An embedded HTML image was omitted because the shared import limit was reached.",
                HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Omission,
                imageLimit);
            return;
        }
        using HtmlImportBudgetReservation imageReservationScope = imageReservation;
        if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
            Add(result, HtmlConversionDiagnosticCodes.ResourceDecodeFailed,
                "An embedded HTML image could not be decoded.", HtmlDiagnosticSeverity.Warning, OfficeConversionLossKind.Omission);
            return;
        }
        target.Add(new OneNoteImage {
            AltText = resource.AlternateText,
            Hyperlink = resource.Hyperlink,
            MediaType = dataUri.MediaType,
            FileName = "image" + dataUri.FileExtension,
            Payload = OneNoteBinaryPayload.FromBytes(bytes)
        });
        result.Elements++;
        result.Images++;
        imageReservation.Commit();
    }

    private static void TrimRuns(OneNoteParagraph paragraph) {
        while (paragraph.Runs.Count > 0 && string.IsNullOrWhiteSpace(paragraph.Runs[0].Text)) paragraph.Runs.RemoveAt(0);
        while (paragraph.Runs.Count > 0 && string.IsNullOrWhiteSpace(paragraph.Runs[paragraph.Runs.Count - 1].Text)) paragraph.Runs.RemoveAt(paragraph.Runs.Count - 1);
        if (paragraph.Runs.Count > 0) paragraph.Runs[0].Text = paragraph.Runs[0].Text.TrimStart();
        if (paragraph.Runs.Count > 0) paragraph.Runs[paragraph.Runs.Count - 1].Text = paragraph.Runs[paragraph.Runs.Count - 1].Text.TrimEnd();
    }

    private static void Add(
        HtmlToOneNoteSectionResult result,
        string code,
        string message,
        HtmlDiagnosticSeverity severity,
        OfficeConversionLossKind lossKind,
        string? detail = null) =>
        result.AddImportDiagnostic(new HtmlDiagnostic(ComponentName, code, message, severity, detail: detail, lossKind: lossKind));

    private static T Require<T>(HtmlConversionResult<T> result) where T : class {
        if (result.Succeeded) return result.Value;
        throw new HtmlConversionException(result.Report.Diagnostics);
    }

    private static string CleanName(string? value, string fallback) {
        string normalized = string.Join(" ", (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        return normalized.Length == 0 ? fallback : normalized;
    }

}
