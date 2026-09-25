using OfficeIMO.Drawing;
using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class HtmlPowerPointConverterExtensions {
    private const int MaximumGenericTextChunkCharacters = 4096;
    private const double GenericTextWidthPoints = 620D;

    private static bool TryImportGenericTextBlock(
        HtmlSemanticBlock block,
        PptCore.PowerPointPresentation presentation,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget,
        ref PptCore.PowerPointSlide slide,
        ref double contentTop,
        ref double pictureTop,
        double slideBottom) {
        string text = block.Text;
        if (text.Length == 0) return true;
        if (!budget.IsMetadataWithinLimit(text, out string metadataLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                "A slide text block was omitted because it exceeded the shared field limit.",
                lossKind: OfficeConversionLossKind.Omission, detail: metadataLimit);
            return true;
        }

        double minimumHeight = block.Kind == HtmlSemanticBlockKind.List
            ? Math.Max(52D, CountSemanticListItems(block) * 30D)
            : 52D;
        if (HasGenericAuthoredTextGeometry(block)) {
            if (NeedsGenericContinuation(contentTop, minimumHeight, slideBottom)) {
                if (!TryAddGenericSlide(presentation, result, budget, out slide)) return false;
                contentTop = pictureTop = 30D;
            }
            int previous = result.TextBoxes;
            contentTop = ImportTextBox(block.SourceElement, text, slide, contentTop, result, budget,
                minimumHeight, block);
            ReportGenericFormApproximation(block, result, result.TextBoxes > previous);
            return result.TextBoxes > previous;
        }

        double fontSize = GetGenericTextMeasureFontSize(block);
        OfficeFontStyle fontStyle = block.Runs.Any(run => run.Bold)
            ? OfficeFontStyle.Bold : OfficeFontStyle.Regular;
        OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(new OfficeFontInfo("Aptos", fontSize, fontStyle));
        Func<string?, double, double> measure = (value, size) =>
            measurer.MeasureWidth(value, measurer.CreateStyle(measurer.FallbackFontInfo.WithSize(size)))
            * 72D / OfficeTextMeasurer.DefaultDpi;
        int offset = 0;
        bool reportedStyleLoss = false;
        bool reportedListLoss = false;
        bool importedAny = false;
        while (offset < text.Length) {
            double applicableMinimumHeight = offset == 0 ? minimumHeight : 52D;
            double available = slideBottom - contentTop - 10D;
            if (available < 52D && contentTop > 30D) {
                if (!TryAddGenericSlide(presentation, result, budget, out slide)) return false;
                contentTop = pictureTop = 30D;
                continue;
            }

            int end = FindGenericTextChunkEnd(text, offset, available, fontSize, measure);
            if (end == text.Length && block.Kind == HtmlSemanticBlockKind.List
                && applicableMinimumHeight > available) {
                if (contentTop > 30D && applicableMinimumHeight <= slideBottom - 40D) {
                    if (!TryAddGenericSlide(presentation, result, budget, out slide)) return false;
                    contentTop = pictureTop = 30D;
                    continue;
                }
                int proportionalEnd = offset + Math.Max(1,
                    (int)Math.Floor((text.Length - offset) * available / applicableMinimumHeight));
                end = Math.Min(text.Length - 1, proportionalEnd);
                if (end > offset + 1) {
                    int wordBoundary = text.LastIndexOf(' ', end - 1, end - offset);
                    if (wordBoundary > offset + (end - offset) / 2) end = wordBoundary + 1;
                }
            }
            if (end < text.Length && contentTop > 30D && text.Length - offset <= MaximumGenericTextChunkCharacters
                && MeasureGenericTextHeight(text.Substring(offset), fontSize, measure, applicableMinimumHeight) <=
                slideBottom - 40D) {
                if (!TryAddGenericSlide(presentation, result, budget, out slide)) return false;
                contentTop = pictureTop = 30D;
                continue;
            }
            if (end <= offset) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentOmitted,
                    "A generic HTML text block could not fit within the visible slide area.",
                    lossKind: OfficeConversionLossKind.Omission,
                    detail: "remainingCharacters=" + (text.Length - offset));
                return true;
            }

            bool completeBlock = offset == 0 && end == text.Length;
            string chunk = text.Substring(offset, end - offset);
            double height = MeasureGenericTextHeight(chunk, fontSize, measure,
                completeBlock ? applicableMinimumHeight : 52D);
            int previousTextBoxes = result.TextBoxes;
            contentTop = ImportTextBox(completeBlock ? block.SourceElement : null, chunk, slide,
                contentTop, result, budget, height, completeBlock ? block : null);
            if (result.TextBoxes == previousTextBoxes) return false;
            importedAny = true;

            if (!completeBlock) {
                PptCore.PowerPointTextBox textBox = slide.TextBoxes.Last();
                if (!TryApplyGenericTextSliceRuns(textBox, block, offset, end) && !reportedStyleLoss) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                        "A paginated HTML text block could not retain its rich run formatting.",
                        lossKind: OfficeConversionLossKind.Approximation,
                        detail: "block=" + block.Kind + "; projection=paginatedText");
                    int omittedLinks = block.Runs.Count(run => !string.IsNullOrWhiteSpace(run.Hyperlink));
                    if (omittedLinks > 0) {
                        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentOmitted,
                            "Hyperlinks in a paginated HTML text block were not retained.",
                            lossKind: OfficeConversionLossKind.Omission,
                            detail: "hyperlinkRuns=" + omittedLinks);
                    }
                    reportedStyleLoss = true;
                }
                if (block.Kind == HtmlSemanticBlockKind.List && !reportedListLoss) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                        "A long HTML list was paginated as editable text without native list markers.",
                        lossKind: OfficeConversionLossKind.Approximation,
                        detail: "block=List; projection=paginatedText");
                    reportedListLoss = true;
                }
            }
            offset = end;
            if (offset < text.Length) {
                if (!TryAddGenericSlide(presentation, result, budget, out slide)) return false;
                contentTop = pictureTop = 30D;
            }
        }
        ReportGenericFormApproximation(block, result, importedAny);
        return true;
    }

    private static bool HasGenericAuthoredTextGeometry(HtmlSemanticBlock block) =>
        block.SourceElement.HasAttribute("data-officeimo-left")
        || block.SourceElement.HasAttribute("data-officeimo-top")
        || block.SourceElement.HasAttribute("data-officeimo-width")
        || block.SourceElement.HasAttribute("data-officeimo-height");

    private static double[] EstimateGenericTableRowHeights(HtmlSemanticBlock block, double tableWidth, HtmlImportBudget budget) {
        HtmlSemanticTable? source = block.Table;
        if (source == null) return Array.Empty<double>();
        int sourceRows = Math.Min(source.Rows.Count, budget.Limits.MaxTableCells);
        int columns = (int)Math.Min(Math.Max(1, budget.Limits.MaxTableCells),
            Math.Max(1L, source.Rows.Take(sourceRows).Select(row => row.Cells.Sum(cell => (long)Math.Max(1, cell.ColumnSpan)))
                .DefaultIfEmpty(1L).Max()));
        int rows = sourceRows;
        for (int rowIndex = 0; rowIndex < sourceRows; rowIndex++) {
            foreach (HtmlSemanticTableCell cell in source.Rows[rowIndex].Cells) {
                rows = Math.Max(rows, (int)Math.Min(budget.Limits.MaxTableCells,
                    (long)rowIndex + Math.Max(1, cell.RowSpan)));
            }
        }
        double[] rowHeights = Enumerable.Repeat(34D, rows).ToArray();
        double cellWidth = Math.Max(40D, tableWidth / columns - 16D);
        var measurers = new Dictionary<(double Size, OfficeFontStyle Style), OfficeTextMeasurer>();
        for (int rowIndex = 0; rowIndex < sourceRows; rowIndex++) {
            HtmlSemanticTableRow row = source.Rows[rowIndex];
            foreach (HtmlSemanticTableCell cell in row.Cells) {
                double fontSize = 18D;
                if (TryParseSemanticPixels(cell.Style?.GetValue("font-size"), out double cellPixels)) {
                    fontSize = Math.Max(fontSize, cellPixels * 0.75D);
                }
                foreach (HtmlSemanticRun run in cell.Runs) {
                    if (TryParseSemanticPixels(run.Style?.GetValue("font-size"), out double runPixels)) {
                        fontSize = Math.Max(fontSize, runPixels * 0.75D);
                    }
                }
                OfficeFontStyle fontStyle = cell.IsHeader || cell.Runs.Any(run => run.Bold)
                    ? OfficeFontStyle.Bold : OfficeFontStyle.Regular;
                if (!measurers.TryGetValue((fontSize, fontStyle), out OfficeTextMeasurer? measurer)) {
                    measurer = OfficeTextMeasurer.Create(new OfficeFontInfo("Aptos", fontSize, fontStyle));
                    measurers.Add((fontSize, fontStyle), measurer);
                }
                Func<string?, double, double> measure = (value, size) =>
                    measurer.MeasureWidth(value, measurer.CreateStyle(measurer.FallbackFontInfo.WithSize(size)))
                    * 72D / OfficeTextMeasurer.DefaultDpi;
                OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(cell.Text, fontSize,
                    cellWidth * Math.Min(columns, Math.Max(1, cell.ColumnSpan)), 100000D, 1.2D, 1D, measure,
                    wrap: true, forceSingleLine: false, shrinkToFit: false);
                int span = Math.Min(rows - rowIndex, Math.Max(1, cell.RowSpan));
                double heightPerRow = Math.Ceiling((layout.Height + 16D) / span);
                for (int spannedRow = rowIndex; spannedRow < rowIndex + span; spannedRow++) {
                    rowHeights[spannedRow] = Math.Max(rowHeights[spannedRow], heightPerRow);
                }
            }
        }
        return rowHeights;
    }

    private static double GetGenericTextMeasureFontSize(HtmlSemanticBlock block) {
        double size = 18D;
        foreach (HtmlSemanticRun run in block.Runs) {
            if (TryParseSemanticPixels(run.Style?.GetValue("font-size"), out double pixels)) {
                size = Math.Max(size, pixels * 0.75D);
            }
        }
        return size;
    }

    private static double MeasureGenericTextHeight(
        string text,
        double fontSize,
        Func<string?, double, double> measure,
        double minimumHeight) {
        OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(text, fontSize,
            GenericTextWidthPoints - 14.4D, 100000D, 1.35D, 1D, measure,
            wrap: true, forceSingleLine: false, shrinkToFit: false);
        return Math.Max(minimumHeight, Math.Ceiling(layout.Height + 20D));
    }

    private static int FindGenericTextChunkEnd(
        string text,
        int start,
        double availableHeight,
        double fontSize,
        Func<string?, double, double> measure) {
        int low = start;
        int high = Math.Min(text.Length, start + MaximumGenericTextChunkCharacters);
        while (low < high) {
            int middle = low + (high - low + 1) / 2;
            double height = MeasureGenericTextHeight(text.Substring(start, middle - start),
                fontSize, measure, 52D);
            if (height <= availableHeight) low = middle;
            else high = middle - 1;
        }
        if (low <= start) return start;
        if (low < text.Length) {
            int wordBoundary = text.LastIndexOf(' ', low - 1, low - start);
            if (wordBoundary > start + (low - start) / 2) low = wordBoundary + 1;
        }
        if (char.IsHighSurrogate(text[low - 1]) && low < text.Length) low--;
        return low;
    }

    private static bool TryApplyGenericTextSliceRuns(
        PptCore.PowerPointTextBox textBox,
        HtmlSemanticBlock block,
        int start,
        int end) {
        if (block.Runs.Count == 0 || textBox.Paragraphs.Count != 1
            || textBox.Text.IndexOf('\n') >= 0
            || !string.Equals(string.Concat(block.Runs.Select(run => run.Text)), block.Text,
                StringComparison.Ordinal)) return false;

        var slices = new List<(HtmlSemanticRun Run, string Text)>();
        int runStart = 0;
        foreach (HtmlSemanticRun run in block.Runs) {
            int runEnd = runStart + run.Text.Length;
            int sliceStart = Math.Max(runStart, start);
            int sliceEnd = Math.Min(runEnd, end);
            if (sliceEnd > sliceStart) {
                slices.Add((run, run.Text.Substring(sliceStart - runStart, sliceEnd - sliceStart)));
            }
            runStart = runEnd;
        }
        if (slices.Count == 0 || slices.Sum(slice => slice.Text.Length) != end - start) return false;

        PptCore.PowerPointParagraph paragraph = textBox.Paragraphs[0];
        paragraph.Text = slices[0].Text;
        ApplySemanticRun(paragraph.Runs[0], slices[0].Run, preserveTargetText: true);
        for (int index = 1; index < slices.Count; index++) {
            PptCore.PowerPointTextRun target = paragraph.AddRun(slices[index].Text);
            ApplySemanticRun(target, slices[index].Run, preserveTargetText: true);
        }
        return true;
    }

    private static void ReportGenericFormApproximation(
        HtmlSemanticBlock block,
        HtmlToPowerPointResult result,
        bool imported) {
        if (block.Kind == HtmlSemanticBlockKind.Form && imported) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                "An HTML form was imported as editable visible text without its interactive controls.",
                lossKind: OfficeConversionLossKind.Approximation,
                detail: "block=Form; preserved=visibleText; interaction=omitted");
        }
    }
}
