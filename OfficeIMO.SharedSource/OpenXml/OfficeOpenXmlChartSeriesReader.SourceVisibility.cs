using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        /// <summary>The cached points a visible-cells-only chart plots, by data source (c:val, c:cat, c:xVal, c:yVal,
        /// c:bubbleSize, and the c:plus and c:minus amounts of custom error bars).</summary>
        internal sealed class SourceVisibility {
            private readonly Dictionary<OpenXmlElement, int[]> _visible;

            internal SourceVisibility(Dictionary<OpenXmlElement, int[]> visible) => _visible = visible;

            /// <summary>Keeps the cached points whose source cells are visible, closing the gaps as Office does.</summary>
            internal static IReadOnlyList<T> Select<T>(SourceVisibility? visibility, OpenXmlElement? source, IReadOnlyList<T> values) {
                if (visibility == null || source == null || !visibility._visible.TryGetValue(source, out int[]? offsets)) return values;
                var visible = new List<T>(offsets.Length);
                foreach (int offset in offsets) {
                    if (offset < values.Count) visible.Add(values[offset]);
                }
                return visible;
            }

            internal static bool Filters(SourceVisibility? visibility, OpenXmlElement? source) =>
                visibility != null && source != null && visibility._visible.ContainsKey(source);
        }

        /// <summary>Resolves which cached points a chart plots when it shows visible cells only. A chart that
        /// shows hidden cells, has no workbook source or has nothing hidden in its ranges plots its caches as written
        /// and returns a null <paramref name="visibility"/>.</summary>
        /// <returns>False with a reason when hidden source cells cannot be mapped to the cached points.</returns>
        internal static bool TryReadSourceVisibility(ChartPart part, C.Chart chart, out SourceVisibility? visibility, out string? reason) {
            visibility = null;
            reason = null;
            if (!OfficeOpenXmlChartSourceVisibility.PlotsVisibleCellsOnly(chart)) return true;
            // Literal cached charts have no workbook source whose hidden rows can affect plotting.
            if (chart.PlotArea?.Descendants<C.Formula>().Any(formula => !string.IsNullOrWhiteSpace(formula.Text)) != true) return true;
            try {
                // Without any workbook reference the cache is the only source, and Office plots it
                // as written. An unresolved or external workbook link fails closed.
                var embedded = OfficeOpenXmlChartWriter.GetSharedEmbeddedWorkbook(part);
                if (embedded == null) return true;
                using var stream = embedded.GetStream(FileMode.Open, FileAccess.Read);
                byte[] bytes = OfficeOpenXmlChartWorkbookSecurity.ReadAndValidate(stream);
                using var workbookStream = new MemoryStream(bytes, writable: false);
                using var workbook = SpreadsheetDocument.Open(workbookStream, false, OfficeOpenXmlChartWorkbookSecurity.CreateOpenSettings());
                if (workbook.WorkbookPart == null) return Fail("The embedded chart workbook has no workbook part.", out reason);
                var sheets = new Dictionary<string, WorksheetPart>(StringComparer.OrdinalIgnoreCase);
                foreach (S.Sheet sheet in workbook.WorkbookPart.Workbook?.Sheets?.Elements<S.Sheet>() ?? Enumerable.Empty<S.Sheet>()) {
                    if (sheet.Name?.Value == null || sheet.Id?.Value == null ||
                        workbook.WorkbookPart.GetPartById(sheet.Id.Value) is not WorksheetPart worksheetPart ||
                        worksheetPart.Worksheet == null || sheets.ContainsKey(sheet.Name.Value))
                        return Fail("The embedded chart workbook has an unresolved or duplicate sheet.", out reason);
                    sheets.Add(sheet.Name.Value, worksheetPart);
                }
                var hiddenBySheet = new Dictionary<WorksheetPart, OfficeOpenXmlChartHiddenCells>();
                foreach (WorksheetPart worksheetPart in sheets.Values.Distinct())
                    hiddenBySheet.Add(worksheetPart, new OfficeOpenXmlChartHiddenCells(worksheetPart.Worksheet!));
                if (!hiddenBySheet.Values.Any(hidden => hidden.HasHiddenCells)) return true;

                C.Formula[] formulas = chart.PlotArea!.Descendants<C.Formula>().Take(10001).ToArray();
                if (formulas.Length > 10000) return Fail("The chart has more source references than are supported.", out reason);
                var visible = new Dictionary<OpenXmlElement, int[]>();
                long cells = 0;
                foreach (C.Formula formula in formulas) {
                    if (string.IsNullOrWhiteSpace(formula.Text)) continue;
                    if (!OfficeOpenXmlChartSourceVisibility.TryParseRange(formula.Text, out string sheetName,
                        out uint firstRow, out uint lastRow, out uint firstColumn, out uint lastColumn))
                        return Fail("A chart source reference is not a single worksheet range.", out reason);
                    if (!sheets.TryGetValue(sheetName, out WorksheetPart? worksheetPart))
                        return Fail("A chart source reference names a sheet the embedded workbook does not have.", out reason);
                    OfficeOpenXmlChartHiddenCells hidden = hiddenBySheet[worksheetPart];
                    if (!hidden.Intersects(firstRow, lastRow, firstColumn, lastColumn)) continue;
                    OpenXmlElement? reference = formula.Parent;
                    OpenXmlElement? source = reference?.Parent;
                    // Office keeps a series whose name cell is hidden.
                    if (source is C.SeriesText && source.Parent?.LocalName == "ser") continue;
                    // Custom error amounts lose their hidden cells with the points they belong to.
                    bool seriesSource = source is C.Values or C.CategoryAxisData or C.XValues or C.YValues or C.BubbleSize &&
                        source.Parent?.LocalName == "ser";
                    bool errorAmounts = source is C.Plus or C.Minus && source.Parent is C.ErrorBars errorBars &&
                        errorBars.Parent?.LocalName == "ser";
                    if (source == null || reference is not (C.NumberReference or C.StringReference) || !seriesSource && !errorAmounts)
                        return Fail("A hidden chart source cell feeds chart content that cannot leave it out.", out reason);
                    long count = ((long)lastRow - firstRow + 1) * ((long)lastColumn - firstColumn + 1);
                    cells += count;
                    if (cells > OfficeOpenXmlChartWriter.MaximumSharedChartPoints)
                        return Fail("The chart source ranges exceed the supported point limit.", out reason);
                    // Each cached point maps to one cell only when the cache covers the whole range.
                    if (reference.Descendants<C.PointCount>().FirstOrDefault()?.Val?.Value != count)
                        return Fail("A chart cache does not cover its source range, so hidden cells cannot be matched to points.", out reason);
                    visible[source] = hidden.GetVisibleOffsets(firstRow, lastRow, firstColumn, lastColumn)!;
                }
                if (visible.Count == 0) return true;
                // Leaving cells out of one series shortens it, and Office draws no point past its end. The flat
                // projection pads shorter series with missing values, which only a zero blank rule would draw.
                if (OfficeOpenXmlChartPlotSemantics.ReadBlankDisplay(chart) == OfficeIMO.Drawing.OfficeChartBlankDisplay.Zero) {
                    long Length(OpenXmlElement source) => visible.TryGetValue(source, out int[]? offsets)
                        ? offsets.Length
                        : (long?)source.Descendants<C.PointCount>().FirstOrDefault()?.Val?.Value ?? 0L;
                    // A series whose value cells are all hidden is left out with its categories, so it needs no
                    // padding and sets no length; the others are padded to the longest series or category list kept.
                    var kept = chart.PlotArea!.Descendants<C.Values>()
                        .Select(values => (Values: Length(values), Categories: values.Parent?.GetFirstChild<C.CategoryAxisData>() is C.CategoryAxisData categories ? Length(categories) : 0L))
                        .Where(item => item.Values != 0L).ToArray();
                    long projected = kept.Select(item => Math.Max(item.Values, item.Categories)).DefaultIfEmpty(0L).Max();
                    if (kept.Any(item => item.Values < projected))
                        return Fail("Hidden cells leave chart series shorter than the chart, which a zero blank rule would draw.", out reason);
                }
                visibility = new SourceVisibility(visible);
                return true;
            } catch (Exception exception) {
                return Fail("The chart's embedded workbook cannot be read to leave out hidden cells: " + exception.Message, out reason);
            }
        }

        private static bool Fail(string message, out string? reason) {
            reason = message;
            return false;
        }
    }
}
