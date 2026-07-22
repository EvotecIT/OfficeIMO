using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

/// <summary>
/// Extension methods enabling HTML conversions for OfficeIMO Excel documents.
/// </summary>
public static partial class ExcelHtmlConverterExtensions {
    /// <summary>
    /// Converts a workbook to HTML.
    /// </summary>
    public static string ToHtml(this ExcelDocument workbook, ExcelHtmlSaveOptions? options = null) {
        return workbook.ToHtmlResult(options).Value;
    }

    /// <summary>Converts a workbook to HTML with the shared structured result contract.</summary>
    public static HtmlTextConversionResult ToHtmlResult(this ExcelDocument workbook, ExcelHtmlSaveOptions? options = null) {
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));
        options ??= new ExcelHtmlSaveOptions();
        options.Validate();
        string html = options.Profile == OfficeHtmlConversionProfile.ExcelVisualReview
            ? ConvertWorkbookVisual(workbook, options)
            : ConvertWorkbookSemantic(workbook, options);
        return new HtmlTextConversionResult(html);
    }

    /// <summary>
    /// Converts a worksheet to HTML.
    /// </summary>
    public static string ToHtml(this ExcelSheet sheet, ExcelHtmlSaveOptions? options = null) {
        return sheet.ToHtmlResult(options).Value;
    }

    /// <summary>Converts a worksheet to HTML with the shared structured result contract.</summary>
    public static HtmlTextConversionResult ToHtmlResult(this ExcelSheet sheet, ExcelHtmlSaveOptions? options = null) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        options ??= new ExcelHtmlSaveOptions();
        options.Validate();
        string html = options.Profile == OfficeHtmlConversionProfile.ExcelVisualReview
            ? ConvertSheetVisual(sheet, options)
            : ConvertSheetSemantic(sheet, options);
        return new HtmlTextConversionResult(html);
    }

    /// <summary>
    /// Saves a workbook as HTML.
    /// </summary>
    public static void SaveAsHtml(this ExcelDocument workbook, string path, ExcelHtmlSaveOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("HTML path cannot be empty.", nameof(path));
        HtmlTextIO.Write(path, workbook.ToHtml(options));
    }

    /// <summary>
    /// Saves a worksheet as HTML.
    /// </summary>
    public static void SaveAsHtml(this ExcelSheet sheet, string path, ExcelHtmlSaveOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("HTML path cannot be empty.", nameof(path));
        HtmlTextIO.Write(path, sheet.ToHtml(options));
    }

    private static string ConvertWorkbookSemantic(ExcelDocument workbook, ExcelHtmlSaveOptions options) {
        var body = new StringBuilder();
        body.Append("<main class=\"officeimo-document\"");
        OfficeHtmlSemanticEnvelope.AppendRootAttributes(body, "excel", options.Profile.ToString());
        body.Append('>');
        body.Append("<h1>").Append(OfficeHtmlText.Escape(GetTitle(options, "Excel Workbook"))).Append("</h1>");
        foreach (ExcelSheet sheet in workbook.Sheets) {
            AppendSheetTable(body, sheet, options);
        }

        body.Append("</main>");
        return Wrap(body.ToString(), options, GetTitle(options, "Excel Workbook"));
    }

    private static string ConvertSheetSemantic(ExcelSheet sheet, ExcelHtmlSaveOptions options) {
        var body = new StringBuilder();
        body.Append("<main class=\"officeimo-document\"");
        OfficeHtmlSemanticEnvelope.AppendRootAttributes(body, "excel", options.Profile.ToString());
        body.Append('>');
        AppendSheetTable(body, sheet, options);
        body.Append("</main>");
        return Wrap(body.ToString(), options, GetTitle(options, sheet.Name));
    }

    private static string ConvertWorkbookVisual(ExcelDocument workbook, ExcelHtmlSaveOptions options) {
        var body = new StringBuilder();
        body.Append("<main class=\"officeimo-document\"");
        OfficeHtmlSemanticEnvelope.AppendRootAttributes(body, "excel", options.Profile.ToString());
        body.Append('>');
        body.Append("<h1>").Append(OfficeHtmlText.Escape(GetTitle(options, "Excel Visual Review"))).Append("</h1>");
        ExcelWorkbookImageExportOptions visualOptions = ResolveWorkbookVisualOptions(options.VisualOptions);
        Dictionary<string, ExcelSheet> sheetsByName = workbook.Sheets.ToDictionary(sheet => sheet.Name, StringComparer.OrdinalIgnoreCase);
        int svgIndex = 0;
        foreach (OfficeImageExportResult result in workbook.ExportImages(OfficeImageExportFormat.Svg, visualOptions)) {
            AppendSvgResult(body, result, CreateSvgNamespacePrefix(result, ++svgIndex));
            string? resultName = result.Name;
            if (!string.IsNullOrWhiteSpace(resultName) && sheetsByName.TryGetValue(resultName!, out ExcelSheet? sheet)) {
                AppendVisualCommentInventory(body, sheet.GetComments());
            }
        }

        body.Append("</main>");
        return Wrap(body.ToString(), options, GetTitle(options, "Excel Visual Review"));
    }

    private static string ConvertSheetVisual(ExcelSheet sheet, ExcelHtmlSaveOptions options) {
        var body = new StringBuilder();
        body.Append("<main class=\"officeimo-document\"");
        OfficeHtmlSemanticEnvelope.AppendRootAttributes(body, "excel", options.Profile.ToString());
        body.Append('>');
        body.Append("<h1>").Append(OfficeHtmlText.Escape(GetTitle(options, sheet.Name))).Append("</h1>");
        AppendSvgResult(body, sheet.ExportImage(OfficeImageExportFormat.Svg, ToWorksheetOptions(ResolveWorkbookVisualOptions(options.VisualOptions))), "officeimo-sheet-svg-1-");
        AppendVisualCommentInventory(body, sheet.GetComments());
        body.Append("</main>");
        return Wrap(body.ToString(), options, GetTitle(options, sheet.Name));
    }

    private static void AppendSheetTable(StringBuilder body, ExcelSheet sheet, ExcelHtmlSaveOptions options) {
        int rowLimit = options.MaxRowsPerSheet ?? ExcelHtmlSaveOptions.DefaultMaxRowsPerSheet;
        int columnLimit = options.MaxColumnsPerSheet ?? ExcelHtmlSaveOptions.DefaultMaxColumnsPerSheet;
        IReadOnlyList<ExcelMergedRangeSnapshot> mergedRanges = sheet.GetMergedRanges(options.MaxMergedRangesPerSheet);
        string reportedUsedRange = sheet.GetUsedRangeA1();
        bool isEmptyDefaultRange = mergedRanges.Count == 0
            && (string.Equals(reportedUsedRange, "A1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(reportedUsedRange, "A1:A1", StringComparison.OrdinalIgnoreCase))
            && !sheet.TryGetCellText(1, 1, out _);
        string usedRange = ExpandUsedRangeForMerges(
            sheet,
            reportedUsedRange,
            mergedRanges,
            rowLimit,
            columnLimit,
            options.MaxCellsPerSheet);
        body.Append("<section class=\"officeimo-sheet\" data-officeimo-sheet=\"")
            .Append(OfficeHtmlText.EscapeAttribute(sheet.Name))
            .Append("\" data-officeimo-range=\"")
            .Append(OfficeHtmlText.EscapeAttribute(usedRange))
            .Append('"');
        AppendSheetVisibilityAttribute(body, sheet);
        body.Append('>');
        body.Append("<h2>").Append(OfficeHtmlText.Escape(sheet.Name)).Append("</h2>");

        ParseUsedRange(usedRange, out int firstRow, out int firstColumn, out int rowCount, out int columnCount);
        int maxColumns = Math.Min(Math.Min(columnCount, columnLimit), options.MaxCellsPerSheet);
        int maxRows = Math.Min(rowCount, rowLimit);
        if (maxColumns > 0 && (long)maxRows * maxColumns > options.MaxCellsPerSheet) {
            maxRows = Math.Min(maxRows, Math.Max(1, options.MaxCellsPerSheet / maxColumns));
        }
        ExcelMergeExportMap mergeMap = BuildMergeExportMap(mergedRanges, firstRow, firstColumn, maxRows, maxColumns);
        if (rowCount == 0 || columnCount == 0 || maxRows == 0 || maxColumns == 0 || (!SheetHasUsedCells(sheet, firstRow, firstColumn, maxRows, maxColumns) && mergeMap.Count == 0)) {
            body.Append(isEmptyDefaultRange
                ? "<p class=\"officeimo-muted\">No used cells.</p>"
                : "<p class=\"officeimo-muted\">No cells within export limits.</p>");
            AppendSheetTruncationDiagnostics(body, maxRows, rowCount, maxColumns, columnCount);
            AppendSheetFeatureInventory(body, sheet, GetFeatureInventoryWindow(firstRow, maxRows, rowCount));
            body.Append("</section>");
            return;
        }

        bool firstRowIsHeader = options.HeaderMode == ExcelHtmlHeaderMode.FirstRow;
        body.Append("<table class=\"officeimo-table\">");
        if (firstRowIsHeader) {
            body.Append("<thead>");
        } else {
            body.Append("<tbody>");
        }

        ExcelMergeExportRowCursor mergeCursor = mergeMap.CreateRowCursor();
        for (int row = 0; row < maxRows; row++) {
            if (row == 1 && firstRowIsHeader) {
                body.Append("</thead><tbody>");
            }

            body.Append("<tr>");
            mergeCursor.MoveToRow(firstRow + row);
            for (int column = 0; column < maxColumns; column++) {
                string tag = firstRowIsHeader && row == 0 ? "th" : "td";
                int cellRow = firstRow + row;
                int cellColumn = firstColumn + column;
                if (mergeCursor.IsCoveredCell(cellColumn)) {
                    continue;
                }

                bool hasSnapshot = sheet.TryGetCellValueSnapshot(cellRow, cellColumn, out ExcelCellValueSnapshot? snapshot) && snapshot != null;
                string cellText = ReadCellText(sheet, cellRow, cellColumn, options.EmptyCellText);
                body.Append('<').Append(tag)
                    .Append(" data-officeimo-cell=\"")
                    .Append(OfficeHtmlText.EscapeAttribute(A1.CellReference(cellRow, cellColumn)))
                    .Append('"');
                if (mergeCursor.TryGetOrigin(cellColumn, out ExcelMergeExportRange merge)) {
                    AppendMergeAttributes(body, merge);
                }

                if (tag == "th") {
                    body.Append(" scope=\"col\"");
                }

                if (hasSnapshot) {
                    ExcelCellValueSnapshot cellSnapshot = snapshot!;
                    body.Append(" data-officeimo-value-kind=\"")
                        .Append(OfficeHtmlText.EscapeAttribute(ToHtmlValueKind(cellSnapshot, cellText)))
                        .Append("\" data-officeimo-value=\"")
                        .Append(OfficeHtmlText.EscapeAttribute(ToHtmlRawValue(cellSnapshot, cellText)))
                        .Append('"');
                } else {
                    body.Append(" data-officeimo-empty=\"true\"");
                }

                body.Append('>');
                body.Append(OfficeHtmlText.Escape(cellText));
                body.Append("</").Append(tag).Append('>');
            }

            body.Append("</tr>");
        }

        body.Append(firstRowIsHeader && maxRows == 1 ? "</thead></table>" : "</tbody></table>");
        AppendSheetTruncationDiagnostics(body, maxRows, rowCount, maxColumns, columnCount);

        AppendSheetFeatureInventory(body, sheet, GetFeatureInventoryWindow(firstRow, maxRows, rowCount));
        body.Append("</section>");
    }

    private static void AppendSheetTruncationDiagnostics(
        StringBuilder body,
        int exportedRows,
        int totalRows,
        int exportedColumns,
        int totalColumns) {
        if (exportedColumns < totalColumns) {
            body.Append("<p class=\"officeimo-diagnostic\">Columns truncated: ")
                .Append(exportedColumns.ToString(CultureInfo.InvariantCulture))
                .Append(" of ")
                .Append(totalColumns.ToString(CultureInfo.InvariantCulture))
                .Append(" exported.</p>");
        }
        if (exportedRows < totalRows) {
            body.Append("<p class=\"officeimo-diagnostic\">Rows truncated: ")
                .Append(exportedRows.ToString(CultureInfo.InvariantCulture))
                .Append(" of ")
                .Append(totalRows.ToString(CultureInfo.InvariantCulture))
                .Append(" exported.</p>");
        }
    }

    private static bool SheetHasUsedCells(ExcelSheet sheet, int firstRow, int firstColumn, int rowCount, int columnCount) {
        for (int row = 0; row < rowCount; row++) {
            for (int column = 0; column < columnCount; column++) {
                if (sheet.TryGetCellValueSnapshot(firstRow + row, firstColumn + column, out _)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static void AppendSheetFeatureInventory(StringBuilder body, ExcelSheet sheet, ExportedRowWindow? rowWindow) {
        AppendFormulaInventory(body, FilterFormulas(sheet.GetFormulaCells(), rowWindow));
        AppendCommentInventory(body, FilterComments(sheet.GetComments(), rowWindow));
        AppendChartInventory(body, FilterCharts(sheet.Charts, rowWindow));
        AppendImageInventory(body, FilterImages(sheet.Images, rowWindow));
    }

    private static ExportedRowWindow? GetFeatureInventoryWindow(int firstRow, int exportedRows, int totalRows) =>
        exportedRows < totalRows ? new ExportedRowWindow(firstRow, exportedRows) : null;

    private static IReadOnlyList<ExcelFormulaCellInfo> FilterFormulas(IReadOnlyList<ExcelFormulaCellInfo> formulas, ExportedRowWindow? rowWindow) {
        if (rowWindow == null || formulas.Count == 0) {
            return formulas;
        }

        return formulas.Where(formula => rowWindow.Value.ContainsCellReference(formula.CellReference)).ToList();
    }

    private static IReadOnlyList<ExcelCommentInfo> FilterComments(IReadOnlyList<ExcelCommentInfo> comments, ExportedRowWindow? rowWindow) {
        if (rowWindow == null || comments.Count == 0) {
            return comments;
        }

        return comments.Where(comment => rowWindow.Value.ContainsRow(comment.Row)).ToList();
    }

    private static IEnumerable<ExcelChart> FilterCharts(IEnumerable<ExcelChart> charts, ExportedRowWindow? rowWindow) {
        if (rowWindow == null) {
            return charts;
        }

        return charts.Where(chart => !chart.TryGetSnapshot(out ExcelChartSnapshot snapshot) || rowWindow.Value.ContainsRow(snapshot.RowIndex));
    }

    private static IEnumerable<ExcelImage> FilterImages(IEnumerable<ExcelImage> images, ExportedRowWindow? rowWindow) {
        if (rowWindow == null) {
            return images;
        }

        return images.Where(image => !image.HasAbsoluteAnchor && rowWindow.Value.ContainsRow(image.RowIndex));
    }

    private static void AppendFormulaInventory(StringBuilder body, IReadOnlyList<ExcelFormulaCellInfo> formulas) {
        if (formulas.Count == 0) {
            return;
        }

        body.Append("<section class=\"officeimo-feature officeimo-formulas\"><h3>Formulas</h3><ul class=\"officeimo-feature-list\">");
        foreach (ExcelFormulaCellInfo formula in formulas) {
            body.Append("<li class=\"officeimo-feature-item\" data-officeimo-cell=\"")
                .Append(OfficeHtmlText.EscapeAttribute(formula.CellReference))
                .Append("\"><span class=\"officeimo-feature-label\">")
                .Append(OfficeHtmlText.Escape(formula.CellReference))
                .Append("</span><code>")
                .Append(OfficeHtmlText.Escape(formula.Formula))
                .Append("</code>");
            if (formula.HasCachedValue) {
                body.Append("<div class=\"officeimo-feature-meta\">Cached value: ")
                    .Append(OfficeHtmlText.Escape(formula.CachedValue ?? string.Empty))
                    .Append("</div>");
            }

            if (formula.Dependencies.Count > 0) {
                body.Append("<div class=\"officeimo-feature-meta\">Dependencies: ")
                    .Append(OfficeHtmlText.Escape(string.Join(", ", formula.Dependencies)))
                    .Append("</div>");
            }

            body.Append("</li>");
        }

        body.Append("</ul></section>");
    }

    private static void AppendCommentInventory(StringBuilder body, IReadOnlyList<ExcelCommentInfo> comments) {
        if (comments.Count == 0) {
            return;
        }

        body.Append("<section class=\"officeimo-feature officeimo-comments\"><h3>Comments</h3><ul class=\"officeimo-feature-list\">");
        foreach (ExcelCommentInfo comment in comments) {
            body.Append("<li class=\"officeimo-feature-item\" data-officeimo-cell=\"")
                .Append(OfficeHtmlText.EscapeAttribute(comment.CellReference))
                .Append("\"><span class=\"officeimo-feature-label\">")
                .Append(OfficeHtmlText.Escape(comment.CellReference))
                .Append("</span><p>")
                .Append(OfficeHtmlText.Escape(comment.Text))
                .Append("</p>");
            if (!string.IsNullOrWhiteSpace(comment.Author)) {
                body.Append("<div class=\"officeimo-feature-meta\">Author: ")
                    .Append(OfficeHtmlText.Escape(comment.Author!))
                    .Append("</div>");
            }

            body.Append("</li>");
        }

        body.Append("</ul></section>");
    }

    private static void AppendVisualCommentInventory(StringBuilder body, IReadOnlyList<ExcelCommentInfo> comments) {
        if (comments.Count == 0) {
            return;
        }

        body.Append("<section class=\"officeimo-feature officeimo-comments officeimo-visual-comments\" data-officeimo-visual-proof=\"comment-callout\"><h3>Visible comments</h3><ul class=\"officeimo-feature-list\">");
        foreach (ExcelCommentInfo comment in comments) {
            body.Append("<li class=\"officeimo-feature-item\" data-officeimo-cell=\"")
                .Append(OfficeHtmlText.EscapeAttribute(comment.CellReference))
                .Append("\"><span class=\"officeimo-feature-label\">")
                .Append(OfficeHtmlText.Escape(comment.CellReference))
                .Append("</span><p>")
                .Append(OfficeHtmlText.Escape(comment.Text))
                .Append("</p>");
            if (!string.IsNullOrWhiteSpace(comment.Author)) {
                body.Append("<div class=\"officeimo-feature-meta\">Author: ")
                    .Append(OfficeHtmlText.Escape(comment.Author!))
                    .Append("</div>");
            }

            body.Append("</li>");
        }

        body.Append("</ul></section>");
    }

    private static void AppendChartInventory(StringBuilder body, IEnumerable<ExcelChart> charts) {
        List<ExcelChart> chartList = charts.ToList();
        if (chartList.Count == 0) {
            return;
        }

        body.Append("<section class=\"officeimo-feature officeimo-charts\"><h3>Charts</h3><ul class=\"officeimo-feature-list\">");
        foreach (ExcelChart chart in chartList) {
            string label = string.IsNullOrWhiteSpace(chart.Title)
                ? string.IsNullOrWhiteSpace(chart.Name) ? "Chart" : chart.Name
                : chart.Title!;
            bool hasSnapshot = chart.TryGetSnapshot(out ExcelChartSnapshot snapshot);
            body.Append("<li class=\"officeimo-feature-item\" data-officeimo-layer-kind=\"chart\" data-officeimo-layer-index=\"")
                .Append(chart.DrawingOrder.ToString(CultureInfo.InvariantCulture))
                .Append("\" data-officeimo-chart-type=\"")
                .Append(OfficeHtmlText.EscapeAttribute(chart.ChartType.ToString()))
                .Append('"');
            if (hasSnapshot) {
                AppendDataAttribute(body, "data-officeimo-row", snapshot.RowIndex);
                AppendDataAttribute(body, "data-officeimo-column", snapshot.ColumnIndex);
                AppendDataAttribute(body, "data-officeimo-width", snapshot.WidthPixels);
                AppendDataAttribute(body, "data-officeimo-height", snapshot.HeightPixels);
            }

            body.Append("><span class=\"officeimo-feature-label\">")
                .Append(OfficeHtmlText.Escape(label))
                .Append("</span><div class=\"officeimo-feature-meta\">Type: ")
                .Append(OfficeHtmlText.Escape(chart.ChartType.ToString()))
                .Append("</div>");
            if (hasSnapshot) {
                body.Append("<div class=\"officeimo-feature-meta\">Series: ")
                    .Append(snapshot.Data.Series.Count.ToString(CultureInfo.InvariantCulture))
                    .Append("; Categories: ")
                    .Append(snapshot.Data.Categories.Count.ToString(CultureInfo.InvariantCulture))
                    .Append("; Cell: ")
                    .Append(snapshot.RowIndex.ToString(CultureInfo.InvariantCulture))
                    .Append(", ")
                    .Append(snapshot.ColumnIndex.ToString(CultureInfo.InvariantCulture))
                    .Append("; Size: ")
                    .Append(snapshot.WidthPixels.ToString(CultureInfo.InvariantCulture))
                    .Append("x")
                    .Append(snapshot.HeightPixels.ToString(CultureInfo.InvariantCulture))
                    .Append("</div>");
                AppendChartDataTable(body, snapshot.Data, snapshot.ChartType);
            } else {
                body.Append("<div class=\"officeimo-diagnostic\">Chart data snapshot unavailable; visual review may still render drawing geometry.</div>");
            }

            body.Append("</li>");
        }

        body.Append("</ul></section>");
    }

    private static void AppendChartDataTable(StringBuilder body, ExcelChartData data, ExcelChartType defaultChartType) {
        body.Append("<table class=\"officeimo-chart-data\"><thead><tr><th>Series</th>");
        foreach (string category in data.Categories) {
            body.Append("<th>")
                .Append(OfficeHtmlText.Escape(category))
                .Append("</th>");
        }

        body.Append("</tr></thead><tbody>");
        foreach (ExcelChartSeries series in data.Series) {
            ExcelChartType chartType = series.ChartType ?? defaultChartType;
            body.Append("<tr");
            body.Append(" data-officeimo-chart-type=\"")
                .Append(OfficeHtmlText.EscapeAttribute(chartType.ToString()))
                .Append('"');

            body.Append("><th>")
                .Append(OfficeHtmlText.Escape(series.Name))
                .Append("</th>");
            for (int i = 0; i < series.Values.Count; i++) {
                body.Append("<td");
                if (series.XValues != null && i < series.XValues.Count) {
                    body.Append(" data-officeimo-x=\"")
                        .Append(OfficeHtmlText.EscapeAttribute(series.XValues[i].ToString("G17", CultureInfo.InvariantCulture)))
                        .Append('"');
                }

                body.Append('>')
                    .Append(series.Values[i].ToString("G17", CultureInfo.InvariantCulture))
                    .Append("</td>");
            }

            body.Append("</tr>");
        }

        body.Append("</tbody></table>");
    }

    private static void AppendImageInventory(StringBuilder body, IEnumerable<ExcelImage> images) {
        List<ExcelImage> imageList = images.ToList();
        if (imageList.Count == 0) {
            return;
        }

        body.Append("<section class=\"officeimo-feature officeimo-images\"><h3>Images</h3><ul class=\"officeimo-feature-list\">");
        foreach (ExcelImage image in imageList) {
            string label = !string.IsNullOrWhiteSpace(image.Title)
                ? image.Title
                : !string.IsNullOrWhiteSpace(image.Name) ? image.Name : "Image";
            body.Append("<li class=\"officeimo-feature-item\" data-officeimo-layer-kind=\"image\" data-officeimo-layer-index=\"")
                .Append(image.DrawingOrder.ToString(CultureInfo.InvariantCulture))
                .Append("\" data-officeimo-anchor=\"")
                .Append(image.HasAbsoluteAnchor ? "absolute" : image.HasTwoCellAnchor ? "twoCell" : "oneCell")
                .Append('"');
            AppendDataAttribute(body, "data-officeimo-row", image.RowIndex);
            AppendDataAttribute(body, "data-officeimo-column", image.ColumnIndex);
            AppendDataAttribute(body, "data-officeimo-width", image.WidthPixels);
            AppendDataAttribute(body, "data-officeimo-height", image.HeightPixels);
            AppendDataAttribute(body, "data-officeimo-offset-x", image.OffsetXPixels);
            AppendDataAttribute(body, "data-officeimo-offset-y", image.OffsetYPixels);
            AppendDataAttribute(body, "data-officeimo-rotation", image.RotationDegrees);
            AppendDataAttribute(body, "data-officeimo-flip-horizontal", image.FlipHorizontal);
            AppendDataAttribute(body, "data-officeimo-flip-vertical", image.FlipVertical);
            AppendDataAttribute(body, "data-officeimo-crop-left", image.CropLeftRatio);
            AppendDataAttribute(body, "data-officeimo-crop-top", image.CropTopRatio);
            AppendDataAttribute(body, "data-officeimo-crop-right", image.CropRightRatio);
            AppendDataAttribute(body, "data-officeimo-crop-bottom", image.CropBottomRatio);
            if (image.HasAbsoluteAnchor && image.TryGetAbsoluteAnchorBounds(out int xPixels, out int yPixels, out _, out _)) {
                AppendDataAttribute(body, "data-officeimo-x", xPixels);
                AppendDataAttribute(body, "data-officeimo-y", yPixels);
            }

            if (image.HasTwoCellAnchor && image.ToRowIndex.HasValue && image.ToColumnIndex.HasValue) {
                AppendDataAttribute(body, "data-officeimo-to-row", image.ToRowIndex.Value);
                AppendDataAttribute(body, "data-officeimo-to-column", image.ToColumnIndex.Value);
                AppendDataAttribute(body, "data-officeimo-to-offset-x", image.ToOffsetXPixels);
                AppendDataAttribute(body, "data-officeimo-to-offset-y", image.ToOffsetYPixels);
            }

            body.Append("><span class=\"officeimo-feature-label\">")
                .Append(OfficeHtmlText.Escape(label))
                .Append("</span><div class=\"officeimo-feature-meta\">Cell: ")
                .Append(image.RowIndex.ToString(CultureInfo.InvariantCulture))
                .Append(", ")
                .Append(image.ColumnIndex.ToString(CultureInfo.InvariantCulture))
                .Append("; Size: ")
                .Append(image.WidthPixels.ToString(CultureInfo.InvariantCulture))
                .Append("x")
                .Append(image.HeightPixels.ToString(CultureInfo.InvariantCulture))
                .Append("; Offset: ")
                .Append(image.OffsetXPixels.ToString(CultureInfo.InvariantCulture))
                .Append(", ")
                .Append(image.OffsetYPixels.ToString(CultureInfo.InvariantCulture))
                .Append("; Type: ")
                .Append(OfficeHtmlText.Escape(image.ContentType))
                .Append("</div>");
            if (!string.IsNullOrWhiteSpace(image.Description)) {
                body.Append("<p>")
                    .Append(OfficeHtmlText.Escape(image.Description))
                    .Append("</p>");
            }

            AppendImagePreview(body, image.ToBytes(), image.ContentType, label);
            body.Append("</li>");
        }

        body.Append("</ul></section>");
    }

    private static void AppendDataAttribute(StringBuilder body, string name, int value) {
        body.Append(' ')
            .Append(name)
            .Append("=\"")
            .Append(value.ToString(CultureInfo.InvariantCulture))
            .Append('"');
    }

    private static void AppendDataAttribute(StringBuilder body, string name, double value) {
        if (Math.Abs(value) < 0.0000001D) {
            return;
        }

        body.Append(' ')
            .Append(name)
            .Append("=\"")
            .Append(value.ToString("G17", CultureInfo.InvariantCulture))
            .Append('"');
    }

    private static void AppendDataAttribute(StringBuilder body, string name, bool value) {
        if (!value) {
            return;
        }

        body.Append(' ')
            .Append(name)
            .Append("=\"true\"");
    }

    private static void AppendImagePreview(StringBuilder body, byte[] bytes, string contentType, string label) {
        if (bytes.Length == 0 || string.IsNullOrWhiteSpace(contentType)) {
            return;
        }

        body.Append("<img class=\"officeimo-inline-image\" alt=\"")
            .Append(OfficeHtmlText.EscapeAttribute(label))
            .Append("\" src=\"data:")
            .Append(OfficeHtmlText.EscapeAttribute(contentType))
            .Append(";base64,")
            .Append(Convert.ToBase64String(bytes))
            .Append("\">");
    }

    private static void AppendSvgResult(StringBuilder body, OfficeImageExportResult result, string idPrefix) {
        string svg = NamespaceSvgIds(Encoding.UTF8.GetString(result.Bytes), idPrefix);
        body.Append("<section class=\"officeimo-sheet\" data-officeimo-source-anchor=\"")
            .Append(OfficeHtmlText.EscapeAttribute(result.Source))
            .Append("\">");
        body.Append("<h2>").Append(OfficeHtmlText.Escape(string.IsNullOrWhiteSpace(result.Name) ? "Worksheet" : result.Name)).Append("</h2>");
        body.Append("<div class=\"officeimo-visual-page\" data-officeimo-visual-owner=\"OfficeIMO.Drawing\">");
        body.Append(svg);
        body.Append("</div>");
        foreach (IGrouping<string, OfficeImageExportDiagnostic> group in result.Diagnostics.GroupBy(CreateDiagnosticGroupKey)) {
            OfficeImageExportDiagnostic diagnostic = group.First();
            body.Append("<p class=\"officeimo-diagnostic\">")
                .Append(OfficeHtmlText.Escape(diagnostic.Message));
            int count = group.Count();
            if (count > 1) {
                body.Append(" (")
                    .Append(count.ToString(CultureInfo.InvariantCulture))
                    .Append(" occurrences)");
            }

            body.Append("</p>");
        }

        body.Append("</section>");
    }

    private static string CreateSvgNamespacePrefix(OfficeImageExportResult result, int index) {
        string name = !string.IsNullOrWhiteSpace(result.Name)
            ? result.Name!
            : (string.IsNullOrWhiteSpace(result.Source) ? "worksheet" : result.Source!);
        var builder = new StringBuilder("officeimo-sheet-svg-");
        foreach (char value in name) {
            if (char.IsLetterOrDigit(value)) {
                builder.Append(char.ToLowerInvariant(value));
            } else if (builder[builder.Length - 1] != '-') {
                builder.Append('-');
            }
        }

        if (builder[builder.Length - 1] != '-') {
            builder.Append('-');
        }

        builder.Append(index.ToString(CultureInfo.InvariantCulture)).Append('-');
        return builder.ToString();
    }

    private static string NamespaceSvgIds(string svg, string prefix) {
        if (string.IsNullOrEmpty(svg) || string.IsNullOrEmpty(prefix)) {
            return svg;
        }

        var ids = new SortedSet<string>(StringComparer.Ordinal);
        foreach (System.Text.RegularExpressions.Match match in System.Text.RegularExpressions.Regex.Matches(svg, "\\bid=(['\\\"])(?<id>[^'\\\"]+)\\1")) {
            string id = match.Groups["id"].Value;
            if (id.Length > 0 && !id.StartsWith(prefix, StringComparison.Ordinal)) {
                ids.Add(id);
            }
        }

        string namespaced = svg;
        foreach (string id in ids.OrderByDescending(item => item.Length)) {
            string replacement = prefix + id;
            namespaced = namespaced
                .Replace("id=\"" + id + "\"", "id=\"" + replacement + "\"")
                .Replace("id='" + id + "'", "id='" + replacement + "'")
                .Replace("url(#" + id + ")", "url(#" + replacement + ")")
                .Replace("href=\"#" + id + "\"", "href=\"#" + replacement + "\"")
                .Replace("href='#" + id + "'", "href='#" + replacement + "'")
                .Replace("xlink:href=\"#" + id + "\"", "xlink:href=\"#" + replacement + "\"")
                .Replace("xlink:href='#" + id + "'", "xlink:href='#" + replacement + "'");
        }

        return namespaced;
    }

    private static string CreateDiagnosticGroupKey(OfficeImageExportDiagnostic diagnostic) {
        return diagnostic.Severity.ToString() + "|" + diagnostic.Code + "|" + diagnostic.Message;
    }

    private static ExcelWorkbookImageExportOptions ResolveWorkbookVisualOptions(ExcelWorkbookImageExportOptions? workbookOptions) {
        if (workbookOptions == null) {
            return new ExcelWorkbookImageExportOptions {
                ShowCommentBodies = true
            };
        }

        return workbookOptions;
    }

    private static ExcelWorksheetImageExportOptions ToWorksheetOptions(ExcelWorkbookImageExportOptions workbookOptions) {
        return new ExcelWorksheetImageExportOptions {
            Scale = workbookOptions.Scale,
            BackgroundColor = workbookOptions.BackgroundColor,
            GridlineColor = workbookOptions.GridlineColor,
            ShowGridlines = workbookOptions.ShowGridlines,
            IncludeHidden = workbookOptions.IncludeHidden,
            IncludeImages = workbookOptions.IncludeImages,
            IncludeCharts = workbookOptions.IncludeCharts,
            IncludeDrawingObjects = workbookOptions.IncludeDrawingObjects,
            IncludeConditionalFormatting = workbookOptions.IncludeConditionalFormatting,
            ConditionalFormattingDate = workbookOptions.ConditionalFormattingDate,
            ShowHyperlinkHints = workbookOptions.ShowHyperlinkHints,
            ShowCommentBodies = workbookOptions.ShowCommentBodies,
            DefaultColumnWidthPixels = workbookOptions.DefaultColumnWidthPixels,
            DefaultRowHeightPixels = workbookOptions.DefaultRowHeightPixels,
            HeaderFooterDateTime = workbookOptions.HeaderFooterDateTime,
            UsePrintArea = workbookOptions.UseWorksheetPrintAreas,
            SplitByManualPageBreaks = workbookOptions.SplitWorksheetsByManualPageBreaks
        };
    }

    private static string Wrap(string body, ExcelHtmlSaveOptions options, string title) {
        return OfficeHtmlDocumentShell.WrapBody(body, new OfficeHtmlDocumentOptions {
            Title = title,
            Theme = options.Theme,
            IncludeDefaultStyles = options.IncludeDefaultStyles,
            BodyClass = "officeimo-html officeimo-excel-html"
        });
    }

    private static string GetTitle(ExcelHtmlSaveOptions options, string fallback) {
        return string.IsNullOrWhiteSpace(options.Title) ? fallback : options.Title!;
    }

    private static void AppendSheetVisibilityAttribute(StringBuilder body, ExcelSheet sheet) {
        if (sheet.VeryHidden) {
            body.Append(" data-officeimo-visibility=\"veryHidden\"");
        } else if (sheet.Hidden) {
            body.Append(" data-officeimo-visibility=\"hidden\"");
        }
    }

    private static string ReadCellText(ExcelSheet sheet, int row, int column, string emptyCellText) {
        if (!sheet.TryGetCellText(row, column, out string text)) {
            return emptyCellText;
        }

        return string.IsNullOrEmpty(text) ? emptyCellText : text;
    }

    private static string ToHtmlValueKind(ExcelCellValueKind kind) =>
        kind switch {
            ExcelCellValueKind.Number => "number",
            ExcelCellValueKind.Boolean => "boolean",
            ExcelCellValueKind.Error => "error",
            ExcelCellValueKind.Formula => "formula",
            ExcelCellValueKind.DateTime => "date-time",
            ExcelCellValueKind.Other => "other",
            _ => "text"
        };

    private static string ToHtmlValueKind(ExcelCellValueSnapshot snapshot, string cellText) =>
        ShouldUseDisplayTextRawValue(snapshot, cellText) ? "text" : ToHtmlValueKind(snapshot.Kind);

    private static string ToHtmlRawValue(ExcelCellValueSnapshot snapshot, string cellText) =>
        snapshot.Kind == ExcelCellValueKind.DateTime && snapshot.DateTimeValue.HasValue
            ? snapshot.DateTimeValue.Value.ToString("O", CultureInfo.InvariantCulture)
            : ShouldUseDisplayTextRawValue(snapshot, cellText) ? cellText : snapshot.RawValue;

    private static bool ShouldUseDisplayTextRawValue(ExcelCellValueSnapshot snapshot, string cellText) =>
        snapshot.Kind == ExcelCellValueKind.Text ||
        (HasSignificantWhitespace(cellText) &&
            !string.Equals(cellText, snapshot.RawValue, StringComparison.Ordinal) &&
            int.TryParse(snapshot.RawValue, NumberStyles.None, CultureInfo.InvariantCulture, out _));

    private static bool HasSignificantWhitespace(string text) =>
        text.IndexOf('\n') >= 0 ||
        text.IndexOf('\r') >= 0 ||
        text.IndexOf('\t') >= 0 ||
        text.Contains("  ");

    private static void ParseUsedRange(string usedRange, out int firstRow, out int firstColumn, out int rowCount, out int columnCount) {
        if (string.IsNullOrWhiteSpace(usedRange)) {
            firstRow = 1;
            firstColumn = 1;
            rowCount = 0;
            columnCount = 0;
            return;
        }

        string[] parts = usedRange.Split(':');
        ParseCellReference(parts[0], out firstRow, out firstColumn);
        if (parts.Length > 1) {
            ParseCellReference(parts[1], out int lastRow, out int lastColumn);
            rowCount = Math.Max(0, lastRow - firstRow + 1);
            columnCount = Math.Max(0, lastColumn - firstColumn + 1);
        } else {
            rowCount = 1;
            columnCount = 1;
        }
    }

    private static void ParseCellReference(string reference, out int row, out int column) {
        row = 0;
        column = 0;
        foreach (char ch in reference ?? string.Empty) {
            if (ch >= 'A' && ch <= 'Z') {
                column = column * 26 + (ch - 'A' + 1);
            } else if (ch >= 'a' && ch <= 'z') {
                column = column * 26 + (ch - 'a' + 1);
            } else if (ch >= '0' && ch <= '9') {
                row = row * 10 + (ch - '0');
            }
        }

        if (row <= 0) {
            row = 1;
        }

        if (column <= 0) {
            column = 1;
        }
    }

    private readonly struct ExportedRowWindow {
        private readonly int _firstRow;
        private readonly int _lastRow;

        public ExportedRowWindow(int firstRow, int exportedRows) {
            _firstRow = firstRow;
            _lastRow = exportedRows <= 0 ? firstRow - 1 : firstRow + exportedRows - 1;
        }

        public bool ContainsCellReference(string cellReference) {
            ParseCellReference(cellReference, out int row, out _);
            return ContainsRow(row);
        }

        public bool ContainsRow(int row) => row >= _firstRow && row <= _lastRow;
    }
}
