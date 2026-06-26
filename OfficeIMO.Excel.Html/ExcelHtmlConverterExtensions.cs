using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

/// <summary>
/// Extension methods enabling HTML conversions for OfficeIMO Excel documents.
/// </summary>
public static class ExcelHtmlConverterExtensions {
    /// <summary>
    /// Converts a workbook to HTML.
    /// </summary>
    public static string ToHtml(this ExcelDocument workbook, ExcelHtmlSaveOptions? options = null) {
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));
        options ??= new ExcelHtmlSaveOptions();
        return options.Profile == OfficeHtmlConversionProfile.ExcelVisualReview
            ? ConvertWorkbookVisual(workbook, options)
            : ConvertWorkbookSemantic(workbook, options);
    }

    /// <summary>
    /// Converts a worksheet to HTML.
    /// </summary>
    public static string ToHtml(this ExcelSheet sheet, ExcelHtmlSaveOptions? options = null) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        options ??= new ExcelHtmlSaveOptions();
        return options.Profile == OfficeHtmlConversionProfile.ExcelVisualReview
            ? ConvertSheetVisual(sheet, options)
            : ConvertSheetSemantic(sheet, options);
    }

    /// <summary>
    /// Saves a workbook as HTML.
    /// </summary>
    public static void SaveAsHtml(this ExcelDocument workbook, string path, ExcelHtmlSaveOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("HTML path cannot be empty.", nameof(path));
        File.WriteAllText(path, workbook.ToHtml(options), Encoding.UTF8);
    }

    /// <summary>
    /// Saves a worksheet as HTML.
    /// </summary>
    public static void SaveAsHtml(this ExcelSheet sheet, string path, ExcelHtmlSaveOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("HTML path cannot be empty.", nameof(path));
        File.WriteAllText(path, sheet.ToHtml(options), Encoding.UTF8);
    }

    private static string ConvertWorkbookSemantic(ExcelDocument workbook, ExcelHtmlSaveOptions options) {
        var body = new StringBuilder();
        body.Append("<main class=\"officeimo-document\" data-officeimo-source=\"excel\" data-officeimo-profile=\"")
            .Append(OfficeHtmlText.EscapeAttribute(options.Profile.ToString()))
            .Append("\">");
        body.Append("<h1>").Append(OfficeHtmlText.Escape(GetTitle(options, "Excel Workbook"))).Append("</h1>");
        foreach (ExcelSheet sheet in workbook.Sheets) {
            AppendSheetTable(body, sheet, options);
        }

        body.Append("</main>");
        return Wrap(body.ToString(), options, GetTitle(options, "Excel Workbook"));
    }

    private static string ConvertSheetSemantic(ExcelSheet sheet, ExcelHtmlSaveOptions options) {
        var body = new StringBuilder();
        body.Append("<main class=\"officeimo-document\" data-officeimo-source=\"excel\" data-officeimo-profile=\"")
            .Append(OfficeHtmlText.EscapeAttribute(options.Profile.ToString()))
            .Append("\">");
        AppendSheetTable(body, sheet, options);
        body.Append("</main>");
        return Wrap(body.ToString(), options, GetTitle(options, sheet.Name));
    }

    private static string ConvertWorkbookVisual(ExcelDocument workbook, ExcelHtmlSaveOptions options) {
        var body = new StringBuilder();
        body.Append("<main class=\"officeimo-document\" data-officeimo-source=\"excel\" data-officeimo-profile=\"")
            .Append(OfficeHtmlText.EscapeAttribute(options.Profile.ToString()))
            .Append("\">");
        body.Append("<h1>").Append(OfficeHtmlText.Escape(GetTitle(options, "Excel Visual Review"))).Append("</h1>");
        ExcelWorkbookImageExportOptions visualOptions = ResolveWorkbookVisualOptions(options.VisualOptions);
        Dictionary<string, ExcelSheet> sheetsByName = workbook.Sheets.ToDictionary(sheet => sheet.Name, StringComparer.OrdinalIgnoreCase);
        foreach (OfficeImageExportResult result in workbook.ExportImages(OfficeImageExportFormat.Svg, visualOptions)) {
            AppendSvgResult(body, result);
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
        body.Append("<main class=\"officeimo-document\" data-officeimo-source=\"excel\" data-officeimo-profile=\"")
            .Append(OfficeHtmlText.EscapeAttribute(options.Profile.ToString()))
            .Append("\">");
        body.Append("<h1>").Append(OfficeHtmlText.Escape(GetTitle(options, sheet.Name))).Append("</h1>");
        AppendSvgResult(body, sheet.ExportImage(OfficeImageExportFormat.Svg, ToWorksheetOptions(ResolveWorkbookVisualOptions(options.VisualOptions))));
        AppendVisualCommentInventory(body, sheet.GetComments());
        body.Append("</main>");
        return Wrap(body.ToString(), options, GetTitle(options, sheet.Name));
    }

    private static void AppendSheetTable(StringBuilder body, ExcelSheet sheet, ExcelHtmlSaveOptions options) {
        string usedRange = sheet.GetUsedRangeA1();
        body.Append("<section class=\"officeimo-sheet\" data-officeimo-sheet=\"")
            .Append(OfficeHtmlText.EscapeAttribute(sheet.Name))
            .Append("\" data-officeimo-range=\"")
            .Append(OfficeHtmlText.EscapeAttribute(usedRange))
            .Append("\">");
        body.Append("<h2>").Append(OfficeHtmlText.Escape(sheet.Name)).Append("</h2>");

        ParseUsedRange(usedRange, out int firstRow, out int firstColumn, out int rowCount, out int columnCount);
        int maxRows = options.MaxRowsPerSheet.HasValue
            ? Math.Min(rowCount, options.MaxRowsPerSheet.Value)
            : rowCount;
        if (rowCount == 0 || columnCount == 0 || maxRows == 0) {
            body.Append("<p class=\"officeimo-muted\">No used cells.</p>");
            AppendSheetFeatureInventory(body, sheet);
            body.Append("</section>");
            return;
        }

        body.Append("<table class=\"officeimo-table\"><tbody>");
        for (int row = 0; row < maxRows; row++) {
            body.Append("<tr>");
            for (int column = 0; column < columnCount; column++) {
                string tag = row == 0 ? "th" : "td";
                body.Append('<').Append(tag).Append('>');
                body.Append(OfficeHtmlText.Escape(ReadCellText(sheet, firstRow + row, firstColumn + column, options.EmptyCellText)));
                body.Append("</").Append(tag).Append('>');
            }

            body.Append("</tr>");
        }

        body.Append("</tbody></table>");
        if (maxRows < rowCount) {
            body.Append("<p class=\"officeimo-diagnostic\">Rows truncated: ")
                .Append(maxRows.ToString(CultureInfo.InvariantCulture))
                .Append(" of ")
                .Append(rowCount.ToString(CultureInfo.InvariantCulture))
                .Append(" exported.</p>");
        }

        AppendSheetFeatureInventory(body, sheet);
        body.Append("</section>");
    }

    private static void AppendSheetFeatureInventory(StringBuilder body, ExcelSheet sheet) {
        AppendFormulaInventory(body, sheet.GetFormulaCells());
        AppendCommentInventory(body, sheet.GetComments());
        AppendChartInventory(body, sheet.Charts);
        AppendImageInventory(body, sheet.Images);
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
            body.Append("<li class=\"officeimo-feature-item\"><span class=\"officeimo-feature-label\">")
                .Append(OfficeHtmlText.Escape(label))
                .Append("</span><div class=\"officeimo-feature-meta\">Type: ")
                .Append(OfficeHtmlText.Escape(chart.ChartType.ToString()))
                .Append("</div>");
            if (chart.TryGetSnapshot(out ExcelChartSnapshot snapshot)) {
                body.Append("<div class=\"officeimo-feature-meta\">Series: ")
                    .Append(snapshot.Data.Series.Count.ToString(CultureInfo.InvariantCulture))
                    .Append("; Categories: ")
                    .Append(snapshot.Data.Categories.Count.ToString(CultureInfo.InvariantCulture))
                    .Append("</div>");
                AppendChartDataTable(body, snapshot.Data);
            } else {
                body.Append("<div class=\"officeimo-diagnostic\">Chart data snapshot unavailable; visual review may still render drawing geometry.</div>");
            }

            body.Append("</li>");
        }

        body.Append("</ul></section>");
    }

    private static void AppendChartDataTable(StringBuilder body, ExcelChartData data) {
        body.Append("<table class=\"officeimo-chart-data\"><thead><tr><th>Series</th>");
        foreach (string category in data.Categories) {
            body.Append("<th>")
                .Append(OfficeHtmlText.Escape(category))
                .Append("</th>");
        }

        body.Append("</tr></thead><tbody>");
        foreach (ExcelChartSeries series in data.Series) {
            body.Append("<tr><th>")
                .Append(OfficeHtmlText.Escape(series.Name))
                .Append("</th>");
            foreach (double value in series.Values) {
                body.Append("<td>")
                    .Append(value.ToString("G17", CultureInfo.InvariantCulture))
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
            body.Append("<li class=\"officeimo-feature-item\"><span class=\"officeimo-feature-label\">")
                .Append(OfficeHtmlText.Escape(label))
                .Append("</span><div class=\"officeimo-feature-meta\">Cell: ")
                .Append(image.RowIndex.ToString(CultureInfo.InvariantCulture))
                .Append(", ")
                .Append(image.ColumnIndex.ToString(CultureInfo.InvariantCulture))
                .Append("; Size: ")
                .Append(image.WidthPixels.ToString(CultureInfo.InvariantCulture))
                .Append("x")
                .Append(image.HeightPixels.ToString(CultureInfo.InvariantCulture))
                .Append("; Type: ")
                .Append(OfficeHtmlText.Escape(image.ContentType))
                .Append("</div>");
            if (!string.IsNullOrWhiteSpace(image.Description)) {
                body.Append("<p>")
                    .Append(OfficeHtmlText.Escape(image.Description))
                    .Append("</p>");
            }

            AppendImagePreview(body, image.GetBytes(), image.ContentType, label);
            body.Append("</li>");
        }

        body.Append("</ul></section>");
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

    private static void AppendSvgResult(StringBuilder body, OfficeImageExportResult result) {
        string svg = Encoding.UTF8.GetString(result.Bytes);
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

    private static string ReadCellText(ExcelSheet sheet, int row, int column, string emptyCellText) {
        if (!sheet.TryGetCellText(row, column, out string text)) {
            return emptyCellText;
        }

        return string.IsNullOrEmpty(text) ? emptyCellText : text;
    }

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
}
