using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>
/// Extension methods for importing semantic OfficeIMO PowerPoint HTML.
/// </summary>
public static class PowerPointHtmlLoadExtensions {
    /// <summary>
    /// Imports semantic OfficeIMO PowerPoint HTML into a native presentation.
    /// </summary>
    public static PptCore.PowerPointPresentation LoadPowerPointFromHtml(this string html, PowerPointHtmlLoadOptions? options = null) =>
        LoadPowerPointFromHtmlWithResult(html, options).Presentation;

    /// <summary>
    /// Imports semantic OfficeIMO PowerPoint HTML into a native presentation and returns import evidence.
    /// </summary>
    public static PowerPointHtmlLoadResult LoadPowerPointFromHtmlWithResult(this string html, PowerPointHtmlLoadOptions? options = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        options ??= new PowerPointHtmlLoadOptions();

        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
        var stream = new MemoryStream();
        PptCore.PowerPointPresentation presentation = PptCore.PowerPointPresentation.Create(stream);
        var result = new PowerPointHtmlLoadResult(presentation);

        List<IElement> slideSections = document.QuerySelectorAll("section.officeimo-slide").ToList();
        if (slideSections.Count == 0) {
            result.Diagnostics.Add("No semantic PowerPoint slide sections were found.");
            return result;
        }

        foreach (IElement slideSection in slideSections) {
            PptCore.PowerPointSlide slide = presentation.AddSlide();
            ImportSlide(slideSection, slide, options, result);
        }

        return result;
    }

    private static void ImportSlide(IElement section, PptCore.PowerPointSlide slide, PowerPointHtmlLoadOptions options, PowerPointHtmlLoadResult result) {
        result.Slides++;
        double top = 48D;
        foreach (IElement paragraph in section.Children.Where(child => IsElement(child, "p"))) {
            string text = PreserveText(paragraph.TextContent);
            if (text.Length == 0) {
                continue;
            }

            slide.AddTextBoxPoints(text, 64, top, 620, 48);
            result.TextBoxes++;
            top += 58D;
        }

        if (options.ImportTables) {
            foreach (IElement table in section.Children.Where(child => IsElement(child, "table"))) {
                top = ImportTable(table, slide, top, result);
            }
        }

        if (options.ImportPictures) {
            ImportPictures(section, slide, result);
        }

        if (options.ImportChartInventory) {
            ImportCharts(section, slide, result);
        }

        if (options.ImportNotes) {
            ImportNotes(section, slide, result);
        }
    }

    private static double ImportTable(IElement tableElement, PptCore.PowerPointSlide slide, double top, PowerPointHtmlLoadResult result) {
        List<IElement> rows = tableElement.QuerySelectorAll("tr").ToList();
        int rowCount = rows.Count;
        int columnCount = rows.Count == 0
            ? 0
            : rows.Max(row => row.Children.Count(child => IsElement(child, "th") || IsElement(child, "td")));
        if (rowCount == 0 || columnCount == 0) {
            return top;
        }

        PptCore.PowerPointTable table = slide.AddTablePoints(rowCount, columnCount, 64, top, Math.Max(240, columnCount * 150), Math.Max(70, rowCount * 34));
        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            IElement row = rows[rowIndex];
            List<IElement> cells = row.Children.Where(child => IsElement(child, "th") || IsElement(child, "td")).ToList();
            for (int columnIndex = 0; columnIndex < cells.Count; columnIndex++) {
                table.GetCell(rowIndex, columnIndex).Text = PreserveText(cells[columnIndex].TextContent);
            }
        }

        result.Tables++;
        return top + Math.Max(90, rowCount * 40);
    }

    private static void ImportPictures(IElement section, PptCore.PowerPointSlide slide, PowerPointHtmlLoadResult result) {
        double top = 140D;
        foreach (IElement item in section.QuerySelectorAll("section.officeimo-images li")) {
            IElement? image = item.QuerySelector("img[src]");
            if (image == null || !HtmlImageDataUri.TryParse(image.GetAttribute("src"), out HtmlImageDataUri dataUri)) {
                continue;
            }

            if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
                result.Diagnostics.Add("Picture inventory item '" + NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent) + "' could not be decoded.");
                continue;
            }

            if (!TryGetImagePartType(dataUri.MediaType, out PptCore.ImagePartType imagePartType)) {
                result.Diagnostics.Add("Picture inventory item '" + NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent) + "' used unsupported media type '" + dataUri.MediaType + "' and was not imported.");
                continue;
            }

            ReadPictureSize(item, out double width, out double height);
            using var stream = new MemoryStream(bytes);
            PptCore.PowerPointPicture picture = slide.AddPicturePoints(stream, imagePartType, 720, top, width, height);
            string label = NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent);
            string alt = NormalizeText(image.GetAttribute("alt"));
            if (label.Length > 0) {
                picture.Name = label;
            }

            if (alt.Length > 0) {
                picture.AltText = alt;
            }

            result.Pictures++;
            top += height + 18D;
        }
    }

    private static void ImportCharts(IElement section, PptCore.PowerPointSlide slide, PowerPointHtmlLoadResult result) {
        double top = 220D;
        foreach (IElement item in section.QuerySelectorAll("section.officeimo-charts li")) {
            string title = NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent);
            bool restoredFromSemanticData = TryReadChartData(item, out PptCore.PowerPointChartData? semanticData);
            PptCore.PowerPointChartData data = semanticData ?? CreatePlaceholderChartDataFromInventory(item);
            string chartKind = ReadChartKind(item);
            if (!TryAddChartByKind(slide, chartKind, data, 500, top, 320, 180, out PptCore.PowerPointChart? chart) || chart == null) {
                result.Diagnostics.Add("Chart inventory item '" + (title.Length == 0 ? "Imported chart" : title) + "' used unsupported chart kind '" + chartKind + "' and was not imported.");
                continue;
            }

            chart.SetTitle(title.Length == 0 ? "Imported chart" : title);
            result.Charts++;
            if (!restoredFromSemanticData) {
                result.Diagnostics.Add("Chart inventory item '" + (title.Length == 0 ? "Imported chart" : title) + "' was restored as a native chart with reconstructed placeholder values; exact source chart data was not present in semantic HTML.");
            }

            top += 198D;
        }
    }

    private static void ImportNotes(IElement section, PptCore.PowerPointSlide slide, PowerPointHtmlLoadResult result) {
        string notes = ExtractPresenterNotes(section.QuerySelector("pre.officeimo-source-markdown")?.TextContent);
        if (notes.Length == 0) {
            return;
        }

        slide.Notes.Text = notes;
        result.Notes++;
    }

    private static PptCore.PowerPointChartData CreatePlaceholderChartDataFromInventory(IElement item) {
        ReadChartShape(item, out int seriesCount, out int categoryCount);
        return CreatePlaceholderChartData(seriesCount, categoryCount);
    }

    private static PptCore.PowerPointChartData CreatePlaceholderChartData(int seriesCount, int categoryCount) {
        seriesCount = Math.Max(1, seriesCount);
        categoryCount = Math.Max(1, categoryCount);
        string[] categories = Enumerable.Range(1, categoryCount)
            .Select(index => "C" + index.ToString(CultureInfo.InvariantCulture))
            .ToArray();
        PptCore.PowerPointChartSeries[] series = Enumerable.Range(1, seriesCount)
            .Select(seriesIndex => new PptCore.PowerPointChartSeries(
                "Series " + seriesIndex.ToString(CultureInfo.InvariantCulture),
                Enumerable.Range(1, categoryCount).Select(categoryIndex => (double)(seriesIndex * categoryIndex)).ToArray()))
            .ToArray();
        return new PptCore.PowerPointChartData(categories, series);
    }

    private static bool TryReadChartData(IElement item, out PptCore.PowerPointChartData? data) {
        data = null;
        IElement? table = item.QuerySelector("table.officeimo-chart-data");
        if (table == null) {
            return false;
        }

        List<string> categories = table.QuerySelectorAll("thead tr th")
            .Skip(1)
            .Select(header => PreserveText(header.TextContent))
            .ToList();
        if (categories.Count == 0) {
            return false;
        }

        var series = new List<PptCore.PowerPointChartSeries>();
        foreach (IElement row in table.QuerySelectorAll("tbody tr")) {
            string name = PreserveText(row.QuerySelector("th")?.TextContent);
            if (name.Length == 0) {
                name = "Series " + (series.Count + 1).ToString(CultureInfo.InvariantCulture);
            }

            List<IElement> valueCells = row.QuerySelectorAll("td").ToList();
            var values = new double[valueCells.Count];
            var xValues = new double[valueCells.Count];
            bool hasXValues = valueCells.Any(cell => cell.GetAttribute("data-officeimo-x") != null);
            for (int i = 0; i < valueCells.Count; i++) {
                IElement cell = valueCells[i];
                string text = PreserveText(cell.TextContent);
                if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out values[i])) {
                    return false;
                }

                if (hasXValues) {
                    string? rawXValue = cell.GetAttribute("data-officeimo-x");
                    if (rawXValue == null || !double.TryParse(rawXValue, NumberStyles.Float, CultureInfo.InvariantCulture, out xValues[i])) {
                        return false;
                    }
                }
            }

            if (values.Length != categories.Count) {
                return false;
            }

            series.Add(hasXValues
                ? new PptCore.PowerPointChartSeries(name, values, xValues)
                : new PptCore.PowerPointChartSeries(name, values));
        }

        if (series.Count == 0) {
            return false;
        }

        data = new PptCore.PowerPointChartData(categories, series);
        return true;
    }

    private static void ReadChartShape(IElement item, out int seriesCount, out int categoryCount) {
        seriesCount = 1;
        categoryCount = 3;
        string meta = string.Join(" ", item.QuerySelectorAll(".officeimo-feature-meta").Select(element => element.TextContent));
        TryReadMetaInt(meta, "Series:", ref seriesCount);
        TryReadMetaInt(meta, "Categories:", ref categoryCount);
    }

    private static void TryReadMetaInt(string meta, string marker, ref int value) {
        int index = meta.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (index < 0) {
            return;
        }

        string text = meta.Substring(index + marker.Length).Split(';')[0].Trim();
        if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed > 0) {
            value = parsed;
        }
    }

    private static void ReadPictureSize(IElement item, out double width, out double height) {
        width = 72D;
        height = 72D;
        string meta = string.Join("; ", item.QuerySelectorAll(".officeimo-feature-meta").Select(element => element.TextContent));
        const string marker = "Size:";
        int index = meta.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (index < 0) {
            return;
        }

        string text = meta.Substring(index + marker.Length).Split(';')[0].Trim();
        string[] parts = text.Replace("pt", string.Empty).Split('x');
        if (parts.Length == 2) {
            _ = double.TryParse(parts[0].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out width);
            _ = double.TryParse(parts[1].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out height);
        }

        width = Math.Max(1D, width);
        height = Math.Max(1D, height);
    }

    private static string ExtractPresenterNotes(string? markdown) {
        if (string.IsNullOrWhiteSpace(markdown)) {
            return string.Empty;
        }

        string normalized = markdown!.Replace("\r\n", "\n").Replace('\r', '\n');
        int marker = normalized.IndexOf("### Notes", StringComparison.OrdinalIgnoreCase);
        if (marker < 0) {
            return string.Empty;
        }

        string tail = normalized.Substring(marker + "### Notes".Length);
        string[] lines = tail.Split('\n')
            .Select(line => line.Trim())
            .Where(line => line.Length > 0)
            .ToArray();
        return string.Join(Environment.NewLine, lines);
    }

    private static bool TryAddChartByKind(PptCore.PowerPointSlide slide, string chartKind, PptCore.PowerPointChartData data, double left, double top, double width, double height, out PptCore.PowerPointChart? chart) {
        chart = null;
        if (chartKind.Equals("ClusteredColumn", StringComparison.OrdinalIgnoreCase) ||
            chartKind.Equals("ColumnClustered", StringComparison.OrdinalIgnoreCase)) {
            chart = slide.AddChartPoints(data, left, top, width, height);
            return true;
        }

        if (chartKind.Equals("Line", StringComparison.OrdinalIgnoreCase) ||
            chartKind.Equals("StackedLine", StringComparison.OrdinalIgnoreCase) ||
            chartKind.Equals("StackedLine100", StringComparison.OrdinalIgnoreCase)) {
            chart = slide.AddLineChartPoints(data, left, top, width, height);
            if (chartKind.Equals("StackedLine", StringComparison.OrdinalIgnoreCase)) {
                chart.SetLineChartGrouping(C.GroupingValues.Stacked);
            } else if (chartKind.Equals("StackedLine100", StringComparison.OrdinalIgnoreCase)) {
                chart.SetLineChartGrouping(C.GroupingValues.PercentStacked);
            }

            return true;
        }

        if (chartKind.Equals("Pie", StringComparison.OrdinalIgnoreCase)) {
            chart = slide.AddPieChartPoints(data, left, top, width, height);
            return true;
        }

        if (chartKind.Equals("Doughnut", StringComparison.OrdinalIgnoreCase)) {
            chart = slide.AddDoughnutChartPoints(data, left, top, width, height);
            return true;
        }

        if (chartKind.Equals("Scatter", StringComparison.OrdinalIgnoreCase)) {
            if (!TryCreateScatterChartData(data, out PptCore.PowerPointScatterChartData? scatterData) || scatterData == null) {
                return false;
            }

            chart = slide.AddScatterChartPoints(scatterData, left, top, width, height);
            return true;
        }

        return false;
    }

    private static bool TryCreateScatterChartData(PptCore.PowerPointChartData data, out PptCore.PowerPointScatterChartData? scatterData) {
        scatterData = null;
        var series = new List<PptCore.PowerPointScatterChartSeries>();
        foreach (PptCore.PowerPointChartSeries item in data.Series) {
            if (item.XValues == null || item.XValues.Count != item.Values.Count) {
                return false;
            }

            series.Add(new PptCore.PowerPointScatterChartSeries(item.Name, item.XValues, item.Values));
        }

        if (series.Count == 0) {
            return false;
        }

        scatterData = new PptCore.PowerPointScatterChartData(series);
        return true;
    }

    private static string ReadChartKind(IElement item) {
        string meta = string.Join(" ", item.QuerySelectorAll(".officeimo-feature-meta").Select(element => element.TextContent));
        const string marker = "Type:";
        int index = meta.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (index < 0) {
            return "ClusteredColumn";
        }

        string value = meta.Substring(index + marker.Length).Split(';')[0].Trim();
        return value.Length == 0 ? "ClusteredColumn" : value;
    }

    private static bool TryGetImagePartType(string mediaType, out PptCore.ImagePartType imagePartType) {
        if (mediaType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) || mediaType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Jpeg;
            return true;
        }

        if (mediaType.Equals("image/gif", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Gif;
            return true;
        }

        if (mediaType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Bmp;
            return true;
        }

        if (mediaType.Equals("image/tiff", StringComparison.OrdinalIgnoreCase) || mediaType.Equals("image/tif", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Tiff;
            return true;
        }

        if (mediaType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Svg;
            return true;
        }

        if (mediaType.Equals("image/x-emf", StringComparison.OrdinalIgnoreCase) || mediaType.Equals("image/emf", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Emf;
            return true;
        }

        if (mediaType.Equals("image/x-wmf", StringComparison.OrdinalIgnoreCase) || mediaType.Equals("image/wmf", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Wmf;
            return true;
        }

        if (mediaType.Equals("image/x-icon", StringComparison.OrdinalIgnoreCase) ||
            mediaType.Equals("image/vnd.microsoft.icon", StringComparison.OrdinalIgnoreCase) ||
            mediaType.Equals("image/ico", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Icon;
            return true;
        }

        if (mediaType.Equals("image/x-pcx", StringComparison.OrdinalIgnoreCase) || mediaType.Equals("image/pcx", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Pcx;
            return true;
        }

        if (mediaType.Equals("image/png", StringComparison.OrdinalIgnoreCase) || mediaType.Equals("image/x-png", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Png;
            return true;
        }

        imagePartType = PptCore.ImagePartType.Png;
        return false;
    }

    private static bool IsElement(IElement element, string name) =>
        string.Equals(element.LocalName, name, StringComparison.OrdinalIgnoreCase);

    private static string NormalizeText(string? text) =>
        string.IsNullOrWhiteSpace(text) ? string.Empty : string.Join(" ", text!.Split((char[]?)null!, StringSplitOptions.RemoveEmptyEntries));

    private static string PreserveText(string? text) =>
        string.IsNullOrWhiteSpace(text) ? string.Empty : text!.Trim();
}
