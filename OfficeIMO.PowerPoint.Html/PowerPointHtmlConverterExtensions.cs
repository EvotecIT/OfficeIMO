using OfficeIMO.Drawing;
using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>
/// Extension methods enabling HTML conversions for OfficeIMO PowerPoint presentations.
/// </summary>
public static partial class PowerPointHtmlConverterExtensions {
    /// <summary>
    /// Converts a presentation to HTML.
    /// </summary>
    public static string ToHtml(this PptCore.PowerPointPresentation presentation, PowerPointHtmlSaveOptions? options = null) {
        return presentation.ToHtmlResult(options).Value;
    }

    /// <summary>Converts a presentation to HTML with operation-scoped visual diagnostics.</summary>
    public static PowerPointToHtmlResult ToHtmlResult(this PptCore.PowerPointPresentation presentation, PowerPointHtmlSaveOptions? options = null) {
        if (presentation == null) throw new ArgumentNullException(nameof(presentation));
        PowerPointHtmlSaveOptions operation = (options ?? new PowerPointHtmlSaveOptions()).Clone();
        var imageDiagnostics = new List<OfficeImageExportDiagnostic>();
        string html = operation.Profile == OfficeHtmlConversionProfile.PowerPointVisualReview
            ? ConvertVisual(presentation, operation, imageDiagnostics)
            : ConvertSemantic(presentation, operation);
        return new PowerPointToHtmlResult(html, imageDiagnostics);
    }

    /// <summary>
    /// Saves a presentation as HTML.
    /// </summary>
    public static void SaveAsHtml(this PptCore.PowerPointPresentation presentation, string path, PowerPointHtmlSaveOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("HTML path cannot be empty.", nameof(path));
        HtmlTextIO.Write(path, presentation.ToHtml(options));
    }

    private static string ConvertSemantic(PptCore.PowerPointPresentation presentation, PowerPointHtmlSaveOptions options) {
        IReadOnlyList<string> extractionProof = GetExtractionProof(presentation, options);
        var body = new StringBuilder();
        body.Append("<main class=\"officeimo-document\"");
        OfficeHtmlSemanticEnvelope.AppendRootAttributes(body, "powerpoint", options.Profile.ToString());
        body.Append('>');
        body.Append("<h1>").Append(OfficeHtmlText.Escape(GetTitle(options, "PowerPoint Presentation"))).Append("</h1>");

        int visibleIndex = 0;
        for (int i = 0; i < presentation.Slides.Count; i++) {
            PptCore.PowerPointSlide slide = presentation.Slides[i];
            if (!options.IncludeHiddenSlides && slide.Hidden) {
                continue;
            }

            visibleIndex++;
            AppendSemanticSlide(body, slide, i + 1, visibleIndex, extractionProof.Count > i ? extractionProof[i] : null, options);
        }

        body.Append("</main>");
        return Wrap(body.ToString(), options, GetTitle(options, "PowerPoint Presentation"));
    }

    private static string ConvertVisual(PptCore.PowerPointPresentation presentation, PowerPointHtmlSaveOptions options,
        IList<OfficeImageExportDiagnostic> imageDiagnostics) {
        IReadOnlyList<string> extractionProof = GetExtractionProof(presentation, options);
        var body = new StringBuilder();
        body.Append("<main class=\"officeimo-document\"");
        OfficeHtmlSemanticEnvelope.AppendRootAttributes(body, "powerpoint", options.Profile.ToString());
        body.Append('>');
        body.Append("<h1>").Append(OfficeHtmlText.Escape(GetTitle(options, "PowerPoint Visual Review"))).Append("</h1>");

        for (int i = 0; i < presentation.Slides.Count; i++) {
            PptCore.PowerPointSlide slide = presentation.Slides[i];
            if (!options.IncludeHiddenSlides && slide.Hidden) {
                continue;
            }

            AppendVisualSlide(body, presentation, slide, i + 1, extractionProof.Count > i ? extractionProof[i] : null, options, imageDiagnostics);
        }

        body.Append("</main>");
        return Wrap(body.ToString(), options, GetTitle(options, "PowerPoint Visual Review"));
    }

    private static void AppendSemanticSlide(StringBuilder body, PptCore.PowerPointSlide slide, int slideNumber, int visibleIndex, string? extractionProof, PowerPointHtmlSaveOptions options) {
        body.Append("<section class=\"officeimo-slide\" data-officeimo-slide=\"")
            .Append(slideNumber.ToString(CultureInfo.InvariantCulture))
            .Append('"');
        AppendHiddenSlideAttribute(body, slide);
        body.Append('>');
        body.Append("<h2>Slide ").Append(visibleIndex.ToString(CultureInfo.InvariantCulture)).Append("</h2>");

        AppendSemanticShapes(body, slide, options);

        AppendSlideFeatureInventory(body, slide, options);
        AppendExtractionProof(body, extractionProof, options);
        body.Append("</section>");
    }

    private static void AppendVisualSlide(StringBuilder body, PptCore.PowerPointPresentation presentation, PptCore.PowerPointSlide slide, int slideNumber, string? extractionProof, PowerPointHtmlSaveOptions options,
        IList<OfficeImageExportDiagnostic> imageDiagnostics) {
        double width = Math.Max(1D, presentation.SlideSize.WidthPoints);
        double height = Math.Max(1D, presentation.SlideSize.HeightPoints);
        body.Append("<section class=\"officeimo-slide\" data-officeimo-slide=\"")
            .Append(slideNumber.ToString(CultureInfo.InvariantCulture))
            .Append('"');
        AppendHiddenSlideAttribute(body, slide);
        body.Append('>');
        body.Append("<h2>Slide ").Append(slideNumber.ToString(CultureInfo.InvariantCulture)).Append("</h2>");
        body.Append("<div class=\"officeimo-visual-page\" data-officeimo-visual-owner=\"OfficeIMO.PowerPoint\" data-officeimo-visual-boundary=\"positioned-review\">");
        body.Append("<div class=\"officeimo-slide-canvas\" style=\"width:")
            .Append(FormatNumber(width))
            .Append("pt;height:")
            .Append(FormatNumber(height))
            .Append("pt;\">");

        var snapshotOptions = new PptCore.PowerPointImageExportOptions {
            IncludeHiddenShapes = options.IncludeHiddenShapes,
            IncludeTables = options.IncludeTables
        };
        PptCore.PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(snapshotOptions);
        IReadOnlyList<OfficeImageExportDiagnostic> slideDiagnostics = snapshot.Diagnostics;
        body.Append("<div class=\"officeimo-shared-slide-snapshot\" data-officeimo-visual-owner=\"OfficeIMO.Drawing\" data-officeimo-snapshot-diagnostics=\"")
            .Append(snapshot.Diagnostics.Count.ToString(CultureInfo.InvariantCulture))
            .Append("\">")
            .Append(OfficeDrawingSvgExporter.ToSvg(snapshot.Drawing))
            .Append("</div>");
        foreach (OfficeImageExportDiagnostic diagnostic in snapshot.Diagnostics) {
            imageDiagnostics.Add(diagnostic);
        }

        body.Append("<div class=\"officeimo-positioned-shape-metadata\" hidden>");
        foreach (PptCore.PowerPointShape shape in EnumerateVisualExportShapes(slide)) {
            if (!options.IncludeHiddenShapes && shape.Hidden) {
                continue;
            }

            AppendPositionedShape(body, shape, options);
        }

        body.Append("</div></div></div>");
        AppendSnapshotDiagnostics(body, slideDiagnostics);
        AppendExtractionProof(body, extractionProof, options);
        body.Append("</section>");
    }

    private static void AppendSnapshotDiagnostics(StringBuilder body,
        IEnumerable<OfficeImageExportDiagnostic> diagnostics) {
        List<OfficeImageExportDiagnostic> slideDiagnostics = diagnostics.ToList();
        if (slideDiagnostics.Count == 0) return;
        body.Append("<ul class=\"officeimo-snapshot-diagnostics\">");
        foreach (OfficeImageExportDiagnostic diagnostic in slideDiagnostics) {
            body.Append("<li data-officeimo-diagnostic-code=\"")
                .Append(OfficeHtmlText.EscapeAttribute(diagnostic.Code))
                .Append("\">")
                .Append(OfficeHtmlText.Escape(diagnostic.Message))
                .Append("</li>");
        }
        body.Append("</ul>");
    }

    private static IEnumerable<PptCore.PowerPointShape> EnumerateVisualExportShapes(PptCore.PowerPointSlide slide) {
        foreach (PptCore.PowerPointShape shape in slide.GetInheritedShapesForExport()) {
            yield return shape;
        }

        foreach (PptCore.PowerPointShape shape in slide.Shapes) {
            yield return shape;
        }
    }

    private static void AppendHiddenSlideAttribute(StringBuilder body, PptCore.PowerPointSlide slide) {
        if (slide.Hidden) {
            body.Append(" data-officeimo-hidden=\"true\"");
        }
    }

    private static void AppendPositionedShape(StringBuilder body, PptCore.PowerPointShape shape, PowerPointHtmlSaveOptions options) {
        double left = shape.LeftPoints;
        double top = shape.TopPoints;
        double width = Math.Max(1D, shape.WidthPoints);
        double height = Math.Max(1D, shape.HeightPoints);
        string contentClass = shape switch {
            PptCore.PowerPointTable => " officeimo-shape-table",
            PptCore.PowerPointPicture => " officeimo-shape-picture",
            PptCore.PowerPointChart => " officeimo-shape-chart",
            _ => string.Empty
        };
        string transform = BuildShapeTransformStyle(shape);
        string textFlow = shape is PptCore.PowerPointTextBox ? "white-space:pre-wrap;" : string.Empty;

        body.Append("<div class=\"officeimo-shape")
            .Append(contentClass)
            .Append("\" data-officeimo-shape=\"")
            .Append(OfficeHtmlText.EscapeAttribute(shape.ShapeContentType.ToString()))
            .Append("\" style=\"left:")
            .Append(FormatNumber(left))
            .Append("pt;top:")
            .Append(FormatNumber(top))
            .Append("pt;width:")
            .Append(FormatNumber(width))
            .Append("pt;height:")
            .Append(FormatNumber(height))
            .Append("pt;")
            .Append(transform)
            .Append(textFlow)
            .Append("\">");

        if (shape is PptCore.PowerPointTextBox textBox) {
            body.Append(OfficeHtmlText.Escape(NormalizeText(textBox.Text)));
        } else if (shape is PptCore.PowerPointTable table && options.IncludeTables) {
            AppendTable(body, table);
        } else if (shape is PptCore.PowerPointPicture picture) {
            AppendPictureShape(body, picture);
        } else if (shape is PptCore.PowerPointChart chart) {
            AppendChartVisual(body, chart, width, height);
        } else {
            body.Append("<div class=\"officeimo-shape-placeholder\">")
                .Append(OfficeHtmlText.Escape(GetShapeLabel(shape)))
                .Append("</div>");
        }

        body.Append("</div>");
    }

    private static string BuildShapeTransformStyle(PptCore.PowerPointShape shape) {
        var transforms = new List<string>(3);
        if (shape.Rotation.HasValue && Math.Abs(shape.Rotation.Value) > 0.001D) {
            transforms.Add("rotate(" + FormatNumber(shape.Rotation.Value) + "deg)");
        }

        if (shape.HorizontalFlip == true) {
            transforms.Add("scaleX(-1)");
        }

        if (shape.VerticalFlip == true) {
            transforms.Add("scaleY(-1)");
        }

        return transforms.Count == 0 ? string.Empty : "transform:" + string.Join(" ", transforms) + ";";
    }

    private static void AppendSlideFeatureInventory(StringBuilder body, PptCore.PowerPointSlide slide, PowerPointHtmlSaveOptions options) {
        IEnumerable<PptCore.PowerPointPicture> pictures = options.IncludeHiddenShapes
            ? slide.Pictures
            : slide.Pictures.Where(picture => !picture.Hidden);
        IEnumerable<PptCore.PowerPointChart> charts = options.IncludeHiddenShapes
            ? slide.Charts
            : slide.Charts.Where(chart => !chart.Hidden);

        AppendPictureInventory(body, pictures);
        AppendChartInventory(body, charts);
    }

    private static void AppendPictureInventory(StringBuilder body, IEnumerable<PptCore.PowerPointPicture> pictures) {
        List<PptCore.PowerPointPicture> pictureList = pictures.ToList();
        if (pictureList.Count == 0) {
            return;
        }

        body.Append("<section class=\"officeimo-feature officeimo-images\"><h3>Pictures</h3><ul class=\"officeimo-feature-list\">");
        foreach (PptCore.PowerPointPicture picture in pictureList) {
            string label = GetShapeLabel(picture);
            body.Append("<li class=\"officeimo-feature-item\" data-officeimo-layer-kind=\"picture\" data-officeimo-layer-index=\"")
                .Append(picture.DrawingOrder.ToString(CultureInfo.InvariantCulture))
                .Append('"');
            AppendDataAttribute(body, "data-officeimo-left", picture.LeftPoints, omitWhenZero: false);
            AppendDataAttribute(body, "data-officeimo-top", picture.TopPoints, omitWhenZero: false);
            AppendDataAttribute(body, "data-officeimo-width", picture.WidthPoints, omitWhenZero: false);
            AppendDataAttribute(body, "data-officeimo-height", picture.HeightPoints, omitWhenZero: false);
            AppendDataAttribute(body, "data-officeimo-rotation", picture.Rotation ?? 0D);
            AppendDataAttribute(body, "data-officeimo-flip-horizontal", picture.HorizontalFlip == true);
            AppendDataAttribute(body, "data-officeimo-flip-vertical", picture.VerticalFlip == true);
            AppendDataAttribute(body, "data-officeimo-crop-left", picture.CropLeftRatio);
            AppendDataAttribute(body, "data-officeimo-crop-top", picture.CropTopRatio);
            AppendDataAttribute(body, "data-officeimo-crop-right", picture.CropRightRatio);
            AppendDataAttribute(body, "data-officeimo-crop-bottom", picture.CropBottomRatio);
            body.Append("><span class=\"officeimo-feature-label\">")
                .Append(OfficeHtmlText.Escape(label))
                .Append("</span><div class=\"officeimo-feature-meta\">Size: ")
                .Append(FormatNumber(picture.WidthPoints))
                .Append("pt x ")
                .Append(FormatNumber(picture.HeightPoints))
                .Append("pt; Position: ")
                .Append(FormatNumber(picture.LeftPoints))
                .Append("pt, ")
                .Append(FormatNumber(picture.TopPoints))
                .Append("pt; Type: ")
                .Append(OfficeHtmlText.Escape(picture.ContentType ?? string.Empty))
                .Append("</div>");
            AppendPicturePreview(body, picture, label);
            body.Append("</li>");
        }

        body.Append("</ul></section>");
    }

    private static void AppendDataAttribute(StringBuilder body, string name, double value, bool omitWhenZero = true) {
        if (omitWhenZero && Math.Abs(value) < 0.0000001D) {
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

    private static void AppendChartInventory(StringBuilder body, IEnumerable<PptCore.PowerPointChart> charts) {
        List<PptCore.PowerPointChart> chartList = charts.ToList();
        if (chartList.Count == 0) {
            return;
        }

        body.Append("<section class=\"officeimo-feature officeimo-charts\"><h3>Charts</h3><ul class=\"officeimo-feature-list\">");
        foreach (PptCore.PowerPointChart chart in chartList) {
            body.Append("<li class=\"officeimo-feature-item\" data-officeimo-layer-kind=\"chart\" data-officeimo-layer-index=\"")
                .Append(chart.DrawingOrder.ToString(CultureInfo.InvariantCulture))
                .Append('"');
            AppendDataAttribute(body, "data-officeimo-left", chart.LeftPoints, omitWhenZero: false);
            AppendDataAttribute(body, "data-officeimo-top", chart.TopPoints, omitWhenZero: false);
            AppendDataAttribute(body, "data-officeimo-width", chart.WidthPoints, omitWhenZero: false);
            AppendDataAttribute(body, "data-officeimo-height", chart.HeightPoints, omitWhenZero: false);
            AppendDataAttribute(body, "data-officeimo-rotation", chart.Rotation ?? 0D);
            AppendDataAttribute(body, "data-officeimo-flip-horizontal", chart.HorizontalFlip == true);
            AppendDataAttribute(body, "data-officeimo-flip-vertical", chart.VerticalFlip == true);
            body.Append("><div class=\"officeimo-feature-meta\">Size: ")
                .Append(FormatNumber(chart.WidthPoints))
                .Append("pt x ")
                .Append(FormatNumber(chart.HeightPoints))
                .Append("pt; Position: ")
                .Append(FormatNumber(chart.LeftPoints))
                .Append("pt, ")
                .Append(FormatNumber(chart.TopPoints))
                .Append("pt</div>");
            AppendChartSummary(body, chart);
            body.Append("</li>");
        }

        body.Append("</ul></section>");
    }

    private static void AppendPictureShape(StringBuilder body, PptCore.PowerPointPicture picture) {
        string label = GetShapeLabel(picture);
        string? dataUri = TryCreatePictureDataUri(picture);
        if (dataUri == null) {
            body.Append("<div class=\"officeimo-shape-placeholder\">")
                .Append(OfficeHtmlText.Escape(label))
                .Append("</div>");
            return;
        }

        body.Append("<img alt=\"")
            .Append(OfficeHtmlText.EscapeAttribute(label))
            .Append("\" src=\"")
            .Append(OfficeHtmlText.EscapeAttribute(dataUri))
            .Append("\">");
    }

    private static void AppendPicturePreview(StringBuilder body, PptCore.PowerPointPicture picture, string label) {
        string? dataUri = TryCreatePictureDataUri(picture);
        if (dataUri == null) {
            body.Append("<div class=\"officeimo-diagnostic\">Picture bytes unavailable.</div>");
            return;
        }

        body.Append("<img class=\"officeimo-inline-image\" alt=\"")
            .Append(OfficeHtmlText.EscapeAttribute(label))
            .Append("\" src=\"")
            .Append(OfficeHtmlText.EscapeAttribute(dataUri))
            .Append("\">");
    }

    private static string? TryCreatePictureDataUri(PptCore.PowerPointPicture picture) {
        try {
            byte[] bytes = picture.GetImageBytes();
            if (bytes.Length == 0) {
                return null;
            }

            string contentType = string.IsNullOrWhiteSpace(picture.ContentType) ? "image/png" : picture.ContentType!;
            return "data:" + contentType + ";base64," + Convert.ToBase64String(bytes);
        } catch {
            return null;
        }
    }

    private static void AppendChartVisual(StringBuilder body, PptCore.PowerPointChart chart, double width, double height) {
        if (TryCreateOfficeChartSnapshot(chart, width, height, out OfficeChartSnapshot? officeSnapshot, out string? warning) && officeSnapshot != null) {
            try {
                OfficeChartRenderingResult rendering = OfficeChartDrawingRenderer.RenderWithQuality(officeSnapshot);
                body.Append("<div class=\"officeimo-chart-rendered\" data-officeimo-visual-owner=\"OfficeIMO.Drawing\" data-officeimo-chart-kind=\"")
                    .Append(OfficeHtmlText.EscapeAttribute(officeSnapshot.ChartKind.ToString()))
                    .Append("\">")
                    .Append(OfficeDrawingSvgExporter.ToSvg(rendering.Drawing));
                if (rendering.QualityReport.HasIssues) {
                    body.Append("<div class=\"officeimo-diagnostic\">Shared chart renderer reported quality warnings: ")
                        .Append(OfficeHtmlText.Escape(FormatQualityIssues(rendering.QualityReport)))
                        .Append("</div>");
                }

                body.Append("</div>");
                return;
            } catch (Exception ex) {
                warning = "Chart visual rendering fell back to review placeholder because the shared Drawing renderer failed: " + ex.Message;
            }
        }

        body.Append("<div class=\"officeimo-shape-placeholder officeimo-chart-placeholder\">");
        AppendChartSummary(body, chart);
        AppendChartBars(body, chart);
        body.Append("<div class=\"officeimo-diagnostic\">")
            .Append(OfficeHtmlText.Escape(string.IsNullOrWhiteSpace(warning) ? "Chart visual rendering is reported as a review placeholder; chart data is preserved as snapshot metadata." : warning!))
            .Append("</div>");
        body.Append("</div>");
    }

    private static bool TryCreateOfficeChartSnapshot(PptCore.PowerPointChart chart, double width, double height, out OfficeChartSnapshot? officeSnapshot, out string? warning) {
        officeSnapshot = null;
        warning = null;
        if (!chart.TryGetSnapshot(out PptCore.PowerPointChartSnapshot snapshot)) {
            warning = "Chart snapshot unavailable; visual review is reported as a placeholder.";
            return false;
        }

        try {
            var series = snapshot.Data.Series
                .Select(item => new OfficeChartSeries(item.Name, item.Values, item.XValues, item.Color,
                    pointColors: null, showMarkers: true, strokeWidth: item.StrokeWidth,
                    renderKind: item.ChartKind.HasValue ? MapChartKind(item.ChartKind.Value) : null,
                    axisGroup: item.AxisGroup))
                .ToList();
            var data = new OfficeChartData(snapshot.Data.Categories, series);
            officeSnapshot = new OfficeChartSnapshot(
                snapshot.Name,
                snapshot.Title,
                MapChartKind(snapshot.ChartKind),
                data,
                Math.Max(1D, width),
                Math.Max(1D, height));
            return true;
        } catch (Exception ex) {
            warning = "Chart snapshot could not be mapped to the shared Drawing chart model: " + ex.Message;
            return false;
        }
    }

    private static OfficeChartKind MapChartKind(PptCore.PowerPointChartSnapshotKind kind) {
        switch (kind) {
            case PptCore.PowerPointChartSnapshotKind.ClusteredColumn:
                return OfficeChartKind.ColumnClustered;
            case PptCore.PowerPointChartSnapshotKind.StackedColumn:
                return OfficeChartKind.ColumnStacked;
            case PptCore.PowerPointChartSnapshotKind.StackedColumn100:
                return OfficeChartKind.ColumnStacked100;
            case PptCore.PowerPointChartSnapshotKind.ClusteredBar:
                return OfficeChartKind.BarClustered;
            case PptCore.PowerPointChartSnapshotKind.StackedBar:
                return OfficeChartKind.BarStacked;
            case PptCore.PowerPointChartSnapshotKind.StackedBar100:
                return OfficeChartKind.BarStacked100;
            case PptCore.PowerPointChartSnapshotKind.Line:
                return OfficeChartKind.Line;
            case PptCore.PowerPointChartSnapshotKind.StackedLine:
                return OfficeChartKind.LineStacked;
            case PptCore.PowerPointChartSnapshotKind.StackedLine100:
                return OfficeChartKind.LineStacked100;
            case PptCore.PowerPointChartSnapshotKind.Area:
                return OfficeChartKind.Area;
            case PptCore.PowerPointChartSnapshotKind.StackedArea:
                return OfficeChartKind.AreaStacked;
            case PptCore.PowerPointChartSnapshotKind.StackedArea100:
                return OfficeChartKind.AreaStacked100;
            case PptCore.PowerPointChartSnapshotKind.Radar:
                return OfficeChartKind.Radar;
            case PptCore.PowerPointChartSnapshotKind.Scatter:
                return OfficeChartKind.Scatter;
            case PptCore.PowerPointChartSnapshotKind.Pie:
                return OfficeChartKind.Pie;
            case PptCore.PowerPointChartSnapshotKind.Doughnut:
                return OfficeChartKind.Doughnut;
            default:
                throw new NotSupportedException("PowerPoint chart kind '" + kind + "' is not supported by the shared Drawing chart renderer.");
        }
    }

    private static string FormatQualityIssues(OfficeDrawingQualityReport qualityReport) {
        return string.Join("; ", qualityReport.Issues.Select(issue => issue.ToString()));
    }

    private static void AppendChartSummary(StringBuilder body, PptCore.PowerPointChart chart) {
        if (chart.TryGetSnapshot(out PptCore.PowerPointChartSnapshot snapshot)) {
            string label = string.IsNullOrWhiteSpace(snapshot.Title)
                ? string.IsNullOrWhiteSpace(snapshot.Name) ? "Chart" : snapshot.Name
                : snapshot.Title!;
            body.Append("<span class=\"officeimo-feature-label\">")
                .Append(OfficeHtmlText.Escape(label))
                .Append("</span><div class=\"officeimo-feature-meta\">Type: ")
                .Append(OfficeHtmlText.Escape(snapshot.ChartKind.ToString()))
                .Append("; Series: ")
                .Append(snapshot.Data.Series.Count.ToString(CultureInfo.InvariantCulture))
                .Append("; Categories: ")
                .Append(snapshot.Data.Categories.Count.ToString(CultureInfo.InvariantCulture))
                .Append("</div>");
            AppendChartDataTable(body, snapshot.Data);
            return;
        }

        body.Append("<span class=\"officeimo-feature-label\">")
            .Append(OfficeHtmlText.Escape(GetShapeLabel(chart)))
            .Append("</span><div class=\"officeimo-diagnostic\">Chart snapshot unavailable.</div>");
    }

    private static void AppendChartDataTable(StringBuilder body, PptCore.PowerPointChartData data) {
        body.Append("<table class=\"officeimo-chart-data\"><thead><tr><th>Series</th>");
        foreach (string category in data.Categories) {
            body.Append("<th>")
                .Append(OfficeHtmlText.Escape(category))
                .Append("</th>");
        }

        body.Append("</tr></thead><tbody>");
        foreach (PptCore.PowerPointChartSeries series in data.Series) {
            body.Append("<tr><th>")
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

    private static void AppendChartBars(StringBuilder body, PptCore.PowerPointChart chart) {
        if (!chart.TryGetSnapshot(out PptCore.PowerPointChartSnapshot snapshot)) {
            return;
        }

        IReadOnlyList<double> values = snapshot.Data.Series
            .SelectMany(series => series.Values)
            .Take(12)
            .ToList();
        if (values.Count == 0) {
            return;
        }

        double max = Math.Max(1D, values.Max(value => Math.Abs(value)));
        body.Append("<div class=\"officeimo-chart-bars\" aria-hidden=\"true\">");
        foreach (double value in values) {
            double height = Math.Max(6D, Math.Abs(value) / max * 60D);
            body.Append("<span style=\"height:")
                .Append(FormatNumber(height))
                .Append("px\"></span>");
        }

        body.Append("</div>");
    }

    private static void AppendTable(StringBuilder body, PptCore.PowerPointTable table, bool includeShapeMetadata = false) {
        body.Append("<table class=\"officeimo-table\"");
        if (includeShapeMetadata) {
            AppendSemanticShapeAttributes(body, table, "table");
        }

        body.Append("><tbody>");
        foreach (PptCore.PowerPointTableRow row in table.RowItems) {
            body.Append("<tr>");
            foreach (PptCore.PowerPointTableCell cell in row.Cells) {
                if (cell.IsMergedCell) {
                    continue;
                }

                body.Append("<td");
                (int rows, int columns) = cell.Merge;
                if (rows > 1) {
                    body.Append(" rowspan=\"").Append(rows.ToString(CultureInfo.InvariantCulture)).Append('"');
                }

                if (columns > 1) {
                    body.Append(" colspan=\"").Append(columns.ToString(CultureInfo.InvariantCulture)).Append('"');
                }

                body.Append('>').Append(OfficeHtmlText.Escape(NormalizeText(cell.Text))).Append("</td>");
            }

            body.Append("</tr>");
        }

        body.Append("</tbody></table>");
    }

    private static void AppendExtractionProof(StringBuilder body, string? extractionProof, PowerPointHtmlSaveOptions options) {
        if (!options.IncludeExtractionProof || string.IsNullOrWhiteSpace(extractionProof)) {
            return;
        }

        body.Append("<details><summary>Extraction proof</summary><pre class=\"officeimo-source-markdown\">")
            .Append(OfficeHtmlText.Escape(extractionProof))
            .Append("</pre></details>");
    }

    private static IReadOnlyList<string> GetExtractionProof(PptCore.PowerPointPresentation presentation, PowerPointHtmlSaveOptions options) {
        if (!options.IncludeExtractionProof) {
            return Array.Empty<string>();
        }

        return presentation.ExtractMarkdownChunks(
                new PptCore.PowerPointExtractionExtensions.PowerPointExtractOptions {
                    IncludeNotes = options.IncludeNotes,
                    IncludeTables = options.IncludeTables,
                    IncludeHiddenShapes = options.IncludeHiddenShapes
                })
            .Select(chunk => chunk.Markdown ?? chunk.Text ?? string.Empty)
            .ToList();
    }

    private static string Wrap(string body, PowerPointHtmlSaveOptions options, string title) {
        return OfficeHtmlDocumentShell.WrapBody(body, new OfficeHtmlDocumentOptions {
            Title = title,
            Theme = options.Theme,
            IncludeDefaultStyles = options.IncludeDefaultStyles,
            BodyClass = "officeimo-html officeimo-powerpoint-html"
        });
    }

    private static string GetTitle(PowerPointHtmlSaveOptions options, string fallback) {
        return string.IsNullOrWhiteSpace(options.Title) ? fallback : options.Title!;
    }

    private static string NormalizeText(string? text) {
        return (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Trim();
    }

    private static string GetShapeLabel(PptCore.PowerPointShape shape) {
        if (!string.IsNullOrWhiteSpace(shape.AltText)) {
            return shape.AltText!;
        }

        if (!string.IsNullOrWhiteSpace(shape.Name)) {
            return shape.Name!;
        }

        return shape.ShapeContentType.ToString();
    }

    private static string FormatNumber(double value) {
        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }
}
