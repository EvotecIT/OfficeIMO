using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

/// <summary>Projects typed OneNote pages into the shared dependency-free Drawing scene.</summary>
public static partial class OneNotePageRenderer {
    /// <summary>Number of typographic points represented by one native OneNote half-inch unit.</summary>
    public const double PointsPerHalfInch = 36D;

    /// <summary>Creates a Drawing scene and structured rendering diagnostics for a page.</summary>
    public static OneNotePageVisualSnapshot CreateSnapshot(OneNotePage page, OneNotePageRenderingOptions? options = null) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        OneNotePageRenderingOptions effective = options?.Clone() ?? new OneNotePageRenderingOptions();
        effective.Validate();
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var imageCache = new Dictionary<OneNoteImage, ImageRenderData>();
        (double width, double height) = ResolveCanvasSize(page, effective, imageCache);
        var drawing = new OfficeDrawing(width, height);
        OfficeShape background = OfficeShape.Rectangle(width, height);
        background.FillColor = effective.BackgroundColor;
        background.StrokeWidth = 0D;
        drawing.AddShape(background, 0D, 0D);

        var context = new RenderContext(drawing, effective, diagnostics, page.RightToLeft == true, imageCache);
        foreach (OneNoteElement element in page.DirectContent) {
            if (element is OneNoteImage image && image.IsBackground == true) {
                context.RenderElement(element, 0D, 0D, width, 0D, forcePageBounds: true);
            }
        }

        double marginLeft = ResolveMargin(page.Margins.OriginX, page.Margins.Left, 1D) * PointsPerHalfInch;
        double marginTop = ResolveMargin(page.Margins.OriginY, page.Margins.Top, 0.5D) * PointsPerHalfInch;
        double marginRight = Math.Max(0D, (page.Margins.Right ?? 1D) * PointsPerHalfInch);
        double bodyWidth = Math.Max(1D, width - marginLeft - marginRight);
        double flowY = marginTop;
        double pendingFlowSpace = 0D;
        if (effective.IncludeTitle && !string.IsNullOrWhiteSpace(page.Title)) {
            double titleHeight = Math.Min(42D, Math.Max(22D, height - flowY));
            drawing.AddText(
                page.Title,
                Clamp(marginLeft, 0D, Math.Max(0D, width - 1D)),
                Clamp(flowY, 0D, Math.Max(0D, height - 1D)),
                Math.Max(1D, Math.Min(bodyWidth, width - marginLeft)),
                Math.Max(1D, Math.Min(titleHeight, height - flowY)),
                new OfficeFontInfo(effective.DefaultFont.FamilyName, 17D, OfficeFontStyle.Bold),
                OfficeColor.Black,
                page.RightToLeft == true ? OfficeTextAlignment.Right : OfficeTextAlignment.Left,
                wrapText: false);
            flowY += titleHeight + 12D;
        }

        foreach (OneNoteOutline outline in page.Outlines) {
            double x = outline.Layout?.X.HasValue == true ? outline.Layout.X.Value * PointsPerHalfInch : marginLeft;
            double y = outline.Layout?.Y.HasValue == true ? outline.Layout.Y.Value * PointsPerHalfInch : flowY + pendingFlowSpace;
            double outlineWidth = ResolveWidth(outline.Layout, bodyWidth, width - x);
            double used = context.RenderElement(outline, x, y, outlineWidth, 0D);
            if (outline.Layout?.Y.HasValue != true) {
                flowY = Math.Max(flowY, y + used);
                pendingFlowSpace = 6D;
            }
        }

        foreach (OneNoteElement element in page.DirectContent) {
            if (element is OneNoteImage image && image.IsBackground == true) continue;
            double x = element.Layout?.X.HasValue == true ? element.Layout.X.Value * PointsPerHalfInch : marginLeft;
            double y = element.Layout?.Y.HasValue == true
                ? element.Layout.Y.Value * PointsPerHalfInch
                : flowY + Math.Max(pendingFlowSpace, RenderContext.ParagraphSpaceBefore(element));
            double available = ResolveWidth(element.Layout, bodyWidth, width - x);
            double used = context.RenderElement(element, x, y, available, 0D);
            if (element.Layout?.Y.HasValue != true) {
                flowY = Math.Max(flowY, y + used);
                pendingFlowSpace = element is OneNoteParagraph ? RenderContext.ParagraphSpaceAfter(element) : 6D;
            }
        }

        return new OneNotePageVisualSnapshot(drawing, diagnostics.AsReadOnly());
    }

    /// <summary>Creates a Drawing scene for a page.</summary>
    public static OfficeDrawing Render(OneNotePage page, OneNotePageRenderingOptions? options = null) =>
        CreateSnapshot(page, options).Drawing;

    private static (double Width, double Height) ResolveCanvasSize(
        OneNotePage page,
        OneNotePageRenderingOptions options,
        IDictionary<OneNoteImage, ImageRenderData> imageCache) {
        (double namedWidth, double namedHeight) = page.PageSize.HasValue &&
            page.PageSize != OneNotePageSize.Automatic && page.PageSize != OneNotePageSize.Custom
            ? OneNotePageGeometry.GetNamedSizePoints(page.PageSize.Value, page.Orientation)
            : OneNotePageGeometry.GetNamedSizePoints(OneNotePageSize.Letter, page.Orientation);
        double width = page.Width.HasValue ? page.Width.Value * PointsPerHalfInch : namedWidth;
        double height = page.Height.HasValue ? page.Height.Value * PointsPerHalfInch : namedHeight;
        bool automatic = page.PageSize == null || page.PageSize == OneNotePageSize.Automatic;
        if (automatic) {
            width = Math.Max(width, options.AutomaticPageWidthPoints);
            (double preliminaryRight, _) = EstimateContentBounds(page, options, width, imageCache);
            width = Math.Max(width, preliminaryRight + options.AutomaticPagePaddingPoints);
            (double contentRight, double contentBottom) = EstimateContentBounds(page, options, width, imageCache);
            width = Math.Max(width, Math.Max(options.AutomaticPageWidthPoints, contentRight + options.AutomaticPagePaddingPoints));
            height = Math.Max(height, Math.Max(options.AutomaticPageHeightPoints, contentBottom + options.AutomaticPagePaddingPoints));
        }
        return (Math.Max(1D, width), Math.Max(1D, height));
    }

    private static (double Right, double Bottom) EstimateContentBounds(
        OneNotePage page,
        OneNotePageRenderingOptions options,
        double canvasWidth,
        IDictionary<OneNoteImage, ImageRenderData> imageCache) {
        var estimator = new RenderContext(
            new OfficeDrawing(Math.Max(1D, canvasWidth), 1D),
            options,
            new List<OfficeImageExportDiagnostic>(),
            page.RightToLeft == true,
            imageCache);
        double marginLeft = ResolveMargin(page.Margins.OriginX, page.Margins.Left, 1D) * PointsPerHalfInch;
        double marginTop = ResolveMargin(page.Margins.OriginY, page.Margins.Top, 0.5D) * PointsPerHalfInch;
        double marginRight = Math.Max(0D, (page.Margins.Right ?? 1D) * PointsPerHalfInch);
        double bodyWidth = Math.Max(1D, canvasWidth - marginLeft - marginRight);
        double right = 0D;
        double bottom = 0D;
        double flow = marginTop;
        double pendingSpace = 0D;
        if (options.IncludeTitle && !string.IsNullOrWhiteSpace(page.Title)) flow += 54D;
        foreach (OneNoteOutline outline in page.Outlines) {
            double x = outline.Layout?.X.HasValue == true ? outline.Layout.X.Value * PointsPerHalfInch : marginLeft;
            double y = outline.Layout?.Y.HasValue == true ? outline.Layout.Y.Value * PointsPerHalfInch : flow + pendingSpace;
            double width = ResolveEstimatedWidth(outline, bodyWidth, options);
            (double childRight, double childBottom) = estimator.MeasureElementsBounds(outline.Children, width);
            double height = outline.Layout?.Height.HasValue == true
                ? outline.Layout.Height.Value * PointsPerHalfInch
                : childBottom;
            right = Math.Max(right, x + Math.Max(width, childRight));
            bottom = Math.Max(bottom, y + height);
            if (outline.Layout?.Y.HasValue != true) {
                flow = Math.Max(flow, y + height);
                pendingSpace = 6D;
            }
        }
        foreach (OneNoteElement element in page.DirectContent) {
            if (element is OneNoteImage background && background.IsBackground == true) continue;
            double x = element.Layout?.X.HasValue == true ? element.Layout.X.Value * PointsPerHalfInch : marginLeft;
            double y = element.Layout?.Y.HasValue == true
                ? element.Layout.Y.Value * PointsPerHalfInch
                : flow + Math.Max(pendingSpace, RenderContext.ParagraphSpaceBefore(element));
            double elementWidth = ResolveEstimatedWidth(element, bodyWidth, options);
            double elementHeight = estimator.MeasureElementHeight(element, elementWidth);
            right = Math.Max(right, x + estimator.MeasureElementWidthExtent(element, elementWidth));
            bottom = Math.Max(bottom, y + elementHeight);
            if (element.Layout?.Y.HasValue != true) {
                flow = Math.Max(flow, y + elementHeight);
                pendingSpace = element is OneNoteParagraph ? RenderContext.ParagraphSpaceAfter(element) : 6D;
            }
        }
        return (right, Math.Max(bottom, flow + pendingSpace));
    }

    private static double ResolveEstimatedWidth(
        OneNoteElement element,
        double fallback,
        OneNotePageRenderingOptions options) {
        if (element.Layout?.Width.HasValue == true) return Math.Max(1D, element.Layout.Width.Value * PointsPerHalfInch);
        if (element is OneNoteImage image && image.WidthHalfInches.HasValue) {
            return Math.Max(1D, image.WidthHalfInches.Value * PointsPerHalfInch);
        }
        if (element is OneNoteInk ink) {
            OfficeInkBounds bounds = ink.Ink.GetBounds();
            if (!bounds.IsEmpty) return Math.Max(1D, (bounds.X + bounds.Width) * PointsPerHalfInch);
        }
        if (element is OneNoteMath math) {
            return Math.Max(1D, OfficeMathRenderer.Measure(math.GetExpression(), options.Math).Width);
        }
        return Math.Max(1D, fallback);
    }

    private static double ResolveMargin(double? origin, double? margin, double fallback) => origin ?? margin ?? fallback;

    private static double ResolveWidth(OneNoteLayout? layout, double fallback, double remaining) {
        double value = layout?.Width.HasValue == true ? layout.Width.Value * PointsPerHalfInch : fallback;
        return Math.Max(1D, Math.Min(value, Math.Max(1D, remaining)));
    }

    private static double Clamp(double value, double minimum, double maximum) => Math.Max(minimum, Math.Min(maximum, value));
}
