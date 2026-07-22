using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>
/// Dependency-free text-selection and hit-test projection for one rendered PDF page.
/// Coordinates use the visual page's top-left origin after CropBox and page rotation are applied.
/// </summary>
public sealed class PdfPageInteractionMap {
    private PdfPageInteractionMap(
        int pageNumber,
        double width,
        double height,
        IReadOnlyList<PdfPageInteractionRegion> regions) {
        PageNumber = pageNumber;
        Width = width;
        Height = height;
        Regions = regions;
        TextRegions = regions.Where(static region => region.Kind == PdfInteractionKind.Text).ToArray();
    }

    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }

    /// <summary>Visual page width in PDF points after page rotation.</summary>
    public double Width { get; }

    /// <summary>Visual page height in PDF points after page rotation.</summary>
    public double Height { get; }

    /// <summary>All text, link, annotation, and form-widget regions.</summary>
    public IReadOnlyList<PdfPageInteractionRegion> Regions { get; }

    /// <summary>Text-element regions in content extraction order.</summary>
    public IReadOnlyList<PdfPageInteractionRegion> TextRegions { get; }

    /// <summary>Builds an interaction map from PDF bytes.</summary>
    public static PdfPageInteractionMap Create(
        byte[] pdf,
        int pageNumber,
        PdfPageInteractionOptions? options = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), "Page number must be positive.");
        }

        PdfPageInteractionOptions effective = options ?? new PdfPageInteractionOptions();
        if (effective.MaxTextRegions <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "Maximum text regions must be positive.");
        }

        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        if (pageNumber > document.Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), "Page number exceeds the PDF page count.");
        }

        PdfReadPage page = document.Pages[pageNumber - 1];
        (double Width, double Height) size = page.GetInteractionPageSize();
        PdfPageInfo pageInfo = PdfInspector.Inspect(pdf, readOptions).Pages[pageNumber - 1];
        var regions = new List<PdfPageInteractionRegion>();
        AddTextRegions(page, size.Width, size.Height, effective, regions);
        AddLinkRegions(page, pageInfo, size.Height, regions);
        AddAnnotationRegions(page, pageInfo, size.Height, regions);
        AddFormWidgetRegions(page, pageInfo, size.Height, regions);
        return new PdfPageInteractionMap(pageNumber, size.Width, size.Height, regions.AsReadOnly());
    }

    /// <summary>Returns all regions containing a visual top-left page coordinate.</summary>
    public IReadOnlyList<PdfPageInteractionRegion> HitTest(double x, double y, double tolerance = 0D) {
        if (!IsFinite(x) || !IsFinite(y)) {
            throw new ArgumentOutOfRangeException(nameof(x), "Hit-test coordinates must be finite.");
        }

        var matches = new List<PdfPageInteractionRegion>();
        for (int i = Regions.Count - 1; i >= 0; i--) {
            if (Regions[i].Quad.Contains(x, y, tolerance)) {
                matches.Add(Regions[i]);
            }
        }

        return matches.Count == 0 ? Array.Empty<PdfPageInteractionRegion>() : matches.AsReadOnly();
    }

    /// <summary>Returns text elements whose quads intersect a visual top-left selection rectangle.</summary>
    public IReadOnlyList<PdfPageInteractionRegion> SelectText(double x1, double y1, double x2, double y2) {
        if (!IsFinite(x1) || !IsFinite(y1) || !IsFinite(x2) || !IsFinite(y2)) {
            throw new ArgumentOutOfRangeException(nameof(x1), "Selection coordinates must be finite.");
        }

        double left = Math.Min(x1, x2);
        double top = Math.Min(y1, y2);
        double right = Math.Max(x1, x2);
        double bottom = Math.Max(y1, y2);
        var matches = new List<PdfPageInteractionRegion>();
        for (int i = 0; i < TextRegions.Count; i++) {
            if (TextRegions[i].Quad.Intersects(left, top, right, bottom)) {
                matches.Add(TextRegions[i]);
            }
        }

        return matches.Count == 0 ? Array.Empty<PdfPageInteractionRegion>() : matches.AsReadOnly();
    }

    /// <summary>Concatenates selected text elements in extraction order.</summary>
    public string GetSelectedText(double x1, double y1, double x2, double y2) {
        IReadOnlyList<PdfPageInteractionRegion> selected = SelectText(x1, y1, x2, y2);
        var text = new StringBuilder();
        for (int i = 0; i < selected.Count; i++) {
            text.Append(selected[i].Text);
        }

        return text.ToString();
    }

    private static void AddTextRegions(
        PdfReadPage page,
        double pageWidth,
        double pageHeight,
        PdfPageInteractionOptions options,
        List<PdfPageInteractionRegion> regions) {
        IReadOnlyList<PdfTextSpan> spans = page.GetInteractionTextSpans();
        int textIndex = 0;
        int processedTextElements = 0;
        for (int spanIndex = 0; spanIndex < spans.Count; spanIndex++) {
            PdfTextSpan span = spans[spanIndex];
            if (string.IsNullOrEmpty(span.Text) || (!span.IsVisible && !options.IncludeInvisibleText)) {
                continue;
            }

            TextElementEnumerator enumerator = StringInfo.GetTextElementEnumerator(span.Text);
            int elementCount = 0;
            while (enumerator.MoveNext()) {
                elementCount++;
                processedTextElements++;
                if (processedTextElements > options.MaxTextRegions) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.InteractionRegions, options.MaxTextRegions, processedTextElements);
                }
            }

            if (elementCount == 0) {
                continue;
            }

            double totalAdvance = Math.Max(Math.Abs(span.Advance), elementCount * Math.Max(1D, span.FontSize * 0.5D));
            double elementAdvance = totalAdvance / elementCount;
            double radians = span.RotationDegrees * Math.PI / 180D;
            double directionX = Math.Cos(radians);
            double directionY = Math.Sin(radians);
            double normalX = -directionY;
            double normalY = directionX;
            double ascent = Math.Max(1D, span.FontSize);
            double descent = Math.Max(0.5D, span.FontSize * 0.2D);
            enumerator = StringInfo.GetTextElementEnumerator(span.Text);
            int elementIndex = 0;
            while (enumerator.MoveNext()) {
                string element = (string)enumerator.Current!;
                int currentElementIndex = elementIndex++;
                double startAdvance = currentElementIndex * elementAdvance;
                double endAdvance = (currentElementIndex + 1) * elementAdvance;
                double startX = span.X + directionX * startAdvance;
                double startY = span.Y + directionY * startAdvance;
                double endX = span.X + directionX * endAdvance;
                double endY = span.Y + directionY * endAdvance;
                PdfSelectionQuad quad = FromVisualBaseline(
                    startX, startY, endX, endY,
                    normalX, normalY, ascent, descent, pageHeight);
                if (!quad.Intersects(0D, 0D, pageWidth, pageHeight)) {
                    continue;
                }
                regions.Add(new PdfPageInteractionRegion(
                    PdfInteractionKind.Text,
                    quad,
                    text: element,
                    textIndex: textIndex));
                textIndex++;
            }
        }
    }

    private static void AddLinkRegions(PdfReadPage page, PdfPageInfo info, double pageHeight, List<PdfPageInteractionRegion> regions) {
        for (int i = 0; i < info.LinkAnnotations.Count; i++) {
            PdfLinkAnnotation link = info.LinkAnnotations[i];
            if (!TryFromUserRectangle(page, link.X1, link.Y1, link.X2, link.Y2, pageHeight, out PdfSelectionQuad? quad)) {
                continue;
            }
            regions.Add(new PdfPageInteractionRegion(
                PdfInteractionKind.Link,
                quad!,
                target: GetLinkTarget(link),
                subtype: "Link"));
        }
    }

    private static void AddAnnotationRegions(PdfReadPage page, PdfPageInfo info, double pageHeight, List<PdfPageInteractionRegion> regions) {
        for (int i = 0; i < info.Annotations.Count; i++) {
            PdfAnnotation annotation = info.Annotations[i];
            if (annotation.Subtype == "Link" || annotation.Subtype == "Widget") {
                continue;
            }

            if (!TryFromUserRectangle(page, annotation.X1, annotation.Y1, annotation.X2, annotation.Y2, pageHeight, out PdfSelectionQuad? quad)) {
                continue;
            }

            regions.Add(new PdfPageInteractionRegion(
                PdfInteractionKind.Annotation,
                quad!,
                text: annotation.Contents,
                subtype: annotation.Subtype,
                objectNumber: annotation.ObjectNumber));
        }
    }

    private static void AddFormWidgetRegions(PdfReadPage page, PdfPageInfo info, double pageHeight, List<PdfPageInteractionRegion> regions) {
        for (int i = 0; i < info.FormWidgets.Count; i++) {
            PdfFormWidget widget = info.FormWidgets[i];
            if (!TryFromUserRectangle(page, widget.X1, widget.Y1, widget.X2, widget.Y2, pageHeight, out PdfSelectionQuad? quad)) {
                continue;
            }
            regions.Add(new PdfPageInteractionRegion(
                PdfInteractionKind.FormWidget,
                quad!,
                fieldName: widget.FieldName,
                objectNumber: widget.ObjectNumber,
                subtype: "Widget"));
        }
    }

    private static bool TryFromUserRectangle(
        PdfReadPage page,
        double x1,
        double y1,
        double x2,
        double y2,
        double pageHeight,
        out PdfSelectionQuad? clipped) {
        (double X, double Y) topLeft = page.TransformPointToVisual(x1, y2);
        (double X, double Y) topRight = page.TransformPointToVisual(x2, y2);
        (double X, double Y) bottomRight = page.TransformPointToVisual(x2, y1);
        (double X, double Y) bottomLeft = page.TransformPointToVisual(x1, y1);
        var quad = new PdfSelectionQuad(
            ToTopLeft(topLeft, pageHeight),
            ToTopLeft(topRight, pageHeight),
            ToTopLeft(bottomRight, pageHeight),
            ToTopLeft(bottomLeft, pageHeight));
        (double Width, double Height) pageSize = page.GetInteractionPageSize();
        double left = Math.Max(0D, quad.Left);
        double top = Math.Max(0D, quad.Top);
        double right = Math.Min(pageSize.Width, quad.Right);
        double bottom = Math.Min(pageSize.Height, quad.Bottom);
        if (right <= left || bottom <= top) {
            clipped = null;
            return false;
        }

        clipped = new PdfSelectionQuad(
            new PdfSelectionPoint(left, top),
            new PdfSelectionPoint(right, top),
            new PdfSelectionPoint(right, bottom),
            new PdfSelectionPoint(left, bottom));
        return true;
    }

    private static PdfSelectionQuad FromVisualBaseline(
        double startX,
        double startY,
        double endX,
        double endY,
        double normalX,
        double normalY,
        double ascent,
        double descent,
        double pageHeight) {
        return new PdfSelectionQuad(
            ToTopLeft((startX + normalX * ascent, startY + normalY * ascent), pageHeight),
            ToTopLeft((endX + normalX * ascent, endY + normalY * ascent), pageHeight),
            ToTopLeft((endX - normalX * descent, endY - normalY * descent), pageHeight),
            ToTopLeft((startX - normalX * descent, startY - normalY * descent), pageHeight));
    }

    private static PdfSelectionPoint ToTopLeft((double X, double Y) point, double pageHeight) =>
        new PdfSelectionPoint(point.X, pageHeight - point.Y);

    private static string? GetLinkTarget(PdfLinkAnnotation link) {
        if (link.Uri is not null) return link.Uri;
        if (link.DestinationName is not null) return link.DestinationName;
        if (link.NamedAction is not null) return link.NamedAction;
        if (link.RemoteFile is not null) return link.RemoteDestinationName is null
            ? link.RemoteFile
            : link.RemoteFile + "#" + link.RemoteDestinationName;
        return link.DestinationPageNumber.HasValue
            ? "page:" + link.DestinationPageNumber.Value.ToString(CultureInfo.InvariantCulture)
            : null;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
