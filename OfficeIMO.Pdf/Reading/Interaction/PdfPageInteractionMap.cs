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

    /// <summary>All text, image, link, annotation, and form-widget regions.</summary>
    public IReadOnlyList<PdfPageInteractionRegion> Regions { get; }

    /// <summary>Text-element regions in content extraction order.</summary>
    public IReadOnlyList<PdfPageInteractionRegion> TextRegions { get; }

    /// <summary>Builds an interaction map from PDF bytes.</summary>
    public static PdfPageInteractionMap Create(
        byte[] pdf,
        int pageNumber,
        PdfPageInteractionOptions? options = null,
        PdfLoadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), "Page number must be positive.");
        }

        PdfPageInteractionOptions effective = options ?? new PdfPageInteractionOptions();
        if (effective.MaxTextRegions <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "Maximum text regions must be positive.");
        }
        if (effective.MaxImageRegions <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "Maximum image regions must be positive.");
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
        AddImageRegions(document, pdf, page, pageNumber, size.Width, size.Height, effective, regions);
        AddLinkRegions(page, pageInfo, size.Height, regions);
        AddAnnotationRegions(page, pageInfo, size.Height, regions);
        AddFormWidgetRegions(page, pageInfo, size.Height, regions);
        return new PdfPageInteractionMap(pageNumber, size.Width, size.Height, regions.AsReadOnly());
    }

    private static void AddImageRegions(
        PdfReadDocument document,
        byte[] pdf,
        PdfReadPage page,
        int pageNumber,
        double pageWidth,
        double pageHeight,
        PdfPageInteractionOptions options,
        List<PdfPageInteractionRegion> regions) {
        IReadOnlyList<PdfImagePlacement> placements = PdfImageEditor.Placements(document, pdf, pageNumber);
        (double originX, double originY) = page.GetPageBoundaryOrigin();
        int emitted = 0;
        for (int i = 0; i < placements.Count; i++) {
            PdfImagePlacement placement = placements[i];
            PdfSelectionQuad quad = FromImagePlacement(page, placement, originX, originY, pageHeight);
            if (!TryApplyImageClip(page, placement, pageHeight, ref quad)) continue;
            if (!quad.Intersects(0D, 0D, pageWidth, pageHeight)) continue;
            if (emitted >= options.MaxImageRegions) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.InteractionRegions, options.MaxImageRegions, emitted + 1L);
            }
            regions.Add(new PdfPageInteractionRegion(
                PdfInteractionKind.Image,
                quad,
                subtype: "Image",
                objectNumber: placement.ObjectNumber == 0 ? null : placement.ObjectNumber,
                imagePlacement: placement));
            emitted++;
        }
    }

    internal static IReadOnlyList<PdfSelectionQuad> GetOcrOverlapTextSpanBounds(PdfReadPage page) {
        IReadOnlyList<PdfTextSpan> spans = page.GetInteractionTextSpans();
        (double pageWidth, double pageHeight) = page.GetInteractionPageSize();
        var bounds = new List<PdfSelectionQuad>(spans.Count);
        var geometryBudget = new PdfReadPage.VisualGeometryBudget();
        for (int spanIndex = 0; spanIndex < spans.Count; spanIndex++) {
            PdfTextSpan span = spans[spanIndex];
            // Rendering mode 3 is the standard invisible searchable-text layer. Include it
            // for OCR deduplication without allowing other concealed/clipped text to mask pixels.
            if (string.IsNullOrEmpty(span.Text) || (!span.IsVisible && span.TextRenderingMode != 3)) continue;

            TextElementEnumerator enumerator = StringInfo.GetTextElementEnumerator(span.Text);
            int elementCount = 0;
            while (enumerator.MoveNext()) elementCount++;
            if (elementCount == 0) continue;

            double totalAdvance = Math.Abs(span.Advance);
            if (totalAdvance <= 0D) {
                totalAdvance = elementCount * Math.Max(1D, span.FontSize * 0.5D);
            }
            double radians = span.RotationDegrees * Math.PI / 180D;
            double directionX = Math.Cos(radians);
            double directionY = Math.Sin(radians);
            double normalX = -directionY;
            double normalY = directionX;
            PdfSelectionQuad quad = FromVisualBaseline(
                span.X,
                span.Y,
                span.X + directionX * totalAdvance,
                span.Y + directionY * totalAdvance,
                normalX,
                normalY,
                Math.Max(1D, span.FontSize),
                Math.Max(0.5D, span.FontSize * 0.2D),
                pageHeight);
            if (span.ClipPath.HasValue) {
                PdfPageClipPath clip = span.ClipPath.Value;
                if (clip.Width <= 0D || clip.Height <= 0D ||
                    clip.CanProveNoPositiveAreaIntersection(
                        PdfPageClipPath.Rectangle(quad.Left, quad.Top, quad.Width, quad.Height),
                        geometryBudget)) {
                    continue;
                }
            }
            if (quad.Intersects(0D, 0D, pageWidth, pageHeight)) bounds.Add(quad);
        }
        return bounds.Count == 0 ? Array.Empty<PdfSelectionQuad>() : bounds.AsReadOnly();
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
        for (int spanIndex = 0; spanIndex < spans.Count; spanIndex++) {
            PdfTextSpan span = spans[spanIndex];
            if (string.IsNullOrEmpty(span.Text) || (!span.IsVisible && !options.IncludeInvisibleText)) {
                continue;
            }

            int[] elementStarts = StringInfo.ParseCombiningCharacters(span.Text);
            int elementCount = elementStarts.Length;
            if (elementCount == 0) {
                continue;
            }

            double[] elementBoundaries = GetTextElementAdvanceBoundaries(span, elementStarts);
            double radians = span.RotationDegrees * Math.PI / 180D;
            double directionX = Math.Cos(radians);
            double directionY = Math.Sin(radians);
            double normalX = -directionY;
            double normalY = directionX;
            double ascent = Math.Max(1D, span.FontSize);
            double descent = Math.Max(0.5D, span.FontSize * 0.2D);
            for (int elementIndex = 0; elementIndex < elementCount; elementIndex++) {
                int startIndex = elementStarts[elementIndex];
                int endIndex = elementIndex + 1 < elementCount ? elementStarts[elementIndex + 1] : span.Text.Length;
                string element = span.Text.Substring(startIndex, endIndex - startIndex);
                double startAdvance = elementBoundaries[elementIndex];
                double endAdvance = elementBoundaries[elementIndex + 1];
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
                if (textIndex >= options.MaxTextRegions) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.InteractionRegions, options.MaxTextRegions, textIndex + 1L);
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

    private static double[] GetTextElementAdvanceBoundaries(PdfTextSpan span, int[] elementStarts) {
        IReadOnlyList<double>? characterAdvances = span.CharacterAdvances;
        if (characterAdvances is not null && characterAdvances.Count == span.Text.Length) {
            var characterBoundaries = new double[span.Text.Length + 1];
            bool usable = true;
            double signedTotalAdvance = 0D;
            for (int characterIndex = 0; characterIndex < characterAdvances.Count; characterIndex++) {
                double advance = characterAdvances[characterIndex];
                if (!IsFinite(advance)) {
                    usable = false;
                    break;
                }
                signedTotalAdvance += advance;
                if (!IsFinite(signedTotalAdvance)) {
                    usable = false;
                    break;
                }
            }

            // RotationDegrees already follows the resolved origin-to-end direction. Character
            // advances retain their text-space sign, so mirrored text must be projected onto that
            // direction once rather than reversing the run for a second time.
            double directionSign = signedTotalAdvance < 0D ? -1D : 1D;
            for (int characterIndex = 0; usable && characterIndex < characterAdvances.Count; characterIndex++) {
                double advance = characterAdvances[characterIndex];
                characterBoundaries[characterIndex + 1] = characterBoundaries[characterIndex] + advance * directionSign;
                if (!IsFinite(characterBoundaries[characterIndex + 1])) usable = false;
            }

            if (usable && Math.Abs(characterBoundaries[characterBoundaries.Length - 1]) > double.Epsilon) {
                var elementBoundaries = new double[elementStarts.Length + 1];
                for (int elementIndex = 0; elementIndex < elementStarts.Length; elementIndex++) {
                    elementBoundaries[elementIndex] = characterBoundaries[elementStarts[elementIndex]];
                }
                elementBoundaries[elementBoundaries.Length - 1] = characterBoundaries[characterBoundaries.Length - 1];
                return elementBoundaries;
            }
        }

        double totalAdvance = Math.Abs(span.Advance);
        if (!IsFinite(totalAdvance) || totalAdvance <= 0D) {
            totalAdvance = elementStarts.Length * Math.Max(1D, span.FontSize * 0.5D);
        }

        var fallbackBoundaries = new double[elementStarts.Length + 1];
        double elementAdvance = totalAdvance / elementStarts.Length;
        for (int elementIndex = 1; elementIndex < fallbackBoundaries.Length; elementIndex++) {
            fallbackBoundaries[elementIndex] = elementIndex * elementAdvance;
        }
        return fallbackBoundaries;
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

    private static PdfSelectionQuad FromImagePlacement(
        PdfReadPage page,
        PdfImagePlacement placement,
        double originX,
        double originY,
        double pageHeight) {
        double e = placement.E + originX;
        double f = placement.F + originY;
        (double X, double Y) bottomLeft = page.TransformPointToVisual(e, f);
        (double X, double Y) bottomRight = page.TransformPointToVisual(placement.A + e, placement.B + f);
        (double X, double Y) topRight = page.TransformPointToVisual(
            placement.A + placement.C + e,
            placement.B + placement.D + f);
        (double X, double Y) topLeft = page.TransformPointToVisual(placement.C + e, placement.D + f);
        return new PdfSelectionQuad(
            ToTopLeft(topLeft, pageHeight),
            ToTopLeft(topRight, pageHeight),
            ToTopLeft(bottomRight, pageHeight),
            ToTopLeft(bottomLeft, pageHeight));
    }

    private static bool TryApplyImageClip(
        PdfReadPage page,
        PdfImagePlacement placement,
        double pageHeight,
        ref PdfSelectionQuad quad) {
        PdfImageClipInfo? clip = placement.Clip;
        if (clip is null) return true;
        if (clip.Width <= 0D || clip.Height <= 0D) return false;
        double sourceHeight = page.GetPageSize().Height;
        if (!TryFromUserRectangle(
                page,
                clip.X,
                sourceHeight - (clip.Y + clip.Height),
                clip.X + clip.Width,
                sourceHeight - clip.Y,
                pageHeight,
                out PdfSelectionQuad? clipQuad)) {
            return false;
        }

        const double tolerance = 0.001D;
        if (clipQuad!.Left <= quad.Left + tolerance &&
            clipQuad.Top <= quad.Top + tolerance &&
            clipQuad.Right >= quad.Right - tolerance &&
            clipQuad.Bottom >= quad.Bottom - tolerance) {
            return true;
        }

        // Interaction regions are quadrilateral, while PDF clips may be arbitrary paths.
        // Retain the editable image identity using the intersection of their conservative
        // visual bounds instead of dropping every nonrectangular or inexact clip.
        double left = Math.Max(quad.Left, clipQuad.Left);
        double top = Math.Max(quad.Top, clipQuad.Top);
        double right = Math.Min(quad.Right, clipQuad.Right);
        double bottom = Math.Min(quad.Bottom, clipQuad.Bottom);
        if (right <= left || bottom <= top) return false;
        quad = new PdfSelectionQuad(
            new PdfSelectionPoint(left, top),
            new PdfSelectionPoint(right, top),
            new PdfSelectionPoint(right, bottom),
            new PdfSelectionPoint(left, bottom));
        return true;
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
