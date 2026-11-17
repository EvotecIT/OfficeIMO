namespace OfficeIMO.Pdf;

public sealed partial class PdfPageInteractionMap {
    // Text-markup annotations commonly cover an OCR text layer, so painted-invisible OCR text counts as covered text.
    private static readonly PdfPageInteractionOptions MarkupTextOptions = new PdfPageInteractionOptions { IncludeTextRenderingMode3 = true };

    /// <summary>Builds the selectable text elements of a page, in extraction order, for text-markup coverage.</summary>
    internal static IReadOnlyList<PdfPageInteractionRegion> CreateMarkupTextRegions(PdfReadPage page) {
        (double Width, double Height) size = page.GetInteractionPageSize();
        var regions = new List<PdfPageInteractionRegion>();
        AddTextRegions(page, size.Width, size.Height, MarkupTextOptions, regions, static _ => null);
        return regions.AsReadOnly();
    }

    /// <summary>Joins the text elements whose centers lie inside the annotation's /QuadPoints, or inside /Rect when
    /// the annotation has no valid quad points, using the same separators as <see cref="GetSelectedText"/>.</summary>
    internal static string GetMarkupText(PdfReadPage page, IReadOnlyList<PdfPageInteractionRegion> textRegions, PdfAnnotation annotation) {
        List<PdfSelectionQuad> quads = GetMarkupQuads(page, annotation);
        if (quads.Count == 0) return string.Empty;
        var selected = new List<PdfPageInteractionRegion>();
        for (int i = 0; i < textRegions.Count; i++) {
            PdfSelectionQuad glyph = textRegions[i].Quad;
            double x = (glyph.TopLeft.X + glyph.TopRight.X + glyph.BottomRight.X + glyph.BottomLeft.X) / 4D;
            double y = (glyph.TopLeft.Y + glyph.TopRight.Y + glyph.BottomRight.Y + glyph.BottomLeft.Y) / 4D;
            for (int q = 0; q < quads.Count; q++) {
                if (quads[q].Contains(x, y)) {
                    selected.Add(textRegions[i]);
                    break;
                }
            }
        }

        return JoinSelectedText(selected).Trim();
    }

    private static List<PdfSelectionQuad> GetMarkupQuads(PdfReadPage page, PdfAnnotation annotation) {
        double pageHeight = page.GetInteractionPageSize().Height;
        IReadOnlyList<double> points = annotation.QuadPoints;
        var quads = new List<PdfSelectionQuad>(Math.Max(1, points.Count / 8));
        for (int offset = 0; offset + 8 <= points.Count; offset += 8) {
            var corners = new PdfSelectionPoint[4];
            bool finite = true;
            for (int corner = 0; corner < 4; corner++) {
                double x = points[offset + corner * 2];
                double y = points[offset + corner * 2 + 1];
                if (!IsFinite(x) || !IsFinite(y)) { finite = false; break; }
                corners[corner] = ToTopLeft(page.TransformPointToVisual(x, y), pageHeight);
            }
            if (finite) quads.Add(ToConvexQuad(corners));
        }

        if (quads.Count == 0 && annotation.Width > 0D && annotation.Height > 0D &&
            IsFinite(annotation.X1) && IsFinite(annotation.Y1) && IsFinite(annotation.X2) && IsFinite(annotation.Y2)) {
            quads.Add(ToConvexQuad(new[] {
                ToTopLeft(page.TransformPointToVisual(annotation.X1, annotation.Y2), pageHeight),
                ToTopLeft(page.TransformPointToVisual(annotation.X2, annotation.Y2), pageHeight),
                ToTopLeft(page.TransformPointToVisual(annotation.X2, annotation.Y1), pageHeight),
                ToTopLeft(page.TransformPointToVisual(annotation.X1, annotation.Y1), pageHeight)
            }));
        }

        return quads;
    }

    // Producers disagree on /QuadPoints corner order: the specification describes a counterclockwise quadrilateral,
    // while common viewers write top-left, top-right, bottom-left, bottom-right. Order the corners around their
    // centroid so containment sees one convex outline either way.
    private static PdfSelectionQuad ToConvexQuad(PdfSelectionPoint[] corners) {
        double centerX = (corners[0].X + corners[1].X + corners[2].X + corners[3].X) / 4D;
        double centerY = (corners[0].Y + corners[1].Y + corners[2].Y + corners[3].Y) / 4D;
        PdfSelectionPoint[] ordered = corners
            .OrderBy(point => Math.Atan2(point.Y - centerY, point.X - centerX))
            .ToArray();
        return new PdfSelectionQuad(ordered[0], ordered[1], ordered[2], ordered[3]);
    }
}
