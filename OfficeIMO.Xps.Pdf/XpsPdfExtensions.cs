using OfficeIMO.Pdf;

namespace OfficeIMO.Xps;

/// <summary>Thin PDF export over the native XPS reader and existing PDF drawing engine.</summary>
public static class XpsPdfExtensions {
    /// <summary>Produces vector PDF pages. Glyph outlines retain their appearance; native Unicode clusters provide searchable text in markup order.</summary>
    public static byte[] ToPdf(this XpsDocument document, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (document.Pages.Count == 0) throw new InvalidOperationException("PDF export requires at least one page.");
        var pdf = PdfDocument.Create();
        pdf.Compose(builder => {
            foreach (var page in document.Pages) {
                cancellationToken.ThrowIfCancellationRequested();
                var projection = page.ToSvg(false, cancellationToken);
                var drawing = XpsPage.ImportDrawing(projection, cancellationToken);
                double width = page.Width * 72D / 96D, height = page.Height * 72D / 96D;
                builder.Page(p => p.Size(width, height).Margin(new PageMargins(0, 0, 0, 0))
                    .Canvas(c => {
                        if (drawing.Elements.Count > 0) c.Drawing(drawing, 0, 0, width, height);
                        foreach (var span in projection.TextSpans) {
                            cancellationToken.ThrowIfCancellationRequested();
                            c.SearchableText(span.Text, new PdfSelectionQuad(Point(span.TopLeft), Point(span.TopRight), Point(span.BottomRight), Point(span.BottomLeft)));
                        }
                    }));
            }
        });
        using var output = new MemoryStream();
        pdf.Save(output, cancellationToken);
        return output.ToArray();
    }
    private static PdfSelectionPoint Point(OfficeIMO.Drawing.OfficePoint point) => new(point.X * 0.75, point.Y * 0.75);
}
