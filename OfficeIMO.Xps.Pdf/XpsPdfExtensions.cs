using OfficeIMO.Pdf;

namespace OfficeIMO.Xps;

/// <summary>Thin PDF export over the native XPS reader and existing PDF drawing engine.</summary>
public static class XpsPdfExtensions {
    /// <summary>Produces vector PDF pages. Glyphs remain positioned outlines, not searchable PDF text.</summary>
    public static byte[] ToPdf(this XpsDocument document, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (document.Pages.Count == 0) throw new InvalidOperationException("PDF export requires at least one page.");
        var pdf = PdfDocument.Create();
        pdf.Compose(builder => {
            foreach (var page in document.Pages) {
                cancellationToken.ThrowIfCancellationRequested();
                var drawing = page.ToDrawing(cancellationToken);
                double width = page.Width * 72D / 96D, height = page.Height * 72D / 96D;
                builder.Page(p => p.Size(width, height).Margin(new PageMargins(0, 0, 0, 0))
                    .Canvas(c => c.Drawing(drawing, 0, 0, width, height)));
            }
        });
        using var output = new MemoryStream();
        pdf.Save(output, cancellationToken);
        return output.ToArray();
    }
}
