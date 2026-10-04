using OfficeIMO.Pdf;

namespace OfficeIMO.Xps;

/// <summary>Thin PDF export over the native XPS reader and existing PDF drawing engine.</summary>
public static class XpsPdfExtensions {
    /// <summary>Produces vector PDF pages with native semantic structure, retaining the two-parameter conversion entrypoint.</summary>
    public static byte[] ToPdf(this XpsDocument document, CancellationToken cancellationToken) =>
        ToPdf(document, cancellationToken, preserveLogicalStructure: true);

    /// <summary>Produces vector PDF pages with searchable native Unicode. Authored native structure supplies semantic tags and logical reading order when present.</summary>
    /// <param name="document">The native document.</param>
    /// <param name="cancellationToken">Cancels conversion and PDF serialization.</param>
    /// <param name="preserveLogicalStructure">Maps native semantics strictly. False exports paint and searchable text in markup order, including documents with unsupported structure extensions.</param>
    /// <remarks>Page-local fragments use page order when no story addresses were authored. Native figure descriptions and PDF/UA conformance are not inferred.</remarks>
    public static byte[] ToPdf(this XpsDocument document, CancellationToken cancellationToken = default, bool preserveLogicalStructure = true) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (document.Pages.Count == 0) throw new InvalidOperationException("PDF export requires at least one page.");
        var navigation = new XpsPdfNavigation(document, cancellationToken);
        var structure = preserveLogicalStructure ? new XpsPdfLogicalStructure(document, cancellationToken) : null;
        var pdf = PdfDocument.Create();
        if (structure?.HasNativeStructure == true) pdf.TaggedStructure(PdfTaggedStructureMode.CatalogMarkers);
        pdf.Compose(builder => {
            for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
                var page = document.Pages[pageIndex];
                cancellationToken.ThrowIfCancellationRequested();
                var projection = new XpsSvgConverter(page, cancellationToken, explicitPageLinks: true).Convert(false);
                var drawing = XpsPage.ImportDrawing(navigation.MapLinks(projection.Svg, pageIndex, cancellationToken), cancellationToken, structure?.HasNativeStructure == true);
                double width = page.Width * 72D / 96D, height = page.Height * 72D / 96D;
                builder.Page(p => p.Size(width, height).Margin(new PageMargins(0, 0, 0, 0))
                    .Canvas(c => {
                        navigation.AddDestinations(c, projection, pageIndex, width, height, cancellationToken);
                        if (drawing.Elements.Count > 0) {
                            if (structure?.HasNativeStructure == true) c.SourceStructuredDrawing(drawing, width, height, structure.Paint(pageIndex));
                            else c.Drawing(drawing, 0, 0, width, height);
                        }
                        if (structure?.HasNativeStructure == true) structure.AddText(c, projection.TextSpans, pageIndex);
                        else foreach (var span in projection.TextSpans) {
                            cancellationToken.ThrowIfCancellationRequested();
                            c.SearchableText(span.Text, XpsPdfLogicalStructure.Quad(span));
                        }
                    }));
            }
        });
        using var output = new MemoryStream();
        pdf.Save(output, cancellationToken);
        return output.ToArray();
    }
}
