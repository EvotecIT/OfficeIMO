using OfficeIMO.Pdf;

namespace OfficeIMO.Xps;

/// <summary>Thin PDF export over the native XPS reader and existing PDF drawing engine.</summary>
public static partial class XpsPdfExtensions {
    internal static PdfDocumentConversionResult ConvertToPdfWithEvidence(this XpsDocument document, XpsPdfExportOptions? options,
        CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        var snapshot = (options ?? new XpsPdfExportOptions()).Clone();
        if (snapshot.MaximumSvgCharacters < 1) throw new ArgumentOutOfRangeException(nameof(options), "MaximumSvgCharacters must be positive.");
        if (snapshot.MaximumPathCommands < 1) throw new ArgumentOutOfRangeException(nameof(options), "MaximumPathCommands must be positive.");
        var report = new PdfConversionReport();
        PdfCreateOptions pdfOptions = snapshot.PdfOptions?.Clone() ?? new PdfCreateOptions();
        pdfOptions.ReportDiagnosticsTo(report, "OfficeIMO.Xps.Pdf");
        snapshot.PdfOptions = pdfOptions;
        return new PdfDocumentConversionResult(CreatePdf(document, snapshot, cancellationToken), report);
    }

    private static PdfDocument CreatePdf(XpsDocument document, XpsPdfExportOptions options, CancellationToken cancellationToken) {
        if (document.Pages.Count == 0) throw new InvalidOperationException("PDF export requires at least one page.");
        var navigation = new XpsPdfNavigation(document, cancellationToken);
        var structure = options.PreserveLogicalStructure ? new XpsPdfLogicalStructure(document, cancellationToken) : null;
        var pdf = PdfDocument.CreateForImportedLayout(options.PdfOptions);
        if (structure?.HasNativeStructure == true) pdf.TaggedStructure(PdfTaggedStructureMode.CatalogMarkers);
        long projectedSvgCharacters = 0, projectedPathCommands = 0;
        pdf.Compose(builder => {
            for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
                var page = document.Pages[pageIndex];
                cancellationToken.ThrowIfCancellationRequested();
                var projection = new XpsSvgConverter(page, cancellationToken, explicitPageLinks: true, nativeDrawing: true).Convert(false);
                if (projection.Svg.Length > options.MaximumSvgCharacters - projectedSvgCharacters)
                    throw new InvalidDataException("XPS PDF projection exceeds MaximumSvgCharacters.");
                if (projection.PathCommands > options.MaximumPathCommands - projectedPathCommands)
                    throw new InvalidDataException("XPS PDF projection exceeds MaximumPathCommands.");
                projectedSvgCharacters += projection.Svg.Length;
                projectedPathCommands += projection.PathCommands;
                var drawing = XpsPage.ImportDrawing(navigation.MapLinks(projection.Svg, pageIndex, cancellationToken), cancellationToken, structure?.HasNativeStructure == true);
                double width = page.Width * 72D / 96D, height = page.Height * 72D / 96D;
                builder.Page(p => p.Size(global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height)).Margin(new PdfPageMargins(OfficeIMO.OfficeLength.Points(0),OfficeIMO.OfficeLength.Points(0),OfficeIMO.OfficeLength.Points(0),OfficeIMO.OfficeLength.Points(0)))
                    .Canvas(c => {
                        navigation.AddDestinations(c, projection, pageIndex, width, height, cancellationToken);
                        if (drawing.Elements.Count > 0) {
                            if (structure?.HasNativeStructure == true) c.SourceStructuredDrawing(drawing, width, height, structure.Paint(pageIndex));
                            else c.Drawing(drawing, global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(0), global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height));
                        }
                        if (structure?.HasNativeStructure == true) structure.AddText(c, projection.TextSpans, pageIndex);
                        else foreach (var span in projection.TextSpans) {
                            cancellationToken.ThrowIfCancellationRequested();
                            c.SearchableText(span.Text, XpsPdfLogicalStructure.Quad(span));
                        }
                    }));
            }
        });
        return pdf;
    }
}
