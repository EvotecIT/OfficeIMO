using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Creates exact review and delivery pixels from the canonical sheet-placement plan.</summary>
public static class PdfPrintRenderer {
    /// <summary>Prepares raster sheets from an authenticated snapshot without reopening the source location.</summary>
    public static PdfPreparedPrintDocument Prepare(PdfDocument document, PdfPrintPlanRequest request,
        PdfPrintRenderOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(request);
        options ??= new();
        double dpi = options.Dpi;
        int maximumPages = options.MaximumPages;
        long maximumPixels = options.MaximumPixelsPerImage, maximumOutputBytes = options.MaximumOutputBytes;
        if (!double.IsFinite(dpi) || dpi < 72 || dpi > 600) throw new ArgumentOutOfRangeException(nameof(options.Dpi));
        if (maximumPages <= 0 || maximumPixels <= 0 || maximumOutputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options));
        PdfPrintPlan plan = PdfPrintPlanner.Create(document, request, cancellationToken);
        if (plan.SelectedPages.Count > maximumPages) throw new InvalidOperationException($"Print preparation is limited to {maximumPages} selected pages.");
        var sheets = new List<PdfRenderedPrintSheet>();
        var diagnostics = new List<string>();
        long retainedBytes = 0;
        foreach (PdfPrintSheet sheet in plan.Sheets) {
            cancellationToken.ThrowIfCancellationRequested();
            var drawing = new OfficeDrawing(sheet.PaperSize.Width, sheet.PaperSize.Height);
            foreach (PdfPrintPlacement placement in sheet.Placements) {
                PdfPageRenderResult page = document.Render.PrintPage(placement.PageNumber,
                    new PdfPagePrintOptions { Dpi = dpi, MaximumPixels = maximumPixels, MaximumOutputBytes = maximumOutputBytes }, cancellationToken);
                byte[] png = page.Bytes ?? throw new InvalidOperationException("The source page did not produce printable pixels.");
                diagnostics.AddRange(page.Diagnostics);
                drawing.AddClippedImage(png, "image/png",
                    new OfficeImageProjection(new OfficeImagePlacement(placement.X, placement.Y, placement.Width, placement.Height)),
                    placement.SlotX, placement.SlotY, OfficeClipPath.Rectangle(placement.SlotWidth, placement.SlotHeight));
            }
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
                Scale = dpi / 72, Background = OfficeColor.White, MaximumRasterPixels = maximumPixels, CancellationToken = cancellationToken
            });
            byte[] encoded = OfficePngWriter.Encode(raster, cancellationToken);
            retainedBytes = checked(retainedBytes + encoded.LongLength);
            if (retainedBytes > maximumOutputBytes) throw new InvalidOperationException("Prepared print sheets exceed the configured output-byte limit.");
            sheets.Add(new PdfRenderedPrintSheet(sheet, encoded));
        }
        return new PdfPreparedPrintDocument(request.InputPath, plan, sheets, diagnostics);
    }
}
