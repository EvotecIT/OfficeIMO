using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Builds the single PDF that preview and every transport share.</summary>
internal static class PdfPrintJobComposer {
    internal static PdfPrintJob Compose(PdfDocument document, PdfPrintPlanRequest request, PdfPrintJobOptions options, CancellationToken token) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(options);
        int maximumSheets = options.MaximumSheets;
        long maximumBytes = options.MaximumContentBytes, maximumPixels = options.MaximumPixelsPerImage;
        double rasterDpi = options.RasterContentDpi;
        if (maximumSheets <= 0) throw new ArgumentOutOfRangeException(nameof(options), "MaximumSheets must be positive.");
        if (maximumBytes <= 0 || maximumBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(options), "MaximumContentBytes must be positive and below 2 GiB.");
        if (!OfficeRuntimeCompatibility.IsFinite(rasterDpi) || rasterDpi < 72 || rasterDpi > 600) throw new ArgumentOutOfRangeException(nameof(options), "RasterContentDpi must be between 72 and 600.");
        // The shared decoder bounds every raster that preview or delivery decodes.
        _ = new OfficeRasterDecodeOptions { MaximumDecodedPixels = maximumPixels };
        PdfPrintPlan plan = PdfPrintPlanner.Create(document, request, token);
        if (plan.Sheets.Count > maximumSheets) throw new InvalidOperationException($"Print jobs are limited to {maximumSheets} sheets.");
        token.ThrowIfCancellationRequested();

        var diagnostics = new List<string>();
        byte[] source = document.GetBytesForOperation(token);
        PdfLoadOptions readOptions = document.ReadOptions;
        string? blocker = FindVectorBlocker(ref source, ref readOptions, document, plan, token);
        if (blocker is null) {
            try {
                byte[] vector = ComposeVector(source, readOptions, plan, maximumBytes, token);
                return new PdfPrintJob(request.InputPath, plan, vector, PdfPrintContentKind.Vector, diagnostics, maximumPixels);
            } catch (Exception error) when (error is NotSupportedException or PdfPermissionDeniedException) {
                blocker = error.Message;
            }
        }
        diagnostics.Add("print.raster-content: " + blocker + " Sheets are printed as raster images.");
        byte[] raster = ComposeRaster(document, plan, rasterDpi, maximumPixels, maximumBytes, diagnostics, token);
        return new PdfPrintJob(request.InputPath, plan, raster, PdfPrintContentKind.Raster, diagnostics, maximumPixels);
    }

    private static string? FindVectorBlocker(ref byte[] source, ref PdfLoadOptions readOptions, PdfDocument document, PdfPrintPlan plan, CancellationToken token) {
        PdfDocumentPreflight preflight = PdfInspector.Preflight(source, readOptions);
        PdfDocumentSecurityInfo security = preflight.Probe.Security;
        if (!PdfPermissionAuthorization.CanMutate(security, preflight.PermissionPolicy, PdfMutationOperation.ExtractPages))
            return "The source does not permit page import.";
        if (security.HasEncryption && !security.HasOwnerAuthorization && preflight.PermissionPolicy != PdfPermissionPolicy.IgnoreRestrictions &&
            security.AllowsHighQualityPrinting != true)
            return "The source permits only low-quality printing.";
        token.ThrowIfCancellationRequested();
        PdfDocumentInfo info = PdfInspector.Inspect(source, readOptions);
        if (info.HasOptionalContent) return "Vector sheets cannot preserve the source's optional-content visibility.";
        if (!HasPrintedAnnotations(info, plan.SelectedPages)) return null;
        // Imported pages carry content streams only, so printed annotation and form appearances are flattened first.
        PdfDocument flattened;
        try {
            flattened = document.FlattenInteractiveContent().ToDocument();
        } catch (Exception error) when (error is NotSupportedException or InvalidOperationException or PdfPermissionDeniedException) {
            return "Printed annotations could not be flattened: " + error.Message;
        }
        token.ThrowIfCancellationRequested();
        byte[] flattenedBytes = flattened.GetBytesForOperation(token);
        PdfLoadOptions flattenedOptions = flattened.ReadOptions;
        if (HasPrintedAnnotations(PdfInspector.Inspect(flattenedBytes, flattenedOptions), plan.SelectedPages))
            return "Some printed annotations have no flattenable appearance.";
        source = flattenedBytes;
        readOptions = flattenedOptions;
        return null;
    }

    internal static bool HasPrintedAnnotations(PdfDocumentInfo info, IReadOnlyList<int> pages) {
        foreach (int page in pages.Distinct()) {
            foreach (PdfAnnotation annotation in info.Pages[page - 1].Annotations) {
                if (annotation.Subtype is "Link" or "Popup" || !annotation.HasNormalAppearance) continue;
                int flags = annotation.Flags ?? 0;
                if ((flags & 4) != 0 && (flags & 2) == 0) return true; // Print set, Hidden clear.
            }
        }
        return false;
    }

    private static byte[] ComposeVector(byte[] source, PdfLoadOptions readOptions, PdfPrintPlan plan, long maximumBytes, CancellationToken token) {
        PdfDocument blank = PdfDocument.Create(builder => {
            foreach (PdfPrintSheet sheet in plan.Sheets)
                builder.Page(page => page.Size(sheet.PaperSize.Width, sheet.PaperSize.Height).Margin(OfficeLength.Zero));
        });
        byte[] sheets = blank.ToBytes();
        token.ThrowIfCancellationRequested();
        var overlays = new List<PdfPageOverlayOptions>();
        foreach (PdfPrintSheet sheet in plan.Sheets) {
            double paperHeight = sheet.PaperSize.Height.ToPoints();
            OfficePageSelector target = OfficePageSelector.Parse(sheet.SheetNumber.ToString(CultureInfo.InvariantCulture));
            foreach (PdfPrintPlacement placement in sheet.Placements) {
                // Plan coordinates have a top-left origin; overlays use the PDF bottom-left frame.
                double slotBottom = paperHeight - placement.SlotY - placement.SlotHeight;
                overlays.Add(new PdfPageOverlayOptions {
                    SourcePageNumber = placement.PageNumber,
                    SourceReadOptions = readOptions,
                    TargetPages = target,
                    X = placement.X,
                    Y = paperHeight - placement.Y - placement.Height,
                    Width = placement.Width,
                    Height = placement.Height,
                    Fit = PdfPageOverlayFit.Stretch,
                    ClipRectangle = new PdfPageRectangle(placement.SlotX, slotBottom, placement.SlotX + placement.SlotWidth, slotBottom + placement.SlotHeight),
                    CancellationToken = token
                });
            }
        }
        byte[] output = PdfStamper.StampPages(sheets, source, overlays, maximumBytes);
        if (output.LongLength > maximumBytes) throw new InvalidOperationException("The print job exceeds its content byte limit.");
        return output;
    }

    private static byte[] ComposeRaster(PdfDocument document, PdfPrintPlan plan, double dpi, long maximumPixels, long maximumBytes,
        List<string> diagnostics, CancellationToken token) =>
        WrapRasterSheets(plan.Sheets, sheet => {
            var drawing = new OfficeDrawing(sheet.PaperSize.Width.ToPoints(), sheet.PaperSize.Height.ToPoints());
            foreach (PdfPrintPlacement placement in sheet.Placements) {
                PdfPageRenderResult page = document.Render.PrintPage(placement.PageNumber,
                    new PdfPagePrintOptions { Dpi = dpi, PageScale = placement.Scale, MaximumPixels = maximumPixels, MaximumOutputBytes = maximumBytes }, token);
                byte[] png = page.Bytes ?? throw new InvalidOperationException("The source page did not produce printable pixels.");
                diagnostics.AddRange(page.Diagnostics);
                drawing.AddClippedImage(png, "image/png",
                    new OfficeImageProjection(new OfficeImagePlacement(placement.X, placement.Y, placement.Width, placement.Height)),
                    placement.SlotX, placement.SlotY, OfficeClipPath.Rectangle(placement.SlotWidth, placement.SlotHeight));
            }
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
                Scale = dpi / 72, Background = OfficeColor.White, MaximumRasterPixels = maximumPixels,
                ThrowOnImageDecodeFailure = true, CancellationToken = token
            });
            return OfficePngWriter.Encode(raster, token);
        }, maximumBytes, token);

    /// <summary>Wraps one PNG per sheet into a PDF page at that sheet's paper size, rendering one sheet at a time.</summary>
    internal static byte[] WrapRasterSheets(IReadOnlyList<PdfPrintSheet> sheets, Func<PdfPrintSheet, byte[]> renderPng, long maximumBytes, CancellationToken token) {
        var pages = new List<PdfDocument>(sheets.Count);
        long retained = 0;
        foreach (PdfPrintSheet sheet in sheets) {
            token.ThrowIfCancellationRequested();
            byte[] encoded = renderPng(sheet);
            retained = checked(retained + encoded.LongLength);
            if (retained > maximumBytes) throw new InvalidOperationException("The print job exceeds its content byte limit.");
            pages.Add(PdfDocument.CreateFromImages([new PdfImageDocumentSource(encoded)],
                new PdfImageImportOptions { FixedPageSize = sheet.PaperSize, Margin = OfficeLength.Zero, Fit = OfficeImageFit.Stretch }, token));
        }
        byte[] bytes = (pages.Count == 1 ? pages[0] : PdfDocument.Merge(pages, cancellationToken: token)).ToBytes(new PdfSaveOptions { AllowLoss = true }, token);
        if (bytes.LongLength > maximumBytes) throw new InvalidOperationException("The print job exceeds its content byte limit.");
        return bytes;
    }
}
