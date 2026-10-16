using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.PowerPoint;

/// <summary>Selection and encoding over the canonical slide PDF layout; no independent slide layout.</summary>
internal static class PowerPointPageImageExport {
    internal static OfficeImageExportResult Render(PowerPointSlide slide, OfficeImageExportFormat format,
        PowerPointImageExportOptions options, CancellationToken cancellationToken) {
        PowerPointPresentation presentation = GetPresentation(slide);
        int number = GetSlideNumber(presentation, slide);
        var batch = options.CopyPowerPointOptionsTo(new PowerPointPresentationImageExportOptions());
        batch.IncludeHiddenSlides = true;
        batch.SlideNumbers = new[] { number };
        OfficeImageExportResult? result = null;
        RenderPages(presentation, format, batch, value => result = value, cancellationToken);
        return result!;
    }

    internal static void RenderPages(PowerPointPresentation presentation, OfficeImageExportFormat format,
        PowerPointPresentationImageExportOptions options, OfficeImageExportConsumer consumer, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var slideNumbers = Enumerable.Range(1, presentation.Slides.Count)
            .Where(number => options.IncludeHiddenSlides || !presentation.Slides[number - 1].Hidden).ToArray();
        if (options.SlideNumbers?.Any(number => number < 1 || number > presentation.Slides.Count) == true)
            throw new ArgumentOutOfRangeException(nameof(options.SlideNumbers), "Selected slide numbers must exist in the presentation.");
        var selected = options.SlideNumbers == null || options.SlideNumbers.Count == 0
            ? null : new HashSet<int>(options.SlideNumbers);
        int[] selectedNumbers = slideNumbers.Where(number => selected == null || selected.Contains(number)).ToArray();
        if (selectedNumbers.Length == 0) return;
        int outputIndex = 0;
        PdfCore.PdfDocumentConversionResult? conversion = null;
        PdfCore.PdfImageExportEngine.ExportEach(
            token => {
                PowerPointToPdfOptions layout = CreateLayout(options, options.IncludeHiddenSlides);
                layout.SelectedSlides = new HashSet<PowerPointSlide>(selectedNumbers.Select(number => presentation.Slides[number - 1]));
                conversion = presentation.ToPdfDocumentResult(layout, token);
                return PdfCore.PdfReadDocument.Open(conversion.ToBytes(token), options: null, cancellationToken: token);
            }, format, options.ToPdfImageOptions(), _ => null,
            result => {
                // The PDF exporter emits in page order, which is the presentation's original slide order.
                int number = selectedNumbers[outputIndex++];
                consumer(new OfficeImageExportResult(result.Format, result.Width, result.Height, result.Bytes,
                    "Slide " + number, "PowerPoint slide " + number, result.Diagnostics));
            }, initialDiagnostics: null,
            () => PdfCore.PdfImageExportEngine.MapConversionDiagnostics(conversion!), cancellationToken);
    }

    internal static PowerPointSlideVisualSnapshot CreateSnapshot(PowerPointSlide slide, PowerPointImageExportOptions options) {
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(options.RenderTimeout, CancellationToken.None);
        PowerPointPresentation presentation = GetPresentation(slide);
        int number = GetSlideNumber(presentation, slide);
        PowerPointToPdfOptions layout = CreateLayout(options, true);
        layout.SelectedSlides = new HashSet<PowerPointSlide> { slide };
        PdfCore.PdfDocumentConversionResult conversion = presentation.ToPdfDocumentResult(layout, execution.Token);
        PdfCore.PdfReadDocument document = PdfCore.PdfReadDocument.Open(conversion.ToBytes(execution.Token), options: null, cancellationToken: execution.Token);
        PdfCore.PdfReadPage page = document.Pages[0];
        var size = page.GetPageSize();
        OfficeDrawing drawing = page.ToDrawing(execution.Token, scene => scene.ApplyImageExportOptions(options),
            options.GetEffectiveScale(size.Width, size.Height), options.MaximumRasterPixels);
        var diagnostics = new List<OfficeImageExportDiagnostic>(PdfCore.PdfImageExportEngine.MapConversionDiagnostics(conversion));
        foreach (PdfCore.PdfRenderCapabilityDiagnostic diagnostic in page.GetRenderCapabilityDiagnostics(cancellationToken: execution.Token))
            diagnostics.Add(new OfficeImageExportDiagnostic(
                diagnostic.SupportLevel == PdfCore.PdfRenderSupportLevel.Supported ? OfficeImageExportDiagnosticSeverity.Info : OfficeImageExportDiagnosticSeverity.Warning,
                diagnostic.Code, diagnostic.Message, "PowerPoint slide " + number, diagnostic.LossKind));
        drawing.AppendFontDiagnostics(diagnostics, "PowerPoint slide " + number);
        return new PowerPointSlideVisualSnapshot(drawing, diagnostics.AsReadOnly(), page.GetTextSpans(includeArtifactText: true, execution.Token));
    }

    private static PowerPointToPdfOptions CreateLayout(PowerPointImageExportOptions options, bool includeHiddenSlides) {
        PowerPointToPdfOptions layout = options.LayoutOptions.Clone();
        layout.PageLayout = PowerPointPdfPageLayout.Slides;
        layout.IncludeHiddenSlides = includeHiddenSlides;
        layout.IncludeSlideBackgrounds &= options.IncludeSlideBackground;
        layout.IncludePictures &= options.IncludeSlideContent && options.IncludePictures;
        layout.IncludeAutoShapes &= options.IncludeSlideContent && options.IncludeAutoShapes;
        layout.IncludeTextBoxes &= options.IncludeSlideContent && options.IncludeTextBoxes;
        layout.IncludeTables &= options.IncludeSlideContent && options.IncludeTables;
        layout.IncludeCharts &= options.IncludeSlideContent && options.IncludeCharts;
        layout.IncludeSmartArt &= options.IncludeSlideContent && options.IncludeSmartArt;
        layout.IncludeHiddenShapes = options.IncludeHiddenShapes;
        layout.MaxGroupShapeDepth = options.MaxGroupShapeDepth;
        layout.MaximumEmbeddedImageBytes = options.MaximumEmbeddedImageBytes;
        layout.MaximumTotalEmbeddedImageBytes = options.MaximumTotalEmbeddedImageBytes;
        if (options.Fonts.Faces.Count > 0 || options.TextShapingProvider != null || options.TextShapingLanguage != null || options.ImageCodec != null)
            layout.UseRenderingProfile(new OfficeRenderingProfile("PowerPoint image export", options.Fonts,
                options.TextShapingProvider, options.TextShapingLanguage, options.ImageCodec), OfficeRenderingProfileApplyMode.Overlay);
        return layout;
    }

    private static PowerPointPresentation GetPresentation(PowerPointSlide slide) => slide.OwnerPresentation
        ?? throw new InvalidOperationException("The slide is not attached to a presentation.");

    private static int GetSlideNumber(PowerPointPresentation presentation, PowerPointSlide slide) {
        for (int index = 0; index < presentation.Slides.Count; index++)
            if (ReferenceEquals(presentation.Slides[index], slide)) return index + 1;
        throw new InvalidOperationException("The slide no longer belongs to its presentation.");
    }
}
