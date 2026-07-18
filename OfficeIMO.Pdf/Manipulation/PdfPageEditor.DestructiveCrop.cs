using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    /// <summary>
    /// Destructively crops selected pages by replacing the retained visual rectangle with an opaque raster page.
    /// Original selected-page content streams, resources, and annotations are not retained in the rewritten object graph.
    /// </summary>
    public static PdfDestructiveCropResult DestructiveCropPages(byte[] pdf, double left, double bottom, double right, double top, PdfDestructiveCropOptions? options = null, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf)); Guard.NotNull(pageNumbers, nameof(pageNumbers)); ValidatePageBoxCoordinates(left, bottom, right, top);
        PdfDestructiveCropOptions effective = options ?? new PdfDestructiveCropOptions();
        if (effective.Dpi <= 0D || double.IsNaN(effective.Dpi) || double.IsInfinity(effective.Dpi)) throw new ArgumentOutOfRangeException(nameof(options), "Destructive crop DPI must be positive and finite.");
        if (effective.MaxPixelsPerPage <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Maximum pixels per page must be positive.");
        PdfMutationPlan pageTreePlan = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageTree);
        PdfMutationPlan contentPlan = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageContent);
        PdfReadDocument source = PdfReadDocument.Open(pdf);
        int[] selected = pageNumbers.Length == 0 ? Enumerable.Range(1, source.Pages.Count).ToArray() : pageNumbers;
        ValidatePageNumbers(selected, source.Pages.Count, nameof(pageNumbers));
        EnsureNoSelectedFormWidgets(source, selected);

        byte[] translated = CropAndTranslatePages(pdf, left, bottom, right, top, selected);
        var renderOptions = new PdfPageRenderOptions { Format = PdfPageRenderFormat.Png, Dpi = effective.Dpi, Background = effective.Background, MaxPages = selected.Length, MaxPixelsPerPage = effective.MaxPixelsPerPage, ContinueOnError = false };
        IReadOnlyList<PdfPageRenderResult> renders = PdfPageImageRenderer.RenderPages(translated, PdfPageSelection.From(selected), renderOptions);
        ValidateDestructiveCropRenders(renders, effective);
        PdfReadDocument translatedDocument = PdfReadDocument.Open(translated);
        var renderByPage = renders.ToDictionary(static render => render.PageNumber);
        var selectedSet = new HashSet<int>(selected);

        byte[] output = PdfDocumentObjectGraphRewriter.Rewrite(translated, null, null, (objects, security) => {
            int nextObjectNumber = objects.Keys.Max() + 1;
            for (int pageIndex = 0; pageIndex < translatedDocument.Pages.Count; pageIndex++) {
                int pageNumber = pageIndex + 1;
                if (!selectedSet.Contains(pageNumber)) continue;
                PdfReadPage readPage = translatedDocument.Pages[pageIndex];
                if (!objects.TryGetValue(readPage.ObjectNumber, out PdfIndirectObject? pageObject) || pageObject.Value is not PdfDictionary page) throw new InvalidOperationException("Selected PDF page object was not readable.");
                PdfPageRenderResult render = renderByPage[pageNumber];
                byte[] png = render.Bytes!;
                var info = new OfficeImageInfo(OfficeImageFormat.Png, render.Width, render.Height);
                if (!PdfWriter.TryBuildImageStream(png, info, right - left, top - bottom, out PdfWriter.PdfImageStream image, out string? reason)) throw new NotSupportedException(reason ?? "Rendered crop PNG could not be embedded.");
                int imageNumber = nextObjectNumber++;
                int contentNumber = nextObjectNumber++;
                objects[imageNumber] = new PdfIndirectObject(imageNumber, 0, PdfWriter.BuildImageXObject(image));
                string content = "q\n" + FormatNumber(right - left) + " 0 0 " + FormatNumber(top - bottom) + " 0 0 cm\n/CropImage Do\nQ\n";
                objects[contentNumber] = new PdfIndirectObject(contentNumber, 0, new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(content)));
                var xObjects = new PdfDictionary(); xObjects.Items["CropImage"] = new PdfReference(imageNumber, 0);
                var resources = new PdfDictionary(); resources.Items["XObject"] = xObjects;
                page.Items["Resources"] = resources;
                page.Items["Contents"] = new PdfReference(contentNumber, 0);
                page.Items.Remove("Annots"); page.Items.Remove("StructParents");
            }
            return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value) ? security.InfoObjectNumber : null;
        });

        PdfReadDocument verified = PdfReadDocument.Open(output);
        foreach (int pageNumber in selected) {
            if (!string.IsNullOrWhiteSpace(verified.Pages[pageNumber - 1].ExtractText())) throw new InvalidOperationException("Destructive crop validation found extractable text on a replaced page.");
            if (verified.Pages[pageNumber - 1].GetAnnotations().Count != 0) throw new InvalidOperationException("Destructive crop validation found a live annotation on a replaced page.");
        }
        var preservationOptions = new PdfRewritePreservationOptions { PreservePageGeometry = false, PreserveAnnotations = false, PreserveLinkAnnotations = false, PreserveRevisionStructure = false };
        PdfRewritePreservationReport preservation = PdfRewritePreservation.AssertPreserved(pdf, output, preservationOptions);
        return new PdfDestructiveCropResult(output, pageTreePlan, contentPlan, preservation, renders);
    }

    private static void EnsureNoSelectedFormWidgets(PdfReadDocument document, IReadOnlyList<int> selected) {
        var pages = new HashSet<int>(selected);
        if (document.FormFields.Any(field => field.Widgets.Any(widget => widget.PageNumber.HasValue && pages.Contains(widget.PageNumber.Value)))) throw new NotSupportedException("Destructive crop blocks selected pages with AcroForm widgets so field values are not left reachable outside the replacement page content.");
    }

    private static void ValidateDestructiveCropRenders(IReadOnlyList<PdfPageRenderResult> renders, PdfDestructiveCropOptions options) {
        foreach (PdfPageRenderResult render in renders) {
            if (!render.Succeeded) throw new InvalidOperationException("Destructive crop page rendering failed: " + string.Join(" ", render.Diagnostics));
            PdfRenderCapabilityDiagnostic? blocked = render.CapabilityDiagnostics.FirstOrDefault(diagnostic => diagnostic.SupportLevel == PdfRenderSupportLevel.Unsupported || !options.AllowSimplifiedRendering && diagnostic.SupportLevel == PdfRenderSupportLevel.Simplified);
            if (blocked != null) throw new NotSupportedException("Destructive crop blocked renderer diagnostic " + blocked.Code + ": " + blocked.Message);
        }
    }

    private static string FormatNumber(double value) => value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
}
