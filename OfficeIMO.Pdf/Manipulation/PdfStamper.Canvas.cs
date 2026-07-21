namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    internal static byte[] StampCanvas(
        byte[] pdf,
        Action<PdfPageCanvas, PdfStampPageContext> build,
        PdfCanvasStampOptions? options = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(build, nameof(build));
        PdfCanvasStampOptions effective = options ?? new PdfCanvasStampOptions();
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageContent, readOptions);

        PdfReadDocument target = PdfReadDocument.Open(pdf, readOptions);
        if (target.Pages.Count == 0) {
            throw new ArgumentException("PDF does not contain any pages.", nameof(pdf));
        }

        IReadOnlyList<int> selectedPages = effective.TargetPages?.Resolve(target.Pages.Count) ??
            Enumerable.Range(1, target.Pages.Count).ToArray();
        byte[] output = pdf;
        for (int selectedIndex = 0; selectedIndex < selectedPages.Count; selectedIndex++) {
            int pageNumber = selectedPages[selectedIndex];
            PdfReadPage page = target.Pages[pageNumber - 1];
            (double width, double height, _) = page.GetImportGeometry();
            var context = new PdfStampPageContext(pageNumber, target.Pages.Count, width, height, page.GetRotationDegrees());
            var canvas = new PdfPageCanvas();
            build(canvas, context);
            RejectNonVisualCanvasItems(canvas.Items);

            var overlayOptions = new PdfOptions {
                PageWidth = width,
                PageHeight = height,
                MarginLeft = 0D,
                MarginTop = 0D,
                MarginRight = 0D,
                MarginBottom = 0D
            };
            byte[] overlay = PdfDocument.Create(overlayOptions)
                .Canvas(generatedCanvas => CopyCanvasItems(canvas, generatedCanvas))
                .ToBytes();
            var pageOptions = new PdfPageOverlayOptions {
                TargetPages = PdfPageSelector.Parse(pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                Fit = PdfPageOverlayFit.Stretch,
                X = 0D,
                Y = 0D,
                Width = width,
                Height = height,
                Opacity = effective.Opacity,
                BehindContent = effective.BehindContent
            };
            output = StampPage(output, overlay, pageOptions, readOptions);
        }

        return output;
    }

    private static void CopyCanvasItems(PdfPageCanvas source, PdfPageCanvas target) {
        target.AddItems(source.Items);
    }

    private static void RejectNonVisualCanvasItems(IReadOnlyList<PdfCanvasItem> items) {
        for (int i = 0; i < items.Count; i++) {
            PdfCanvasItem item = items[i];
            if (item is PdfCanvasTextAnnotationItem || item is PdfCanvasFreeTextAnnotationItem || item is PdfCanvasHighlightAnnotationItem) {
                throw new NotSupportedException("Existing-page canvas stamping accepts visual content only. Use the annotation editor for interactive annotations.");
            }

            if (item is PdfCanvasOutlineItem) {
                throw new NotSupportedException("Existing-page canvas stamping accepts visual content only. Use the bookmark editor for document outlines.");
            }

            if (item is PdfCanvasClipItem clip) {
                RejectNonVisualCanvasItems(clip.Items);
            } else if (item is PdfCanvasEffectItem effect) {
                RejectNonVisualCanvasItems(effect.Items);
            } else if (item is PdfCanvasFigureItem figure) {
                RejectNonVisualCanvasItems(figure.Items);
            } else if (item is PdfCanvasStructureItem structure) {
                RejectNonVisualCanvasItems(structure.Items);
            } else if (item is PdfCanvasActualTextItem actualText) {
                RejectNonVisualCanvasItems(actualText.Items);
            }
        }
    }
}
