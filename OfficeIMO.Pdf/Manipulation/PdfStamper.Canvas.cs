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
        var requests = new List<PageStampRequest>(selectedPages.Count);
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
            requests.Add(new PageStampRequest(overlay, pageOptions));
        }

        return StampPageSetCore(pdf, requests, readOptions);
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

            if (HasInteractiveMetadata(item)) {
                throw new NotSupportedException("Existing-page canvas stamping accepts visual content only. Links, named destinations, and form controls require their dedicated document editors.");
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

    private static bool HasInteractiveMetadata(PdfCanvasItem item) {
        if (item is PdfCanvasTextItem text) return HasInteractiveTextRuns(text.Runs);
        if (item is PdfCanvasTextBoxItem textBox) return HasInteractiveTextRuns(textBox.Runs);
        if (item is PdfCanvasShapeItem shape) return shape.Block.LinkUri != null;
        if (item is PdfCanvasDrawingItem drawing) return drawing.Block.LinkUri != null;
        if (item is PdfCanvasImageItem image) return image.Block.LinkUri != null;
        if (item is not PdfCanvasTableItem table) return false;

        if (table.Block.Links.Count > 0) return true;
        for (int rowIndex = 0; rowIndex < table.Block.Cells.Count; rowIndex++) {
            IReadOnlyList<PdfTableCell> row = table.Block.Cells[rowIndex];
            for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                PdfTableCell cell = row[cellIndex];
                if (cell.LinkUri != null || cell.LinkDestinationName != null || cell.NamedDestinationName != null ||
                    cell.CheckBoxes.Count > 0 || cell.FormFields.Count > 0 ||
                    HasInteractiveTextRuns(cell.Runs) ||
                    cell.Images.Any(static cellImage => cellImage.LinkUri != null) ||
                    cell.Paragraphs.Any(static paragraph => HasInteractiveTextRuns(paragraph.Runs))) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool HasInteractiveTextRuns(IReadOnlyList<TextRun> runs) {
        for (int i = 0; i < runs.Count; i++) {
            if (runs[i].LinkUri != null || runs[i].LinkDestinationName != null) return true;
        }

        return false;
    }
}
