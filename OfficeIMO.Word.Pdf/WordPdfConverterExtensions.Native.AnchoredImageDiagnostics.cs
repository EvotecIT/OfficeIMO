using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// Records unsupported running-picture wrapping before deferred callbacks and
    /// report capture, using the same eligibility owner as actual projection.
    /// </summary>
    private static void RecordNativeRunningImageWrapDiagnostics(WordSection section, WordPdfExportOptions? options,
        bool mapsHeader, bool mapsFooter) {
        if (options == null) return;
        Record(section.ResolveDefaultHeader(), "default header image", mapsHeader);
        Record(section.ResolveDefaultFooter(), "default footer image", mapsFooter);
        if (section.DifferentFirstPage) {
            Record(section.ResolveFirstHeader(), "first header image", mapsHeader);
            Record(section.ResolveFirstFooter(), "first footer image", mapsFooter);
        }
        if (section.DocumentOddEvenSettingEnabled) {
            Record(section.ResolveEvenHeader(), "even header image", mapsHeader);
            Record(section.ResolveEvenFooter(), "even footer image", mapsFooter);
        }

        void Record(WordHeaderFooter? story, string source, bool mapped) {
            OpenXmlElement? root = (OpenXmlElement?)story?._header ?? story?._footer;
            if (root == null) return;
            foreach (W.Paragraph element in root.Descendants<W.Paragraph>()) {
                // Rich tables suppress merge-continuation content; the flat
                // projection has a different established image enumeration.
                if (mapped && element.Ancestors<W.TableCell>().Any(cell =>
                    cell.TableCellProperties?.HorizontalMerge is { } horizontal && horizontal.Val?.Value != W.MergedCellValues.Restart ||
                    cell.TableCellProperties?.VerticalMerge is { } vertical && vertical.Val?.Value != W.MergedCellValues.Restart)) continue;
                var paragraph = new WordParagraph(story!.Document, element);
                foreach (WordImage image in EnumerateNativeParagraphImages(paragraph, options.CancellationToken)) {
                    if (image._Image == null || !HasNativeImageTextWrap(image._Image)) continue;
                    WordDrawingLayoutReader.TryRead(image._Image, out WordDrawingLayoutInfo layout);
                    PdfCore.PdfPageSize pageSize = GetNativePageSize(section, options);
                    PdfCore.PdfPageMargins margins = GetNativeMargins(section, options, (0D, 0D));
                    bool supported = TryResolveNativeAnchoredImageSource(image._Image, section,
                        pageSize, margins, layout.WidthPoints, layout.HeightPoints,
                        out _, out NativeAnchoredImagePlacement placement, out _, out string? reason);
                    if (supported && mapped) {
                        RecordNativeWrappedImagePageClipping(options, image._Image, source, pageSize,
                            layout.WidthPoints, layout.HeightPoints, new PdfCore.PdfAnchoredImagePosition(
                                placement.X, placement.Y, placement.MarginX, margins.Left.ToPoints(), placement.ParagraphY));
                        continue;
                    }
                    RecordNativeImageWrapApproximation(options, image._Image, source,
                        supported
                            ? "The wrapped picture uses the bounded header/footer projection because this story's presentation or content is outside rich running layout."
                            : "The wrapped picture uses a bounded header/footer projection. " + reason);
                }
            }
        }
    }

    /// <summary>
    /// Records statically known page clipping before running callbacks. Paragraph-relative
    /// vertical clipping remains a resolved-layout concern, just as for existing None anchors.
    /// </summary>
    private static void RecordNativeWrappedImagePageClipping(WordPdfExportOptions? options, W.Drawing drawing,
        string source, PdfCore.PdfPageSize pageSize, double width, double height,
        PdfCore.PdfAnchoredImagePosition position) {
        if (options == null || !(position.X < 0D || position.X + width > pageSize.Width.ToPoints() + 0.001D ||
                !position.ParagraphY && (position.Y < 0D || position.Y + height > pageSize.Height.ToPoints() + 0.001D)) ||
            !options.DiagnosedWrappedImageClippingElements.Add(drawing)) return;
        AddNativeExportWarning(options, "native-anchored-image-clipped", source,
            "The anchored image extends beyond the page; the page boundary clips its visible appearance.");
    }

    private static bool HasNativeImageTextWrap(W.Drawing drawing) =>
        WordDrawingLayoutReader.TryRead(drawing, out WordDrawingLayoutInfo layout) &&
        layout.Placement == WordDrawingPlacementKind.Anchored && layout.Wrap != WordDrawingWrapKind.None;

    /// <summary>Retains one approximation per visible source drawing, across deferred materializations.</summary>
    private static void RecordNativeImageWrapApproximation(WordPdfExportOptions? options, W.Drawing drawing,
        string source, string message) {
        if (options == null || !options.DiagnosedImageWrapElements.Add(drawing)) return;
        AddNativeExportWarning(options, "native-anchored-image-wrap-approximated", source, message);
    }
}
