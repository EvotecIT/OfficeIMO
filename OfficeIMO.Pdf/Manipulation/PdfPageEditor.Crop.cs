using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    /// <summary>
    /// Non-destructively crops selected pages to a source rectangle and translates that rectangle to a zero-based page origin.
    /// Content outside the crop remains in the content streams but is clipped from display.
    /// </summary>
    public static byte[] CropAndTranslatePages(
        byte[] pdf,
        double left,
        double bottom,
        double right,
        double top,
        params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        ValidatePageBoxCoordinates(left, bottom, right, top);
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageTree);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        int[] selectedPages = pageNumbers.Length == 0
            ? Enumerable.Range(1, document.Pages.Count).ToArray()
            : pageNumbers;
        ValidatePageNumbers(selectedPages, document.Pages.Count, nameof(pageNumbers));

        var selected = new HashSet<int>(selectedPages);
        int[] pageObjectNumbers = document.Pages.Select(static page => page.ObjectNumber).ToArray();
        var overrides = new Dictionary<int, Dictionary<string, PdfObject>>();
        var additionalObjects = new List<PdfPageExtractor.AdditionalObject>();
        var transforms = new Dictionary<int, PageResizeTransform>();
        int nextPseudoObjectNumber = -1;
        double width = right - left;
        double height = top - bottom;

        for (int i = 0; i < document.Pages.Count; i++) {
            int pageNumber = i + 1;
            if (!selected.Contains(pageNumber)) {
                continue;
            }

            PdfReadPage readPage = document.Pages[i];
            if (!objects.TryGetValue(readPage.ObjectNumber, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfDictionary page) {
                throw new InvalidOperationException("PDF page object " + readPage.ObjectNumber.ToString(CultureInfo.InvariantCulture) + " was not found.");
            }

            ValidateCropInsidePage(readPage, left, bottom, right, top, pageNumber);
            var transform = new PageResizeTransform(
                1D, 0D, 0D, 1D, -left, -bottom,
                readPage.ObjectNumber,
                left, bottom, width, height,
                0,
                0D, 0D, width, height);
            transforms[readPage.ObjectNumber] = transform;
            int prefixObjectNumber = AllocateResizePseudoObjectNumber(ref nextPseudoObjectNumber);
            int suffixObjectNumber = AllocateResizePseudoObjectNumber(ref nextPseudoObjectNumber);
            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(
                prefixObjectNumber,
                BuildResizeContentStream(transform, left, bottom, width, height)));
            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(
                suffixObjectNumber,
                new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes("\nQ\n"))));

            var pageOverrides = new Dictionary<string, PdfObject>(StringComparer.Ordinal) {
                ["MediaBox"] = CreatePageBoxArray(0D, 0D, width, height),
                ["CropBox"] = CreatePageBoxArray(0D, 0D, width, height),
                ["Contents"] = BuildResizedContentsArray(
                    objects,
                    page.Items.TryGetValue("Contents", out PdfObject? contents) ? contents : null,
                    prefixObjectNumber,
                    suffixObjectNumber)
            };
            AddNormalizedProductionBoxes(pageOverrides, readPage.GetGeometry(), new PageSize(width, height));
            overrides[readPage.ObjectNumber] = pageOverrides;
        }

        PdfPageExtractor.CatalogRewriteState catalogState = PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw);
        var visitedReferences = new HashSet<int>();
        var transformedArrays = new HashSet<PdfArray>();
        TransformCatalogDestinationsForResize(objects, catalogState, transforms, visitedReferences, transformedArrays);
        TransformPageAnnotationsForResize(
            objects,
            document,
            selected,
            transforms,
            overrides,
            additionalObjects,
            ref nextPseudoObjectNumber,
            visitedReferences,
            transformedArrays);
        return PdfPageExtractor.ExtractPages(
            objects,
            document.UncheckedMetadata,
            pageObjectNumbers,
            overrides,
            additionalObjects,
            catalogState,
            PdfPageExtractor.GetSourceFileVersion(pdf));
    }

    /// <summary>Non-destructively crops and translates selected pages from a readable stream.</summary>
    public static byte[] CropAndTranslatePages(
        Stream stream,
        double left,
        double bottom,
        double right,
        double top,
        params int[] pageNumbers) {
        return CropAndTranslatePages(ReadStream(stream, nameof(stream)), left, bottom, right, top, pageNumbers);
    }

    private static void ValidateCropInsidePage(
        PdfReadPage page,
        double left,
        double bottom,
        double right,
        double top,
        int pageNumber) {
        PdfPageBox? source = page.GetGeometry().MediaBox ?? page.GetGeometry().EffectiveBox;
        if (source is null) {
            return;
        }

        const double tolerance = 0.001D;
        if (left < source.Left - tolerance || bottom < source.Bottom - tolerance ||
            right > source.Right + tolerance || top > source.Top + tolerance) {
            throw new ArgumentOutOfRangeException(
                nameof(left),
                "Crop rectangle must stay inside page " + pageNumber.ToString(CultureInfo.InvariantCulture) + " MediaBox.");
        }
    }
}
