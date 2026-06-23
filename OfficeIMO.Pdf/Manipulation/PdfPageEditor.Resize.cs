using System.Globalization;

namespace OfficeIMO.Pdf;

public static partial class PdfPageEditor {
    private const int ResizeContentPrefixBasePseudoObjectNumber = -10000;
    private const int ResizeContentSuffixBasePseudoObjectNumber = -20000;

    /// <summary>
    /// Creates a new PDF with selected pages scaled into the supplied target page size.
    /// If no page numbers are supplied, all pages are resized.
    /// </summary>
    public static byte[] ResizePages(byte[] pdf, PageSize pageSize, params int[] pageNumbers) {
        return ResizePages(pdf, new PdfPageResizeOptions(pageSize), pageNumbers);
    }

    /// <summary>
    /// Creates a new PDF with selected pages scaled into the target page size described by <paramref name="options"/>.
    /// If no page numbers are supplied, all pages are resized.
    /// </summary>
    public static byte[] ResizePages(byte[] pdf, PdfPageResizeOptions options, params int[] pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        ValidateResizeOptions(options);
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
        var selectedPages = pageNumbers.Length == 0
            ? Enumerable.Range(1, document.Pages.Count).ToArray()
            : pageNumbers;
        ValidatePageNumbers(selectedPages, document.Pages.Count, nameof(pageNumbers));

        var selected = new HashSet<int>(selectedPages);
        var pageObjectNumbers = document.Pages.Select(static page => page.ObjectNumber).ToArray();
        var overrides = new Dictionary<int, Dictionary<string, PdfObject>>();
        var additionalObjects = new List<PdfPageExtractor.AdditionalObject>();
        var transformsByPageObjectNumber = new Dictionary<int, PageResizeTransform>();

        for (int i = 0; i < document.Pages.Count; i++) {
            int pageNumber = i + 1;
            if (!selected.Contains(pageNumber)) {
                continue;
            }

            int pageObjectNumber = document.Pages[i].ObjectNumber;
            if (!objects.TryGetValue(pageObjectNumber, out var indirect) ||
                indirect.Value is not PdfDictionary pageDictionary) {
                throw new InvalidOperationException("PDF page object " + pageObjectNumber.ToString(CultureInfo.InvariantCulture) + " was not found.");
            }

            var geometry = document.Pages[i].GetGeometry();
            var sourceBox = geometry.EffectiveBox;
            double sourceLeft = sourceBox?.Left ?? 0D;
            double sourceBottom = sourceBox?.Bottom ?? 0D;
            double sourceWidth = sourceBox?.Width ?? document.Pages[i].GetPageSize().Width;
            double sourceHeight = sourceBox?.Height ?? document.Pages[i].GetPageSize().Height;
            ValidateSourcePageSize(sourceWidth, sourceHeight, pageNumber);

            var transform = CalculateResizeTransform(pageObjectNumber, sourceLeft, sourceBottom, sourceWidth, sourceHeight, options);
            transformsByPageObjectNumber[pageObjectNumber] = transform;
            int prefixPseudoObjectNumber = ResizeContentPrefixBasePseudoObjectNumber - i;
            int suffixPseudoObjectNumber = ResizeContentSuffixBasePseudoObjectNumber - i;
            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(prefixPseudoObjectNumber, BuildResizeContentStream(transform, sourceLeft, sourceBottom, sourceWidth, sourceHeight)));
            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(suffixPseudoObjectNumber, new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes("\nQ\n"))));

            var pageOverrides = new Dictionary<string, PdfObject>(StringComparer.Ordinal) {
                ["MediaBox"] = CreatePageBoxArray(0D, 0D, options.PageSize.Width, options.PageSize.Height),
                ["CropBox"] = CreatePageBoxArray(0D, 0D, options.PageSize.Width, options.PageSize.Height),
                ["UserUnit"] = new PdfNumber(1D),
                ["Rotate"] = new PdfNumber(0D),
                ["Contents"] = BuildResizedContentsArray(
                    objects,
                    pageDictionary.Items.TryGetValue("Contents", out var contents) ? contents : null,
                    prefixPseudoObjectNumber,
                    suffixPseudoObjectNumber)
            };

            AddNormalizedProductionBoxes(pageOverrides, geometry, options.PageSize);
            if (TryBuildTransformedAnnotations(objects, pageDictionary, transform, out PdfArray? transformedAnnotations)) {
                pageOverrides["Annots"] = transformedAnnotations!;
            }

            overrides[pageObjectNumber] = pageOverrides;
        }

        PdfFileVersion fileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(pdf));
        PdfPageExtractor.CatalogRewriteState catalogState = PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw);
        TransformCatalogDestinationsForResize(objects, catalogState, transformsByPageObjectNumber);
        return PdfPageExtractor.ExtractPages(
            objects,
            document.Metadata,
            pageObjectNumbers,
            overrides,
            additionalObjects,
            catalogState,
            fileVersion);
    }

    /// <summary>Creates a new PDF with selected pages scaled into the supplied target page size from a readable stream.</summary>
    public static byte[] ResizePages(Stream stream, PageSize pageSize, params int[] pageNumbers) {
        return ResizePages(ReadStream(stream, nameof(stream)), pageSize, pageNumbers);
    }

    /// <summary>Creates a new PDF with selected pages scaled into the target page size described by <paramref name="options"/> from a readable stream.</summary>
    public static byte[] ResizePages(Stream stream, PdfPageResizeOptions options, params int[] pageNumbers) {
        return ResizePages(ReadStream(stream, nameof(stream)), options, pageNumbers);
    }

    /// <summary>Writes a new PDF with selected pages scaled into the supplied target page size.</summary>
    public static void ResizePages(byte[] pdf, Stream outputStream, PageSize pageSize, params int[] pageNumbers) {
        WriteOutput(outputStream, ResizePages(pdf, pageSize, pageNumbers));
    }

    /// <summary>Writes a new PDF with selected pages scaled into the target page size described by <paramref name="options"/>.</summary>
    public static void ResizePages(byte[] pdf, Stream outputStream, PdfPageResizeOptions options, params int[] pageNumbers) {
        WriteOutput(outputStream, ResizePages(pdf, options, pageNumbers));
    }

    /// <summary>Writes a new PDF with selected pages scaled into the supplied target page size from a readable stream.</summary>
    public static void ResizePages(Stream inputStream, Stream outputStream, PageSize pageSize, params int[] pageNumbers) {
        WriteOutput(outputStream, ResizePages(inputStream, pageSize, pageNumbers));
    }

    /// <summary>Writes a new PDF with selected pages scaled into the target page size described by <paramref name="options"/> from a readable stream.</summary>
    public static void ResizePages(Stream inputStream, Stream outputStream, PdfPageResizeOptions options, params int[] pageNumbers) {
        WriteOutput(outputStream, ResizePages(inputStream, options, pageNumbers));
    }

    /// <summary>Writes a new PDF file with selected pages scaled into the supplied target page size.</summary>
    public static void ResizePages(string inputPath, string outputPath, PageSize pageSize, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, ResizePages(File.ReadAllBytes(inputPath), pageSize, pageNumbers));
    }

    /// <summary>Writes a new PDF file with selected pages scaled into the target page size described by <paramref name="options"/>.</summary>
    public static void ResizePages(string inputPath, string outputPath, PdfPageResizeOptions options, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, ResizePages(File.ReadAllBytes(inputPath), options, pageNumbers));
    }

    /// <summary>Creates a new PDF with selected pages scaled into the supplied target page size from a file path.</summary>
    public static byte[] ResizePages(string inputPath, PageSize pageSize, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ResizePages(File.ReadAllBytes(inputPath), pageSize, pageNumbers);
    }

    /// <summary>Creates a new PDF with selected pages scaled into the target page size described by <paramref name="options"/> from a file path.</summary>
    public static byte[] ResizePages(string inputPath, PdfPageResizeOptions options, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ResizePages(File.ReadAllBytes(inputPath), options, pageNumbers);
    }

    /// <summary>Creates a new PDF with the inclusive one-based page range scaled into the supplied target page size.</summary>
    public static byte[] ResizePageRange(byte[] pdf, PageSize pageSize, int firstPage, int lastPage) {
        return ResizePages(pdf, pageSize, BuildInclusivePageRange(firstPage, lastPage, nameof(lastPage)));
    }

    /// <summary>Creates a new PDF with the inclusive one-based page range scaled into the target page size described by <paramref name="options"/>.</summary>
    public static byte[] ResizePageRange(byte[] pdf, PdfPageResizeOptions options, int firstPage, int lastPage) {
        return ResizePages(pdf, options, BuildInclusivePageRange(firstPage, lastPage, nameof(lastPage)));
    }

    /// <summary>Creates a new PDF with the supplied page ranges scaled into the supplied target page size.</summary>
    public static byte[] ResizePageRanges(byte[] pdf, PageSize pageSize, params PdfPageRange[] pageRanges) {
        return ResizePages(pdf, pageSize, ExpandPageRangesDistinct(pageRanges, nameof(pageRanges)));
    }

    /// <summary>Creates a new PDF with the supplied page ranges scaled into the target page size described by <paramref name="options"/>.</summary>
    public static byte[] ResizePageRanges(byte[] pdf, PdfPageResizeOptions options, params PdfPageRange[] pageRanges) {
        return ResizePages(pdf, options, ExpandPageRangesDistinct(pageRanges, nameof(pageRanges)));
    }

    private static PdfArray BuildResizedContentsArray(Dictionary<int, PdfIndirectObject> objects, PdfObject? existingContents, int prefixPseudoObjectNumber, int suffixPseudoObjectNumber) {
        var result = new PdfArray();
        result.Items.Add(new PdfReference(prefixPseudoObjectNumber, 0));
        AppendContentEntries(objects, result, existingContents);
        result.Items.Add(new PdfReference(suffixPseudoObjectNumber, 0));
        return result;
    }

    private static void AppendContentEntries(Dictionary<int, PdfIndirectObject> objects, PdfArray target, PdfObject? contents) {
        if (contents is null) {
            return;
        }

        if (contents is PdfArray directArray) {
            foreach (var item in directArray.Items) {
                target.Items.Add(item);
            }

            return;
        }

        if (contents is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out var indirect) &&
            indirect.Value is PdfArray referencedArray) {
            foreach (var item in referencedArray.Items) {
                target.Items.Add(item);
            }

            return;
        }

        target.Items.Add(contents);
    }

    private static PdfStream BuildResizeContentStream(PageResizeTransform transform, double sourceLeft, double sourceBottom, double sourceWidth, double sourceHeight) {
        string content =
            "q\n" +
            FormatResizeNumber(transform.ScaleX) + " 0 0 " +
            FormatResizeNumber(transform.ScaleY) + " " +
            FormatResizeNumber(transform.TranslateX) + " " +
            FormatResizeNumber(transform.TranslateY) + " cm\n" +
            FormatResizeNumber(sourceLeft) + " " +
            FormatResizeNumber(sourceBottom) + " " +
            FormatResizeNumber(sourceWidth) + " " +
            FormatResizeNumber(sourceHeight) + " re\n" +
            "W n\n";
        return new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(content));
    }

    private static void AddNormalizedProductionBoxes(Dictionary<string, PdfObject> pageOverrides, PdfPageGeometry geometry, PageSize pageSize) {
        PdfArray targetBox = CreatePageBoxArray(0D, 0D, pageSize.Width, pageSize.Height);
        if (geometry.BleedBox is not null) {
            pageOverrides["BleedBox"] = ClonePageBoxArray(targetBox);
        }

        if (geometry.TrimBox is not null) {
            pageOverrides["TrimBox"] = ClonePageBoxArray(targetBox);
        }

        if (geometry.ArtBox is not null) {
            pageOverrides["ArtBox"] = ClonePageBoxArray(targetBox);
        }
    }

    private static PdfArray ClonePageBoxArray(PdfArray source) {
        var clone = new PdfArray();
        foreach (PdfObject item in source.Items) {
            clone.Items.Add(ClonePdfObject(item));
        }

        return clone;
    }

    private static bool TryBuildTransformedAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        PageResizeTransform transform,
        out PdfArray? transformedAnnotations) {
        transformedAnnotations = null;
        if (!pageDictionary.Items.TryGetValue("Annots", out PdfObject? annotationsObject) ||
            PdfObjectLookup.Resolve(objects, annotationsObject) is not PdfArray annotations) {
            return false;
        }

        transformedAnnotations = new PdfArray();
        foreach (PdfObject annotationObject in annotations.Items) {
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, annotationObject);
            if (resolved is PdfDictionary annotationDictionary) {
                var clonedAnnotation = (PdfDictionary)ClonePdfObject(annotationDictionary);
                TransformAnnotationRectangle(clonedAnnotation, transform);
                TransformDestinationsInObject(objects, clonedAnnotation, new Dictionary<int, PageResizeTransform> {
                    [transform.PageObjectNumber] = transform
                }, new HashSet<int>(), new HashSet<PdfArray>());
                transformedAnnotations.Items.Add(clonedAnnotation);
            } else {
                transformedAnnotations.Items.Add(ClonePdfObject(annotationObject));
            }
        }

        return true;
    }

    private static void TransformAnnotationRectangle(PdfDictionary annotation, PageResizeTransform transform) {
        if (!annotation.Items.TryGetValue("Rect", out PdfObject? rectObject) ||
            rectObject is not PdfArray rect ||
            rect.Items.Count < 4 ||
            rect.Items[0] is not PdfNumber x1 ||
            rect.Items[1] is not PdfNumber y1 ||
            rect.Items[2] is not PdfNumber x2 ||
            rect.Items[3] is not PdfNumber y2) {
            return;
        }

        double transformedX1 = TransformX(x1.Value, transform);
        double transformedY1 = TransformY(y1.Value, transform);
        double transformedX2 = TransformX(x2.Value, transform);
        double transformedY2 = TransformY(y2.Value, transform);
        var transformedRect = new PdfArray();
        transformedRect.Items.Add(new PdfNumber(Math.Min(transformedX1, transformedX2)));
        transformedRect.Items.Add(new PdfNumber(Math.Min(transformedY1, transformedY2)));
        transformedRect.Items.Add(new PdfNumber(Math.Max(transformedX1, transformedX2)));
        transformedRect.Items.Add(new PdfNumber(Math.Max(transformedY1, transformedY2)));
        annotation.Items["Rect"] = transformedRect;
    }

    private static double TransformX(double value, PageResizeTransform transform) =>
        value * transform.ScaleX + transform.TranslateX;

    private static double TransformY(double value, PageResizeTransform transform) =>
        value * transform.ScaleY + transform.TranslateY;

    private static void TransformCatalogDestinationsForResize(
        Dictionary<int, PdfIndirectObject> objects,
        PdfPageExtractor.CatalogRewriteState catalogState,
        IReadOnlyDictionary<int, PageResizeTransform> transformsByPageObjectNumber) {
        if (transformsByPageObjectNumber.Count == 0) {
            return;
        }

        var visitedReferences = new HashSet<int>();
        var transformedArrays = new HashSet<PdfArray>();
        TransformDestinationsInObject(objects, catalogState.Outlines, transformsByPageObjectNumber, visitedReferences, transformedArrays);
        TransformDestinationsInObject(objects, catalogState.NamedDestinations, transformsByPageObjectNumber, visitedReferences, transformedArrays);
        TransformDestinationsInObject(objects, catalogState.NamedDestinationNameTree, transformsByPageObjectNumber, visitedReferences, transformedArrays);
        TransformDestinationsInObject(objects, catalogState.OpenAction, transformsByPageObjectNumber, visitedReferences, transformedArrays);
    }

    private static void TransformDestinationsInObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        IReadOnlyDictionary<int, PageResizeTransform> transformsByPageObjectNumber,
        HashSet<int> visitedReferences,
        HashSet<PdfArray> transformedArrays) {
        switch (value) {
            case PdfReference reference:
                if (!visitedReferences.Add(reference.ObjectNumber) ||
                    !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                    IsPageDictionary(indirect.Value)) {
                    return;
                }

                TransformDestinationsInObject(objects, indirect.Value, transformsByPageObjectNumber, visitedReferences, transformedArrays);
                return;
            case PdfArray array:
                TransformDestinationArray(array, transformsByPageObjectNumber, transformedArrays);
                foreach (PdfObject item in array.Items) {
                    TransformDestinationsInObject(objects, item, transformsByPageObjectNumber, visitedReferences, transformedArrays);
                }

                return;
            case PdfDictionary dictionary:
                foreach (PdfObject item in dictionary.Items.Values) {
                    TransformDestinationsInObject(objects, item, transformsByPageObjectNumber, visitedReferences, transformedArrays);
                }

                return;
            case PdfStream stream:
                TransformDestinationsInObject(objects, stream.Dictionary, transformsByPageObjectNumber, visitedReferences, transformedArrays);
                return;
        }
    }

    private static bool IsPageDictionary(PdfObject value) =>
        value is PdfDictionary dictionary &&
        dictionary.Get<PdfName>("Type")?.Name == "Page";

    private static void TransformDestinationArray(
        PdfArray array,
        IReadOnlyDictionary<int, PageResizeTransform> transformsByPageObjectNumber,
        HashSet<PdfArray> transformedArrays) {
        if (!transformedArrays.Add(array) ||
            array.Items.Count < 2 ||
            array.Items[0] is not PdfReference pageReference ||
            !transformsByPageObjectNumber.TryGetValue(pageReference.ObjectNumber, out PageResizeTransform transform) ||
            array.Items[1] is not PdfName destinationMode) {
            return;
        }

        switch (destinationMode.Name) {
            case "XYZ":
                TransformDestinationCoordinate(array, 2, transform, isX: true);
                TransformDestinationCoordinate(array, 3, transform, isX: false);
                break;
            case "FitH":
            case "FitBH":
                TransformDestinationCoordinate(array, 2, transform, isX: false);
                break;
            case "FitV":
            case "FitBV":
                TransformDestinationCoordinate(array, 2, transform, isX: true);
                break;
            case "FitR":
                TransformDestinationCoordinate(array, 2, transform, isX: true);
                TransformDestinationCoordinate(array, 3, transform, isX: false);
                TransformDestinationCoordinate(array, 4, transform, isX: true);
                TransformDestinationCoordinate(array, 5, transform, isX: false);
                break;
        }
    }

    private static void TransformDestinationCoordinate(PdfArray array, int index, PageResizeTransform transform, bool isX) {
        if (index >= array.Items.Count ||
            array.Items[index] is not PdfNumber coordinate) {
            return;
        }

        array.Items[index] = new PdfNumber(isX
            ? TransformX(coordinate.Value, transform)
            : TransformY(coordinate.Value, transform));
    }

    private static PdfObject ClonePdfObject(PdfObject value) {
        switch (value) {
            case PdfNumber number:
                return new PdfNumber(number.Value);
            case PdfBoolean boolean:
                return new PdfBoolean(boolean.Value);
            case PdfName name:
                return new PdfName(name.Name);
            case PdfStringObj text:
                return new PdfStringObj(text.RawBytes, text.UseTextStringEncoding);
            case PdfArray array:
                var clonedArray = new PdfArray();
                foreach (PdfObject item in array.Items) {
                    clonedArray.Items.Add(ClonePdfObject(item));
                }

                return clonedArray;
            case PdfDictionary dictionary:
                var clonedDictionary = new PdfDictionary();
                foreach (var item in dictionary.Items) {
                    clonedDictionary.Items[item.Key] = ClonePdfObject(item.Value);
                }

                return clonedDictionary;
            case PdfReference reference:
                return new PdfReference(reference.ObjectNumber, reference.Generation);
            case PdfStream stream:
                return new PdfStream((PdfDictionary)ClonePdfObject(stream.Dictionary), (byte[])stream.Data.Clone(), stream.DecodingFailed, stream.DecodingError);
            case PdfNull:
                return PdfNull.Instance;
            default:
                return value;
        }
    }

    private static PageResizeTransform CalculateResizeTransform(int pageObjectNumber, double sourceLeft, double sourceBottom, double sourceWidth, double sourceHeight, PdfPageResizeOptions options) {
        double margin = options.Margin;
        double availableWidth = options.PageSize.Width - margin * 2D;
        double availableHeight = options.PageSize.Height - margin * 2D;
        double scaleX = availableWidth / sourceWidth;
        double scaleY = availableHeight / sourceHeight;

        switch (options.Mode) {
            case PdfPageResizeMode.Fit:
                scaleX = scaleY = Math.Min(scaleX, scaleY);
                break;
            case PdfPageResizeMode.Fill:
                scaleX = scaleY = Math.Max(scaleX, scaleY);
                break;
            case PdfPageResizeMode.Stretch:
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(options), "Unsupported page resize mode.");
        }

        double scaledWidth = sourceWidth * scaleX;
        double scaledHeight = sourceHeight * scaleY;
        double translateX = margin + (availableWidth - scaledWidth) / 2D - sourceLeft * scaleX;
        double translateY = margin + (availableHeight - scaledHeight) / 2D - sourceBottom * scaleY;
        return new PageResizeTransform(scaleX, scaleY, translateX, translateY, pageObjectNumber);
    }

    private static void ValidateResizeOptions(PdfPageResizeOptions options) {
        Guard.Positive(options.PageSize.Width, nameof(options));
        Guard.Positive(options.PageSize.Height, nameof(options));
        if (!IsFinite(options.Margin) || options.Margin < 0D) {
            throw new ArgumentOutOfRangeException(nameof(options), "Resize margin must be a finite non-negative number.");
        }

        if (options.Margin * 2D >= options.PageSize.Width || options.Margin * 2D >= options.PageSize.Height) {
            throw new ArgumentOutOfRangeException(nameof(options), "Resize margin must leave a positive content area.");
        }
    }

    private static void ValidateSourcePageSize(double width, double height, int pageNumber) {
        if (!IsFinite(width) || !IsFinite(height) || width <= 0D || height <= 0D) {
            throw new InvalidOperationException("PDF page " + pageNumber.ToString(CultureInfo.InvariantCulture) + " does not expose a valid source page size.");
        }
    }

    private static string FormatResizeNumber(double value) {
        if (Math.Abs(value % 1) < 0.0000001) {
            return ((long)Math.Round(value)).ToString(CultureInfo.InvariantCulture);
        }

        return value.ToString("0.######", CultureInfo.InvariantCulture);
    }

    private readonly struct PageResizeTransform {
        public PageResizeTransform(double scaleX, double scaleY, double translateX, double translateY, int pageObjectNumber) {
            ScaleX = scaleX;
            ScaleY = scaleY;
            TranslateX = translateX;
            TranslateY = translateY;
            PageObjectNumber = pageObjectNumber;
        }

        public double ScaleX { get; }

        public double ScaleY { get; }

        public double TranslateX { get; }

        public double TranslateY { get; }

        public int PageObjectNumber { get; }
    }
}
