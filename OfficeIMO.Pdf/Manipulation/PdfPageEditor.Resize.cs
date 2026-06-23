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

            int rotationDegrees = NormalizeResizeRotation(document.Pages[i].GetRotationDegrees());
            var transform = CalculateResizeTransform(pageObjectNumber, sourceLeft, sourceBottom, sourceWidth, sourceHeight, rotationDegrees, options);
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
            overrides[pageObjectNumber] = pageOverrides;
        }

        PdfFileVersion fileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(pdf));
        PdfPageExtractor.CatalogRewriteState catalogState = PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw);
        TransformCatalogDestinationsForResize(objects, catalogState, transformsByPageObjectNumber);
        TransformPageAnnotationsForResize(objects, document, selected, transformsByPageObjectNumber, overrides);
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
            FormatResizeNumber(transform.A) + " " +
            FormatResizeNumber(transform.B) + " " +
            FormatResizeNumber(transform.C) + " " +
            FormatResizeNumber(transform.D) + " " +
            FormatResizeNumber(transform.E) + " " +
            FormatResizeNumber(transform.F) + " cm\n" +
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

    private static void TransformPageAnnotationsForResize(
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadDocument document,
        HashSet<int> selectedPages,
        IReadOnlyDictionary<int, PageResizeTransform> transformsByPageObjectNumber,
        Dictionary<int, Dictionary<string, PdfObject>> overrides) {
        if (transformsByPageObjectNumber.Count == 0) {
            return;
        }

        var visitedReferences = new HashSet<int>();
        var transformedArrays = new HashSet<PdfArray>();
        for (int i = 0; i < document.Pages.Count; i++) {
            int pageNumber = i + 1;
            int pageObjectNumber = document.Pages[i].ObjectNumber;
            if (!objects.TryGetValue(pageObjectNumber, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfDictionary pageDictionary) {
                continue;
            }

            PageResizeTransform? annotationGeometryTransform = selectedPages.Contains(pageNumber)
                ? transformsByPageObjectNumber[pageObjectNumber]
                : null;
            if (TryBuildTransformedAnnotations(objects, pageDictionary, annotationGeometryTransform, transformsByPageObjectNumber, visitedReferences, transformedArrays, out PdfArray? transformedAnnotations)) {
                if (!overrides.TryGetValue(pageObjectNumber, out Dictionary<string, PdfObject>? pageOverrides)) {
                    pageOverrides = new Dictionary<string, PdfObject>(StringComparer.Ordinal);
                    overrides[pageObjectNumber] = pageOverrides;
                }

                pageOverrides["Annots"] = transformedAnnotations!;
            }
        }
    }

    private static bool TryBuildTransformedAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        PageResizeTransform? annotationGeometryTransform,
        IReadOnlyDictionary<int, PageResizeTransform> transformsByPageObjectNumber,
        HashSet<int> visitedReferences,
        HashSet<PdfArray> transformedArrays,
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
                if (annotationGeometryTransform.HasValue) {
                    TransformAnnotationRectangle(clonedAnnotation, annotationGeometryTransform.Value);
                    TransformAnnotationCoordinateArrays(clonedAnnotation, annotationGeometryTransform.Value);
                }

                TransformDestinationsInObject(objects, clonedAnnotation, transformsByPageObjectNumber, visitedReferences, transformedArrays);
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

        (double transformedX1, double transformedY1) = TransformPoint(x1.Value, y1.Value, transform);
        (double transformedX2, double transformedY2) = TransformPoint(x2.Value, y1.Value, transform);
        (double transformedX3, double transformedY3) = TransformPoint(x1.Value, y2.Value, transform);
        (double transformedX4, double transformedY4) = TransformPoint(x2.Value, y2.Value, transform);
        var transformedRect = new PdfArray();
        transformedRect.Items.Add(new PdfNumber(Min(transformedX1, transformedX2, transformedX3, transformedX4)));
        transformedRect.Items.Add(new PdfNumber(Min(transformedY1, transformedY2, transformedY3, transformedY4)));
        transformedRect.Items.Add(new PdfNumber(Max(transformedX1, transformedX2, transformedX3, transformedX4)));
        transformedRect.Items.Add(new PdfNumber(Max(transformedY1, transformedY2, transformedY3, transformedY4)));
        annotation.Items["Rect"] = transformedRect;
    }

    private static void TransformAnnotationCoordinateArrays(PdfDictionary annotation, PageResizeTransform transform) {
        TransformPairedNumberArray(annotation, "QuadPoints", transform);
        TransformPairedNumberArray(annotation, "L", transform);
        TransformPairedNumberArray(annotation, "Vertices", transform);

        if (!annotation.Items.TryGetValue("InkList", out PdfObject? inkListObject) ||
            inkListObject is not PdfArray inkList) {
            return;
        }

        foreach (PdfObject strokeObject in inkList.Items) {
            if (strokeObject is PdfArray stroke) {
                TransformPairedNumberArray(stroke, transform);
            }
        }
    }

    private static void TransformPairedNumberArray(PdfDictionary dictionary, string key, PageResizeTransform transform) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value) ||
            value is not PdfArray array) {
            return;
        }

        TransformPairedNumberArray(array, transform);
    }

    private static void TransformPairedNumberArray(PdfArray array, PageResizeTransform transform) {
        for (int i = 0; i + 1 < array.Items.Count; i += 2) {
            if (array.Items[i] is PdfNumber x &&
                array.Items[i + 1] is PdfNumber y) {
                (double transformedX, double transformedY) = TransformPoint(x.Value, y.Value, transform);
                array.Items[i] = new PdfNumber(transformedX);
                array.Items[i + 1] = new PdfNumber(transformedY);
            }
        }
    }

    private static double TransformX(double value, PageResizeTransform transform) =>
        value * transform.A + transform.E;

    private static double TransformY(double value, PageResizeTransform transform) =>
        value * transform.D + transform.F;

    private static (double X, double Y) TransformPoint(double x, double y, PageResizeTransform transform) =>
        (transform.A * x + transform.C * y + transform.E, transform.B * x + transform.D * y + transform.F);

    private static double Min(double a, double b, double c, double d) =>
        Math.Min(Math.Min(a, b), Math.Min(c, d));

    private static double Max(double a, double b, double c, double d) =>
        Math.Max(Math.Max(a, b), Math.Max(c, d));

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
                TransformDestinationPoint(array, 2, 3, transform);
                break;
            case "FitH":
            case "FitBH":
                if (TransformRotatedFitHorizontalDestination(array, transform)) {
                    break;
                }

                TransformDestinationCoordinate(array, 2, transform, isX: false);
                break;
            case "FitV":
            case "FitBV":
                if (TransformRotatedFitVerticalDestination(array, transform)) {
                    break;
                }

                TransformDestinationCoordinate(array, 2, transform, isX: true);
                break;
            case "FitR":
                TransformDestinationRectangle(array, 2, 3, 4, 5, transform);
                break;
        }
    }

    private static bool TransformRotatedFitHorizontalDestination(PdfArray array, PageResizeTransform transform) {
        if (!transform.HasAxisSwap ||
            array.Items.Count <= 2 ||
            array.Items[2] is not PdfNumber top) {
            return false;
        }

        double anchorX = transform.SourceLeft;
        (double transformedX, double transformedY) = TransformPoint(anchorX, top.Value, transform);
        ReplaceDestinationWithXyz(array, transformedX, transformedY);
        return true;
    }

    private static bool TransformRotatedFitVerticalDestination(PdfArray array, PageResizeTransform transform) {
        if (!transform.HasAxisSwap ||
            array.Items.Count <= 2 ||
            array.Items[2] is not PdfNumber left) {
            return false;
        }

        double anchorY = transform.SourceBottom + transform.SourceHeight;
        (double transformedX, double transformedY) = TransformPoint(left.Value, anchorY, transform);
        ReplaceDestinationWithXyz(array, transformedX, transformedY);
        return true;
    }

    private static void ReplaceDestinationWithXyz(PdfArray array, double left, double top) {
        while (array.Items.Count > 2) {
            array.Items.RemoveAt(array.Items.Count - 1);
        }

        array.Items[1] = new PdfName("XYZ");
        array.Items.Add(new PdfNumber(left));
        array.Items.Add(new PdfNumber(top));
        array.Items.Add(PdfNull.Instance);
    }

    private static void TransformDestinationPoint(PdfArray array, int xIndex, int yIndex, PageResizeTransform transform) {
        if (xIndex < array.Items.Count &&
            yIndex < array.Items.Count &&
            array.Items[xIndex] is PdfNumber x &&
            array.Items[yIndex] is PdfNumber y) {
            (double transformedX, double transformedY) = TransformPoint(x.Value, y.Value, transform);
            array.Items[xIndex] = new PdfNumber(transformedX);
            array.Items[yIndex] = new PdfNumber(transformedY);
            return;
        }

        TransformDestinationCoordinate(array, xIndex, transform, isX: true);
        TransformDestinationCoordinate(array, yIndex, transform, isX: false);
    }

    private static void TransformDestinationRectangle(PdfArray array, int leftIndex, int bottomIndex, int rightIndex, int topIndex, PageResizeTransform transform) {
        if (leftIndex >= array.Items.Count ||
            bottomIndex >= array.Items.Count ||
            rightIndex >= array.Items.Count ||
            topIndex >= array.Items.Count ||
            array.Items[leftIndex] is not PdfNumber left ||
            array.Items[bottomIndex] is not PdfNumber bottom ||
            array.Items[rightIndex] is not PdfNumber right ||
            array.Items[topIndex] is not PdfNumber top) {
            TransformDestinationCoordinate(array, leftIndex, transform, isX: true);
            TransformDestinationCoordinate(array, bottomIndex, transform, isX: false);
            TransformDestinationCoordinate(array, rightIndex, transform, isX: true);
            TransformDestinationCoordinate(array, topIndex, transform, isX: false);
            return;
        }

        (double x1, double y1) = TransformPoint(left.Value, bottom.Value, transform);
        (double x2, double y2) = TransformPoint(left.Value, top.Value, transform);
        (double x3, double y3) = TransformPoint(right.Value, bottom.Value, transform);
        (double x4, double y4) = TransformPoint(right.Value, top.Value, transform);
        array.Items[leftIndex] = new PdfNumber(Min(x1, x2, x3, x4));
        array.Items[bottomIndex] = new PdfNumber(Min(y1, y2, y3, y4));
        array.Items[rightIndex] = new PdfNumber(Max(x1, x2, x3, x4));
        array.Items[topIndex] = new PdfNumber(Max(y1, y2, y3, y4));
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

    private static PageResizeTransform CalculateResizeTransform(int pageObjectNumber, double sourceLeft, double sourceBottom, double sourceWidth, double sourceHeight, int rotationDegrees, PdfPageResizeOptions options) {
        double margin = options.Margin;
        double availableWidth = options.PageSize.Width - margin * 2D;
        double availableHeight = options.PageSize.Height - margin * 2D;
        double visualSourceWidth = rotationDegrees == 90 || rotationDegrees == 270 ? sourceHeight : sourceWidth;
        double visualSourceHeight = rotationDegrees == 90 || rotationDegrees == 270 ? sourceWidth : sourceHeight;
        double scaleX = availableWidth / visualSourceWidth;
        double scaleY = availableHeight / visualSourceHeight;

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

        double scaledWidth = visualSourceWidth * scaleX;
        double scaledHeight = visualSourceHeight * scaleY;
        double visualTranslateX = margin + (availableWidth - scaledWidth) / 2D;
        double visualTranslateY = margin + (availableHeight - scaledHeight) / 2D;

        return rotationDegrees switch {
            90 => new PageResizeTransform(
                0D,
                -scaleY,
                scaleX,
                0D,
                visualTranslateX - sourceBottom * scaleX,
                visualTranslateY + (sourceWidth + sourceLeft) * scaleY,
                pageObjectNumber,
                sourceLeft,
                sourceBottom,
                sourceWidth,
                sourceHeight,
                rotationDegrees),
            180 => new PageResizeTransform(
                -scaleX,
                0D,
                0D,
                -scaleY,
                visualTranslateX + (sourceWidth + sourceLeft) * scaleX,
                visualTranslateY + (sourceHeight + sourceBottom) * scaleY,
                pageObjectNumber,
                sourceLeft,
                sourceBottom,
                sourceWidth,
                sourceHeight,
                rotationDegrees),
            270 => new PageResizeTransform(
                0D,
                scaleY,
                -scaleX,
                0D,
                visualTranslateX + (sourceHeight + sourceBottom) * scaleX,
                visualTranslateY - sourceLeft * scaleY,
                pageObjectNumber,
                sourceLeft,
                sourceBottom,
                sourceWidth,
                sourceHeight,
                rotationDegrees),
            _ => new PageResizeTransform(
                scaleX,
                0D,
                0D,
                scaleY,
                visualTranslateX - sourceLeft * scaleX,
                visualTranslateY - sourceBottom * scaleY,
                pageObjectNumber,
                sourceLeft,
                sourceBottom,
                sourceWidth,
                sourceHeight,
                rotationDegrees)
        };
    }

    private static int NormalizeResizeRotation(int rotationDegrees) {
        int normalized = rotationDegrees % 360;
        if (normalized < 0) {
            normalized += 360;
        }

        return normalized == 90 || normalized == 180 || normalized == 270 ? normalized : 0;
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
        public PageResizeTransform(double a, double b, double c, double d, double e, double f, int pageObjectNumber, double sourceLeft, double sourceBottom, double sourceWidth, double sourceHeight, int rotationDegrees) {
            A = a;
            B = b;
            C = c;
            D = d;
            E = e;
            F = f;
            PageObjectNumber = pageObjectNumber;
            SourceLeft = sourceLeft;
            SourceBottom = sourceBottom;
            SourceWidth = sourceWidth;
            SourceHeight = sourceHeight;
            RotationDegrees = rotationDegrees;
        }

        public double A { get; }

        public double B { get; }

        public double C { get; }

        public double D { get; }

        public double E { get; }

        public double F { get; }

        public int PageObjectNumber { get; }

        public double SourceLeft { get; }

        public double SourceBottom { get; }

        public double SourceWidth { get; }

        public double SourceHeight { get; }

        public int RotationDegrees { get; }

        public bool HasAxisSwap => RotationDegrees == 90 || RotationDegrees == 270;
    }
}
