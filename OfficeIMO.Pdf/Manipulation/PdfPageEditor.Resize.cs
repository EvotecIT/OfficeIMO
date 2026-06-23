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

            var transform = CalculateResizeTransform(sourceLeft, sourceBottom, sourceWidth, sourceHeight, options);
            int prefixPseudoObjectNumber = ResizeContentPrefixBasePseudoObjectNumber - i;
            int suffixPseudoObjectNumber = ResizeContentSuffixBasePseudoObjectNumber - i;
            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(prefixPseudoObjectNumber, BuildResizeContentStream(transform)));
            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(suffixPseudoObjectNumber, new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes("\nQ\n"))));

            overrides[pageObjectNumber] = new Dictionary<string, PdfObject>(StringComparer.Ordinal) {
                ["MediaBox"] = CreatePageBoxArray(0D, 0D, options.PageSize.Width, options.PageSize.Height),
                ["CropBox"] = CreatePageBoxArray(0D, 0D, options.PageSize.Width, options.PageSize.Height),
                ["Contents"] = BuildResizedContentsArray(
                    objects,
                    pageDictionary.Items.TryGetValue("Contents", out var contents) ? contents : null,
                    prefixPseudoObjectNumber,
                    suffixPseudoObjectNumber)
            };
        }

        PdfFileVersion fileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(pdf));
        return PdfPageExtractor.ExtractPages(
            objects,
            document.Metadata,
            pageObjectNumbers,
            overrides,
            additionalObjects,
            PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw),
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

    private static PdfStream BuildResizeContentStream(PageResizeTransform transform) {
        string content =
            "q\n" +
            FormatResizeNumber(transform.ScaleX) + " 0 0 " +
            FormatResizeNumber(transform.ScaleY) + " " +
            FormatResizeNumber(transform.TranslateX) + " " +
            FormatResizeNumber(transform.TranslateY) + " cm\n";
        return new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(content));
    }

    private static PageResizeTransform CalculateResizeTransform(double sourceLeft, double sourceBottom, double sourceWidth, double sourceHeight, PdfPageResizeOptions options) {
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
        return new PageResizeTransform(scaleX, scaleY, translateX, translateY);
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
        public PageResizeTransform(double scaleX, double scaleY, double translateX, double translateY) {
            ScaleX = scaleX;
            ScaleY = scaleY;
            TranslateX = translateX;
            TranslateY = translateY;
        }

        public double ScaleX { get; }

        public double ScaleY { get; }

        public double TranslateX { get; }

        public double TranslateY { get; }
    }
}
