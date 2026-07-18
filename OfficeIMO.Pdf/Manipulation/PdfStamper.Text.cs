namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    /// <summary>
    /// Adds a simple text stamp to selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] StampText(byte[] pdf, string text, PdfTextStampOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(text, nameof(text));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageContent);
        if (text.Length == 0) {
            throw new ArgumentException("Stamp text cannot be empty.", nameof(text));
        }

        var effectiveOptions = options ?? new PdfTextStampOptions();
        ValidateOptions(effectiveOptions);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Open(pdf);
        if (document.Pages.Count == 0) {
            throw new ArgumentException("PDF does not contain any pages.", nameof(pdf));
        }

        var selectedPages = NormalizePageNumbers(effectiveOptions.PageNumbers, document.Pages.Count);
        var selectedSet = new HashSet<int>(selectedPages);
        var pageObjectNumbers = document.Pages.Select(page => page.ObjectNumber).ToArray();
        var additionalObjects = new List<PdfPageExtractor.AdditionalObject> {
            new PdfPageExtractor.AdditionalObject(FontPseudoObjectNumber, BuildFontObject(effectiveOptions.Font))
        };
        var overrides = new Dictionary<int, Dictionary<string, PdfObject>>();
        string fontResourceName = GetAvailableFontResourceName(objects, pageObjectNumbers);

        for (int i = 0; i < document.Pages.Count; i++) {
            int pageNumber = i + 1;
            if (!selectedSet.Contains(pageNumber)) {
                continue;
            }

            var page = document.Pages[i];
            int stampPseudoId = -1000 - pageNumber;
            var stampStream = BuildStampStream(
                text,
                fontResourceName,
                page.GetPageSize().Width,
                page.GetPageSize().Height,
                effectiveOptions,
                watermarkDefaults: false);

            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(stampPseudoId, stampStream));
            overrides[page.ObjectNumber] = BuildPageOverrides(objects, pageObjectNumbers[i], fontResourceName, stampPseudoId, effectiveOptions.BehindContent);
        }

        PdfFileVersion fileVersion = PdfPageExtractor.GetSourceFileVersion(pdf);
        return PdfPageExtractor.ExtractPages(objects, document.Metadata, pageObjectNumbers, overrides, additionalObjects, PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw), fileVersion);
    }

    /// <summary>
    /// Adds a simple text stamp to selected pages from the current position of a readable stream, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] StampText(Stream stream, string text, PdfTextStampOptions? options = null) {
        return StampText(ReadStream(stream, nameof(stream)), text, options);
    }

    /// <summary>
    /// Adds a simple text stamp to selected pages and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampText(byte[] pdf, Stream outputStream, string text, PdfTextStampOptions? options = null) {
        WriteOutput(outputStream, StampText(pdf, text, options));
    }

    /// <summary>
    /// Adds a simple text stamp to selected pages from the current position of a readable stream and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampText(Stream stream, Stream outputStream, string text, PdfTextStampOptions? options = null) {
        WriteOutput(outputStream, StampText(stream, text, options));
    }

    /// <summary>
    /// Writes a new PDF with a simple text stamp from the current position of a readable stream.
    /// </summary>
    public static void StampText(Stream stream, string outputPath, string text, PdfTextStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, StampText(stream, text, options));
    }

    /// <summary>
    /// Adds a simple text stamp to selected pages from a PDF file and returns the stamped PDF bytes.
    /// </summary>
    public static byte[] StampTextToBytes(string inputPath, string text, PdfTextStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return StampText(File.ReadAllBytes(inputPath), text, options);
    }

    /// <summary>
    /// Adds a simple text stamp to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampText(string inputPath, Stream outputStream, string text, PdfTextStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, StampText(File.ReadAllBytes(inputPath), text, options));
    }

    /// <summary>
    /// Writes a new PDF with a simple text stamp on selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static void StampText(string inputPath, string outputPath, string text, PdfTextStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, StampText(File.ReadAllBytes(inputPath), text, options));
    }

    /// <summary>
    /// Adds a large diagonal text watermark to selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] WatermarkText(byte[] pdf, string text, PdfTextStampOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(text, nameof(text));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageContent);
        if (text.Length == 0) {
            throw new ArgumentException("Watermark text cannot be empty.", nameof(text));
        }

        var effectiveOptions = BuildWatermarkOptions(options);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Open(pdf);
        if (document.Pages.Count == 0) {
            throw new ArgumentException("PDF does not contain any pages.", nameof(pdf));
        }

        var selectedPages = NormalizePageNumbers(effectiveOptions.PageNumbers, document.Pages.Count);
        var selectedSet = new HashSet<int>(selectedPages);
        var pageObjectNumbers = document.Pages.Select(page => page.ObjectNumber).ToArray();
        var additionalObjects = new List<PdfPageExtractor.AdditionalObject> {
            new PdfPageExtractor.AdditionalObject(FontPseudoObjectNumber, BuildFontObject(effectiveOptions.Font))
        };
        var overrides = new Dictionary<int, Dictionary<string, PdfObject>>();
        string fontResourceName = GetAvailableFontResourceName(objects, pageObjectNumbers);

        for (int i = 0; i < document.Pages.Count; i++) {
            int pageNumber = i + 1;
            if (!selectedSet.Contains(pageNumber)) {
                continue;
            }

            var page = document.Pages[i];
            var size = page.GetPageSize();
            int stampPseudoId = -1000 - pageNumber;
            var stampStream = BuildStampStream(
                text,
                fontResourceName,
                size.Width,
                size.Height,
                effectiveOptions,
                watermarkDefaults: true);

            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(stampPseudoId, stampStream));
            overrides[page.ObjectNumber] = BuildPageOverrides(objects, pageObjectNumbers[i], fontResourceName, stampPseudoId, effectiveOptions.BehindContent);
        }

        PdfFileVersion fileVersion = PdfPageExtractor.GetSourceFileVersion(pdf);
        return PdfPageExtractor.ExtractPages(objects, document.Metadata, pageObjectNumbers, overrides, additionalObjects, PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw), fileVersion);
    }

    /// <summary>
    /// Adds a large diagonal text watermark to selected pages from the current position of a readable stream, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] WatermarkText(Stream stream, string text, PdfTextStampOptions? options = null) {
        return WatermarkText(ReadStream(stream, nameof(stream)), text, options);
    }

    /// <summary>
    /// Adds a large diagonal text watermark to selected pages and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void WatermarkText(byte[] pdf, Stream outputStream, string text, PdfTextStampOptions? options = null) {
        WriteOutput(outputStream, WatermarkText(pdf, text, options));
    }

    /// <summary>
    /// Adds a large diagonal text watermark to selected pages from the current position of a readable stream and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void WatermarkText(Stream stream, Stream outputStream, string text, PdfTextStampOptions? options = null) {
        WriteOutput(outputStream, WatermarkText(stream, text, options));
    }

    /// <summary>
    /// Writes a new PDF with a large diagonal text watermark from the current position of a readable stream.
    /// </summary>
    public static void WatermarkText(Stream stream, string outputPath, string text, PdfTextStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkText(stream, text, options));
    }

    /// <summary>
    /// Adds a large diagonal text watermark to selected pages from a PDF file and returns the watermarked PDF bytes.
    /// </summary>
    public static byte[] WatermarkTextToBytes(string inputPath, string text, PdfTextStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return WatermarkText(File.ReadAllBytes(inputPath), text, options);
    }

    /// <summary>
    /// Adds a large diagonal text watermark to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void WatermarkText(string inputPath, Stream outputStream, string text, PdfTextStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, WatermarkText(File.ReadAllBytes(inputPath), text, options));
    }

    /// <summary>
    /// Writes a new PDF with a large diagonal text watermark on selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static void WatermarkText(string inputPath, string outputPath, string text, PdfTextStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkText(File.ReadAllBytes(inputPath), text, options));
    }
    private static PdfTextStampOptions BuildWatermarkOptions(PdfTextStampOptions? options) {
        if (options is null) {
            return new PdfTextStampOptions {
                FontSize = 64,
                Color = PdfColor.LightGray,
                RotationDegrees = -45,
                Font = PdfStandardFont.HelveticaBold
            };
        }

        ValidateOptions(options);
        return options;
    }

}
