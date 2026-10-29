namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    /// <summary>
    /// Adds a simple text stamp to selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] StampText(byte[] pdf, string text, PdfTextStampOptions? options = null, PdfLoadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(text, nameof(text));
        var cancellationToken = options?.CancellationToken ?? default;
        cancellationToken.ThrowIfCancellationRequested();
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageContent, readOptions, cancellationToken: cancellationToken);
        if (text.Length == 0) {
            throw new ArgumentException("Stamp text cannot be empty.", nameof(text));
        }

        var effectiveOptions = options ?? new PdfTextStampOptions();
        ValidateOptions(effectiveOptions);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions, out _, out _, cancellationToken);
        var document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        if (document.Pages.Count == 0) {
            throw new ArgumentException("PDF does not contain any pages.", nameof(pdf));
        }

        var selectedPages = NormalizePageNumbers(effectiveOptions.PageNumbers, document.Pages.Count);
        var selectedSet = new HashSet<int>(selectedPages);
        var pageObjectNumbers = document.Pages.Select(page => page.ObjectNumber).ToArray();
        int nextObjectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;
        StampFont stampFont = PrepareStampFont(effectiveOptions, text, objects, ref nextObjectNumber);
        var additionalObjects = new List<PdfPageExtractor.AdditionalObject> {
            new PdfPageExtractor.AdditionalObject(FontPseudoObjectNumber, stampFont.Font)
        };
        var overrides = new Dictionary<int, Dictionary<string, PdfObject>>();
        string fontResourceName = GetAvailableFontResourceName(objects, pageObjectNumbers);

        for (int i = 0; i < document.Pages.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            int pageNumber = i + 1;
            if (!selectedSet.Contains(pageNumber)) {
                continue;
            }

            var page = document.Pages[i];
            if (!effectiveOptions.BehindContent) IsolateExistingContents(objects, pageObjectNumbers[i], ref nextObjectNumber);
            int stampPseudoId = -1000 - pageNumber;
            var stampStream = BuildStampStream(
                text,
                fontResourceName,
                page.GetPageSize().Width,
                page.GetPageSize().Height,
                effectiveOptions,
                watermarkDefaults: false,
                stampFont.GlyphHex,
                stampFont.TextWidth);

            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(stampPseudoId, stampStream));
            overrides[page.ObjectNumber] = BuildPageOverrides(objects, pageObjectNumbers[i], fontResourceName, stampPseudoId, effectiveOptions.BehindContent);
        }

        PdfFileVersion fileVersion = PdfPageExtractor.GetSourceFileVersion(pdf);
        return PdfPageExtractor.ExtractPages(objects, document.UncheckedMetadata, pageObjectNumbers, overrides, additionalObjects, PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw, document, cancellationToken), fileVersion, cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Adds a simple text stamp to selected pages from the current position of a readable stream, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] StampText(Stream stream, string text, PdfTextStampOptions? options = null) {
        return StampText(ReadStream(stream, nameof(stream), cancellationToken: options?.CancellationToken ?? default), text, options);
    }

    /// <summary>
    /// Adds a simple text stamp to selected pages and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampText(byte[] pdf, Stream outputStream, string text, PdfTextStampOptions? options = null) {
        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, StampText(pdf, text, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds a simple text stamp to selected pages from the current position of a readable stream and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampText(Stream stream, Stream outputStream, string text, PdfTextStampOptions? options = null) {
        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, StampText(stream, text, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Writes a new PDF with a simple text stamp from the current position of a readable stream.
    /// </summary>
    public static void StampText(Stream stream, string outputPath, string text, PdfTextStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, StampText(stream, text, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds a simple text stamp to selected pages from a PDF file and returns the stamped PDF bytes.
    /// </summary>
    public static byte[] StampTextToBytes(string inputPath, string text, PdfTextStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return StampText(ReadPath(inputPath, options?.CancellationToken ?? default), text, options);
    }

    /// <summary>
    /// Adds a simple text stamp to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampText(string inputPath, Stream outputStream, string text, PdfTextStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, StampText(ReadPath(inputPath, options?.CancellationToken ?? default), text, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Writes a new PDF with a simple text stamp on selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static void StampText(string inputPath, string outputPath, string text, PdfTextStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, StampText(ReadPath(inputPath, options?.CancellationToken ?? default), text, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds a large diagonal text watermark to selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] WatermarkText(byte[] pdf, string text, PdfTextStampOptions? options = null, PdfLoadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(text, nameof(text));
        var cancellationToken = options?.CancellationToken ?? default;
        cancellationToken.ThrowIfCancellationRequested();
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageContent, readOptions, cancellationToken: cancellationToken);
        if (text.Length == 0) {
            throw new ArgumentException("Watermark text cannot be empty.", nameof(text));
        }

        var effectiveOptions = BuildWatermarkOptions(options);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions, out _, out _, cancellationToken);
        var document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        if (document.Pages.Count == 0) {
            throw new ArgumentException("PDF does not contain any pages.", nameof(pdf));
        }

        var selectedPages = NormalizePageNumbers(effectiveOptions.PageNumbers, document.Pages.Count);
        var selectedSet = new HashSet<int>(selectedPages);
        var pageObjectNumbers = document.Pages.Select(page => page.ObjectNumber).ToArray();
        int nextObjectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;
        StampFont stampFont = PrepareStampFont(effectiveOptions, text, objects, ref nextObjectNumber);
        var additionalObjects = new List<PdfPageExtractor.AdditionalObject> {
            new PdfPageExtractor.AdditionalObject(FontPseudoObjectNumber, stampFont.Font)
        };
        var overrides = new Dictionary<int, Dictionary<string, PdfObject>>();
        string fontResourceName = GetAvailableFontResourceName(objects, pageObjectNumbers);

        for (int i = 0; i < document.Pages.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            int pageNumber = i + 1;
            if (!selectedSet.Contains(pageNumber)) {
                continue;
            }

            var page = document.Pages[i];
            if (!effectiveOptions.BehindContent) IsolateExistingContents(objects, pageObjectNumbers[i], ref nextObjectNumber);
            var size = page.GetPageSize();
            int stampPseudoId = -1000 - pageNumber;
            var stampStream = BuildStampStream(
                text,
                fontResourceName,
                size.Width,
                size.Height,
                effectiveOptions,
                watermarkDefaults: true,
                stampFont.GlyphHex,
                stampFont.TextWidth);

            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(stampPseudoId, stampStream));
            overrides[page.ObjectNumber] = BuildPageOverrides(objects, pageObjectNumbers[i], fontResourceName, stampPseudoId, effectiveOptions.BehindContent);
        }

        PdfFileVersion fileVersion = PdfPageExtractor.GetSourceFileVersion(pdf);
        return PdfPageExtractor.ExtractPages(objects, document.UncheckedMetadata, pageObjectNumbers, overrides, additionalObjects, PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw, document, cancellationToken), fileVersion, cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Adds a large diagonal text watermark to selected pages from the current position of a readable stream, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] WatermarkText(Stream stream, string text, PdfTextStampOptions? options = null) {
        return WatermarkText(ReadStream(stream, nameof(stream), cancellationToken: options?.CancellationToken ?? default), text, options);
    }

    /// <summary>
    /// Adds a large diagonal text watermark to selected pages and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void WatermarkText(byte[] pdf, Stream outputStream, string text, PdfTextStampOptions? options = null) {
        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, WatermarkText(pdf, text, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds a large diagonal text watermark to selected pages from the current position of a readable stream and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void WatermarkText(Stream stream, Stream outputStream, string text, PdfTextStampOptions? options = null) {
        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, WatermarkText(stream, text, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Writes a new PDF with a large diagonal text watermark from the current position of a readable stream.
    /// </summary>
    public static void WatermarkText(Stream stream, string outputPath, string text, PdfTextStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkText(stream, text, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds a large diagonal text watermark to selected pages from a PDF file and returns the watermarked PDF bytes.
    /// </summary>
    public static byte[] WatermarkTextToBytes(string inputPath, string text, PdfTextStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return WatermarkText(ReadPath(inputPath, options?.CancellationToken ?? default), text, options);
    }

    /// <summary>
    /// Adds a large diagonal text watermark to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void WatermarkText(string inputPath, Stream outputStream, string text, PdfTextStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, WatermarkText(ReadPath(inputPath, options?.CancellationToken ?? default), text, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Writes a new PDF with a large diagonal text watermark on selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static void WatermarkText(string inputPath, string outputPath, string text, PdfTextStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkText(ReadPath(inputPath, options?.CancellationToken ?? default), text, options), options?.CancellationToken ?? default);
    }
    private readonly struct StampFont {
        internal StampFont(PdfObject font, string? glyphHex, double? textWidth) {
            Font = font;
            GlyphHex = glyphHex;
            TextWidth = textWidth;
        }

        internal PdfObject Font { get; }
        internal string? GlyphHex { get; }
        internal double? TextWidth { get; }
    }

    // Stamped text uses the same installed font as text written by OfficeIMO.Pdf, embedded as a subset,
    // so it matches the page and the result stays valid for archival profiles.
    private static StampFont PrepareStampFont(PdfTextStampOptions options, string text, Dictionary<int, PdfIndirectObject> objects, ref int nextObjectNumber) {
        PdfStandardFont font = options.Font;
        PdfStandardFont slot = PdfStandardFontMapper.GetFontFamily(font);
        if (options.EmbedFont &&
            (slot == PdfStandardFont.Helvetica || slot == PdfStandardFont.TimesRoman || slot == PdfStandardFont.Courier) &&
            PdfCreateOptions.TryResolveDefaultEmbeddedFontFamily(slot, out PdfEmbeddedFontFamily? family) && family != null) {
            string name = font.ToString();
            bool bold = name.Contains("Bold");
            bool italic = name.Contains("Italic") || name.Contains("Oblique");
            byte[]? face = bold && italic ? family.BoldItalicSnapshot ?? family.BoldSnapshot
                : bold ? family.BoldSnapshot
                : italic ? family.ItalicSnapshot
                : null;
            face ??= family.RegularSnapshot;
            if (PdfFormFiller.TryCreateEmbeddedTextFont(face, family.FamilyName, text, options.FontSize, objects, ref nextObjectNumber,
                    out PdfObject type0Font, out string glyphHex, out double textWidth)) {
                return new StampFont(type0Font, glyphHex, textWidth);
            }
        }

        return new StampFont(BuildFontObject(font), null, null);
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
