using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides first-party text stamping helpers for PDFs that can be parsed by OfficeIMO.Pdf.
/// </summary>
public static class PdfStamper {
    private const int FontPseudoObjectNumber = -1;
    private const int ImagePseudoObjectNumber = -2;
    private const int ImageSoftMaskPseudoObjectNumber = -3;

    /// <summary>
    /// Adds a simple text stamp to selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] StampText(byte[] pdf, string text, PdfTextStampOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(text, nameof(text));
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);
        if (text.Length == 0) {
            throw new ArgumentException("Stamp text cannot be empty.", nameof(text));
        }

        var effectiveOptions = options ?? new PdfTextStampOptions();
        ValidateOptions(effectiveOptions);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
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

        return PdfPageExtractor.ExtractPages(objects, document.Metadata, pageObjectNumbers, overrides, additionalObjects, PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw));
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
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);
        if (text.Length == 0) {
            throw new ArgumentException("Watermark text cannot be empty.", nameof(text));
        }

        var effectiveOptions = BuildWatermarkOptions(options);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
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

        return PdfPageExtractor.ExtractPages(objects, document.Metadata, pageObjectNumbers, overrides, additionalObjects, PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw));
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

    /// <summary>
    /// Adds an image stamp to selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] StampImage(byte[] pdf, byte[] imageBytes, PdfImageStampOptions? options = null) {
        return StampImageCore(pdf, imageBytes, options, watermarkDefaults: false);
    }

    /// <summary>
    /// Adds an image stamp from the current position of a readable image stream to selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] StampImage(byte[] pdf, Stream imageStream, PdfImageStampOptions? options = null) {
        return StampImage(pdf, ReadStream(imageStream, nameof(imageStream)), options);
    }

    /// <summary>
    /// Adds an image stamp to selected pages from the current position of a readable stream, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] StampImage(Stream stream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        return StampImage(ReadStream(stream, nameof(stream)), imageBytes, options);
    }

    /// <summary>
    /// Adds an image stamp from the current position of readable PDF and image streams to selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] StampImage(Stream stream, Stream imageStream, PdfImageStampOptions? options = null) {
        return StampImage(ReadStream(stream, nameof(stream)), ReadStream(imageStream, nameof(imageStream)), options);
    }

    /// <summary>
    /// Adds an image stamp to selected pages and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampImage(byte[] pdf, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        WriteOutput(outputStream, StampImage(pdf, imageBytes, options));
    }

    /// <summary>
    /// Adds an image stamp from the current position of a readable image stream and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampImage(byte[] pdf, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        WriteOutput(outputStream, StampImage(pdf, imageStream, options));
    }

    /// <summary>
    /// Adds an image stamp from the current position of a readable PDF stream and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampImage(Stream stream, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        WriteOutput(outputStream, StampImage(stream, imageBytes, options));
    }

    /// <summary>
    /// Adds an image stamp from the current position of readable PDF and image streams and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampImage(Stream stream, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        WriteOutput(outputStream, StampImage(stream, imageStream, options));
    }

    private static byte[] StampImageCore(byte[] pdf, byte[] imageBytes, PdfImageStampOptions? options, bool watermarkDefaults) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(imageBytes, nameof(imageBytes));
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);
        if (imageBytes.Length == 0) {
            throw new ArgumentException("Image bytes cannot be empty.", nameof(imageBytes));
        }

        var effectiveOptions = options ?? new PdfImageStampOptions();
        ValidateImageOptions(effectiveOptions);
        var imageInfo = PdfDoc.ValidateImageBytes(imageBytes);

        if (!PdfWriter.TryBuildImageStream(
                imageBytes,
                imageInfo,
                effectiveOptions.Width ?? Math.Max(1, imageInfo.Width),
                effectiveOptions.Height ?? Math.Max(1, imageInfo.Height),
                out var imageStream,
                out string? unsupportedReason)) {
            throw new NotSupportedException(unsupportedReason ?? "Image format is not supported.");
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
        if (document.Pages.Count == 0) {
            throw new ArgumentException("PDF does not contain any pages.", nameof(pdf));
        }

        var selectedPages = NormalizePageNumbers(effectiveOptions.PageNumbers, document.Pages.Count);
        var selectedSet = new HashSet<int>(selectedPages);
        var pageObjectNumbers = document.Pages.Select(page => page.ObjectNumber).ToArray();
        int? softMaskPseudoObjectNumber = imageStream.SoftMask is null ? null : ImageSoftMaskPseudoObjectNumber;
        var additionalObjects = new List<PdfPageExtractor.AdditionalObject>();
        if (imageStream.SoftMask != null) {
            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(ImageSoftMaskPseudoObjectNumber, PdfWriter.BuildImageXObject(imageStream.SoftMask)));
        }

        additionalObjects.Add(new PdfPageExtractor.AdditionalObject(ImagePseudoObjectNumber, PdfWriter.BuildImageXObject(imageStream, softMaskPseudoObjectNumber)));
        var overrides = new Dictionary<int, Dictionary<string, PdfObject>>();
        string imageResourceName = GetAvailableXObjectResourceName(objects, pageObjectNumbers);

        for (int i = 0; i < document.Pages.Count; i++) {
            int pageNumber = i + 1;
            if (!selectedSet.Contains(pageNumber)) {
                continue;
            }

            var page = document.Pages[i];
            var size = page.GetPageSize();
            int stampPseudoId = -2000 - pageNumber;
            var stampStream = BuildImageStampStream(
                imageResourceName,
                size.Width,
                size.Height,
                imageStream.PixelWidth,
                imageStream.PixelHeight,
                effectiveOptions,
                watermarkDefaults);

            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(stampPseudoId, stampStream));
            overrides[page.ObjectNumber] = BuildImagePageOverrides(objects, pageObjectNumbers[i], imageResourceName, stampPseudoId, effectiveOptions.BehindContent);
        }

        return PdfPageExtractor.ExtractPages(objects, document.Metadata, pageObjectNumbers, overrides, additionalObjects, PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw));
    }

    /// <summary>
    /// Writes a new PDF with an image stamp on selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static void StampImage(string inputPath, string outputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, StampImage(File.ReadAllBytes(inputPath), imageBytes, options));
    }

    /// <summary>
    /// Adds an image stamp to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampImage(string inputPath, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, StampImage(File.ReadAllBytes(inputPath), imageBytes, options));
    }

    /// <summary>
    /// Writes a new PDF with an image stamp from the current position of a readable image stream.
    /// </summary>
    public static void StampImage(string inputPath, string outputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, StampImage(File.ReadAllBytes(inputPath), imageStream, options));
    }

    /// <summary>
    /// Adds an image stamp from the current position of a readable image stream to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampImage(string inputPath, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, StampImage(File.ReadAllBytes(inputPath), imageStream, options));
    }

    /// <summary>
    /// Writes a new PDF with an image stamp from the current position of a readable PDF stream.
    /// </summary>
    public static void StampImage(Stream stream, string outputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, StampImage(stream, imageBytes, options));
    }

    /// <summary>
    /// Writes a new PDF with an image stamp from the current position of readable PDF and image streams.
    /// </summary>
    public static void StampImage(Stream stream, string outputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, StampImage(stream, imageStream, options));
    }

    /// <summary>
    /// Adds an image stamp to selected pages from a PDF file and returns the stamped PDF bytes.
    /// </summary>
    public static byte[] StampImageToBytes(string inputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return StampImage(File.ReadAllBytes(inputPath), imageBytes, options);
    }

    /// <summary>
    /// Adds an image stamp from the current position of a readable image stream to selected pages from a PDF file and returns the stamped PDF bytes.
    /// </summary>
    public static byte[] StampImageToBytes(string inputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return StampImage(File.ReadAllBytes(inputPath), imageStream, options);
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages, or every page when no page selection is supplied.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static byte[] WatermarkImage(byte[] pdf, byte[] imageBytes, PdfImageStampOptions? options = null) {
        var effectiveOptions = BuildImageWatermarkOptions(options);
        return StampImageCore(pdf, imageBytes, effectiveOptions, watermarkDefaults: true);
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable image stream to selected pages, or every page when no page selection is supplied.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static byte[] WatermarkImage(byte[] pdf, Stream imageStream, PdfImageStampOptions? options = null) {
        return WatermarkImage(pdf, ReadStream(imageStream, nameof(imageStream)), options);
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages from the current position of a readable stream, or every page when no page selection is supplied.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static byte[] WatermarkImage(Stream stream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        return WatermarkImage(ReadStream(stream, nameof(stream)), imageBytes, options);
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of readable PDF and image streams to selected pages, or every page when no page selection is supplied.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static byte[] WatermarkImage(Stream stream, Stream imageStream, PdfImageStampOptions? options = null) {
        return WatermarkImage(ReadStream(stream, nameof(stream)), ReadStream(imageStream, nameof(imageStream)), options);
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static void WatermarkImage(byte[] pdf, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        WriteOutput(outputStream, WatermarkImage(pdf, imageBytes, options));
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable image stream and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static void WatermarkImage(byte[] pdf, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        WriteOutput(outputStream, WatermarkImage(pdf, imageStream, options));
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable PDF stream and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static void WatermarkImage(Stream stream, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        WriteOutput(outputStream, WatermarkImage(stream, imageBytes, options));
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of readable PDF and image streams and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static void WatermarkImage(Stream stream, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        WriteOutput(outputStream, WatermarkImage(stream, imageStream, options));
    }

    /// <summary>
    /// Writes a new PDF with a centered image watermark on selected pages, or every page when no page selection is supplied.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static void WatermarkImage(string inputPath, string outputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkImage(File.ReadAllBytes(inputPath), imageBytes, options));
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static void WatermarkImage(string inputPath, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, WatermarkImage(File.ReadAllBytes(inputPath), imageBytes, options));
    }

    /// <summary>
    /// Writes a new PDF with a centered image watermark from the current position of a readable image stream.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static void WatermarkImage(string inputPath, string outputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkImage(File.ReadAllBytes(inputPath), imageStream, options));
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable image stream to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static void WatermarkImage(string inputPath, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, WatermarkImage(File.ReadAllBytes(inputPath), imageStream, options));
    }

    /// <summary>
    /// Writes a new PDF with a centered image watermark from the current position of a readable PDF stream.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static void WatermarkImage(Stream stream, string outputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkImage(stream, imageBytes, options));
    }

    /// <summary>
    /// Writes a new PDF with a centered image watermark from the current position of readable PDF and image streams.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static void WatermarkImage(Stream stream, string outputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkImage(stream, imageStream, options));
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages from a PDF file and returns the watermarked PDF bytes.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static byte[] WatermarkImageToBytes(string inputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return WatermarkImage(File.ReadAllBytes(inputPath), imageBytes, options);
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable image stream to selected pages from a PDF file and returns the watermarked PDF bytes.
    /// Simple PNG alpha soft masks are supported for grayscale-alpha/RGBA inputs.
    /// </summary>
    public static byte[] WatermarkImageToBytes(string inputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return WatermarkImage(File.ReadAllBytes(inputPath), imageStream, options);
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

    private static Dictionary<string, PdfObject> BuildPageOverrides(
        Dictionary<int, PdfIndirectObject> objects,
        int pageObjectNumber,
        string fontResourceName,
        int stampPseudoObjectNumber,
        bool behindContent) {
        if (!objects.TryGetValue(pageObjectNumber, out var indirect) || indirect.Value is not PdfDictionary pageDictionary) {
            throw new InvalidOperationException("PDF page object " + pageObjectNumber.ToString(CultureInfo.InvariantCulture) + " was not found.");
        }

        var contents = BuildContentsArray(objects, pageDictionary.Items.TryGetValue("Contents", out var contentsObj) ? contentsObj : null, stampPseudoObjectNumber, behindContent);
        var resources = BuildResourcesDictionary(objects, GetInheritedPageValue(objects, pageDictionary, "Resources"), fontResourceName);

        return new Dictionary<string, PdfObject>(StringComparer.Ordinal) {
            ["Contents"] = contents,
            ["Resources"] = resources
        };
    }

    private static Dictionary<string, PdfObject> BuildImagePageOverrides(
        Dictionary<int, PdfIndirectObject> objects,
        int pageObjectNumber,
        string imageResourceName,
        int stampPseudoObjectNumber,
        bool behindContent) {
        if (!objects.TryGetValue(pageObjectNumber, out var indirect) || indirect.Value is not PdfDictionary pageDictionary) {
            throw new InvalidOperationException("PDF page object " + pageObjectNumber.ToString(CultureInfo.InvariantCulture) + " was not found.");
        }

        var contents = BuildContentsArray(objects, pageDictionary.Items.TryGetValue("Contents", out var contentsObj) ? contentsObj : null, stampPseudoObjectNumber, behindContent);
        var resources = BuildImageResourcesDictionary(objects, GetInheritedPageValue(objects, pageDictionary, "Resources"), imageResourceName);

        return new Dictionary<string, PdfObject>(StringComparer.Ordinal) {
            ["Contents"] = contents,
            ["Resources"] = resources
        };
    }

    private static PdfArray BuildContentsArray(Dictionary<int, PdfIndirectObject> objects, PdfObject? existingContents, int stampPseudoObjectNumber, bool behindContent) {
        var result = new PdfArray();
        var stampReference = new PdfReference(stampPseudoObjectNumber, 0);

        if (behindContent) {
            result.Items.Add(stampReference);
        }

        AppendContentEntries(objects, result, existingContents);

        if (!behindContent) {
            result.Items.Add(stampReference);
        }

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
            objects.TryGetValue(reference.ObjectNumber, out var indirect) &&
            indirect.Value is PdfArray referencedArray) {
            foreach (var item in referencedArray.Items) {
                target.Items.Add(item);
            }

            return;
        }

        target.Items.Add(contents);
    }

    private static PdfDictionary BuildResourcesDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? existingResources,
        string fontResourceName) {
        var resources = CloneDictionary(ResolveDictionary(objects, existingResources));
        var fonts = CloneDictionary(ResolveDictionary(objects, resources.Items.TryGetValue("Font", out var fontObj) ? fontObj : null));
        fonts.Items[fontResourceName] = new PdfReference(FontPseudoObjectNumber, 0);
        resources.Items["Font"] = fonts;
        return resources;
    }

    private static PdfDictionary BuildImageResourcesDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? existingResources,
        string imageResourceName) {
        var resources = CloneDictionary(ResolveDictionary(objects, existingResources));
        var xObjects = CloneDictionary(ResolveDictionary(objects, resources.Items.TryGetValue("XObject", out var xObjectObj) ? xObjectObj : null));
        xObjects.Items[imageResourceName] = new PdfReference(ImagePseudoObjectNumber, 0);
        resources.Items["XObject"] = xObjects;
        return resources;
    }

    private static PdfDictionary CloneDictionary(PdfDictionary? source) {
        var clone = new PdfDictionary();
        if (source is null) {
            return clone;
        }

        foreach (var entry in source.Items) {
            clone.Items[entry.Key] = entry.Value;
        }

        return clone;
    }

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? obj) {
        if (obj is PdfDictionary dictionary) {
            return dictionary;
        }

        if (obj is PdfReference reference &&
            objects.TryGetValue(reference.ObjectNumber, out var indirect) &&
            indirect.Value is PdfDictionary referencedDictionary) {
            return referencedDictionary;
        }

        return null;
    }

    private static string GetAvailableFontResourceName(Dictionary<int, PdfIndirectObject> objects, int[] pageObjectNumbers) {
        var usedNames = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < pageObjectNumbers.Length; i++) {
            if (!objects.TryGetValue(pageObjectNumbers[i], out var indirect) || indirect.Value is not PdfDictionary pageDictionary) {
                continue;
            }

            var resources = ResolveDictionary(objects, GetInheritedPageValue(objects, pageDictionary, "Resources"));
            var fonts = ResolveDictionary(objects, resources?.Items.TryGetValue("Font", out var fontObj) == true ? fontObj : null);
            if (fonts is null) {
                continue;
            }

            foreach (string name in fonts.Items.Keys) {
                usedNames.Add(name);
            }
        }

        const string baseName = "OIMOStampF";
        for (int i = 1; i < 1000; i++) {
            string candidate = baseName + i.ToString(CultureInfo.InvariantCulture);
            if (!usedNames.Contains(candidate)) {
                return candidate;
            }
        }

        throw new InvalidOperationException("Unable to find an available PDF font resource name for the stamp.");
    }

    private static string GetAvailableXObjectResourceName(Dictionary<int, PdfIndirectObject> objects, int[] pageObjectNumbers) {
        var usedNames = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < pageObjectNumbers.Length; i++) {
            if (!objects.TryGetValue(pageObjectNumbers[i], out var indirect) || indirect.Value is not PdfDictionary pageDictionary) {
                continue;
            }

            var resources = ResolveDictionary(objects, GetInheritedPageValue(objects, pageDictionary, "Resources"));
            var xObjects = ResolveDictionary(objects, resources?.Items.TryGetValue("XObject", out var xObjectObj) == true ? xObjectObj : null);
            if (xObjects is null) {
                continue;
            }

            foreach (string name in xObjects.Items.Keys) {
                usedNames.Add(name);
            }
        }

        const string baseName = "OIMOStampIm";
        for (int i = 1; i < 1000; i++) {
            string candidate = baseName + i.ToString(CultureInfo.InvariantCulture);
            if (!usedNames.Contains(candidate)) {
                return candidate;
            }
        }

        throw new InvalidOperationException("Unable to find an available PDF image resource name for the stamp.");
    }

    private static PdfObject? GetInheritedPageValue(Dictionary<int, PdfIndirectObject> objects, PdfDictionary pageDictionary, string key) {
        PdfDictionary? current = pageDictionary;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (current.Items.TryGetValue(key, out var value)) {
                return value;
            }

            if (!current.Items.TryGetValue("Parent", out var parentObj) ||
                parentObj is not PdfReference parentReference ||
                !objects.TryGetValue(parentReference.ObjectNumber, out var parentIndirect) ||
                parentIndirect.Value is not PdfDictionary parentDictionary) {
                return null;
            }

            current = parentDictionary;
        }

        return null;
    }

    private static PdfDictionary BuildFontObject(PdfStandardFont font) {
        return PdfStandardFontDictionaryBuilder.BuildStandardType1FontDictionary(font);
    }

    private static PdfStream BuildStampStream(
        string text,
        string fontResourceName,
        double pageWidth,
        double pageHeight,
        PdfTextStampOptions options,
        bool watermarkDefaults) {
        double fontSize = options.FontSize;
        double x = options.X ?? (watermarkDefaults ? (pageWidth - PdfWriter.EstimateSimpleTextWidth(text, options.Font, fontSize)) / 2.0 : 36);
        double y = options.Y ?? (watermarkDefaults ? pageHeight / 2.0 : 36);
        double radians = options.RotationDegrees * Math.PI / 180.0;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);

        var sb = new StringBuilder();
        new ContentStreamBuilder(sb)
            .SaveState()
            .FillColor(options.Color)
            .BeginText()
            .Font(fontResourceName, fontSize)
            .TextMatrix(cos, sin, -sin, cos, x, y)
            .ShowHexText(EncodeWinAnsiHex(text))
            .EndText()
            .RestoreState();

        return new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(sb.ToString()));
    }

    private static PdfStream BuildImageStampStream(
        string imageResourceName,
        double pageWidth,
        double pageHeight,
        int pixelWidth,
        int pixelHeight,
        PdfImageStampOptions options,
        bool watermarkDefaults) {
        double imageWidth = options.Width ?? pixelWidth;
        double imageHeight = options.Height ?? pixelHeight;
        double x = options.X ?? (watermarkDefaults ? (pageWidth - imageWidth) / 2.0 : 36);
        double y = options.Y ?? (watermarkDefaults ? (pageHeight - imageHeight) / 2.0 : 36);
        double radians = options.RotationDegrees * Math.PI / 180.0;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);

        var sb = new StringBuilder();
        new ContentStreamBuilder(sb)
            .SaveState()
            .TransformMatrix(imageWidth * cos, imageWidth * sin, -imageHeight * sin, imageHeight * cos, x, y)
            .XObject(imageResourceName)
            .RestoreState();

        return new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(sb.ToString()));
    }

    private static int[] NormalizePageNumbers(int[]? pageNumbers, int pageCount) {
        if (pageNumbers is null || pageNumbers.Length == 0) {
            return Enumerable.Range(1, pageCount).ToArray();
        }

        var seen = new HashSet<int>();
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            if (pageNumber < 1 || pageNumber > pageCount) {
                throw new ArgumentOutOfRangeException(nameof(pageNumbers), "Page number " + pageNumber.ToString(CultureInfo.InvariantCulture) + " is outside the document page range 1-" + pageCount.ToString(CultureInfo.InvariantCulture) + ".");
            }

            if (!seen.Add(pageNumber)) {
                throw new ArgumentException("Duplicate page selections are not supported.", nameof(pageNumbers));
            }
        }

        return pageNumbers;
    }

    private static void ValidateOptions(PdfTextStampOptions options) {
        if (options.FontSize <= 0 || double.IsNaN(options.FontSize) || double.IsInfinity(options.FontSize)) {
            throw new ArgumentOutOfRangeException(nameof(options), "Font size must be a positive finite value.");
        }

        if ((options.X.HasValue && (double.IsNaN(options.X.Value) || double.IsInfinity(options.X.Value))) ||
            (options.Y.HasValue && (double.IsNaN(options.Y.Value) || double.IsInfinity(options.Y.Value))) ||
            double.IsNaN(options.RotationDegrees) ||
            double.IsInfinity(options.RotationDegrees)) {
            throw new ArgumentOutOfRangeException(nameof(options), "Text stamp coordinates and rotation must be finite.");
        }
    }

    private static PdfImageStampOptions BuildImageWatermarkOptions(PdfImageStampOptions? options) {
        if (options is null) {
            return new PdfImageStampOptions {
                BehindContent = true
            };
        }

        var effective = new PdfImageStampOptions {
            PageNumbers = options.PageNumbers,
            X = options.X,
            Y = options.Y,
            Width = options.Width,
            Height = options.Height,
            RotationDegrees = options.RotationDegrees,
            BehindContent = true
        };

        ValidateImageOptions(effective);
        return effective;
    }

    private static void ValidateImageOptions(PdfImageStampOptions options) {
        if (options.Width.HasValue && (options.Width.Value <= 0 || double.IsNaN(options.Width.Value) || double.IsInfinity(options.Width.Value))) {
            throw new ArgumentOutOfRangeException(nameof(options), "Image stamp width must be a positive finite value.");
        }

        if (options.Height.HasValue && (options.Height.Value <= 0 || double.IsNaN(options.Height.Value) || double.IsInfinity(options.Height.Value))) {
            throw new ArgumentOutOfRangeException(nameof(options), "Image stamp height must be a positive finite value.");
        }

        if ((options.X.HasValue && (double.IsNaN(options.X.Value) || double.IsInfinity(options.X.Value))) ||
            (options.Y.HasValue && (double.IsNaN(options.Y.Value) || double.IsInfinity(options.Y.Value))) ||
            double.IsNaN(options.RotationDegrees) ||
            double.IsInfinity(options.RotationDegrees)) {
            throw new ArgumentOutOfRangeException(nameof(options), "Image stamp coordinates and rotation must be finite.");
        }
    }

    private static string EncodeWinAnsiHex(string text) {
        var bytes = PdfWinAnsiEncoding.Encode(text);
        var sb = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            sb.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
        }

        return sb.ToString();
    }

    private static byte[] ReadStream(Stream stream, string paramName) {
        Guard.NotNull(stream, paramName);
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", paramName);
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static void WriteOutput(Stream outputStream, byte[] bytes) {
        ValidateWritableOutputStream(outputStream);

        outputStream.Write(bytes, 0, bytes.Length);
    }

    private static void ValidateWritableOutputStream(Stream outputStream) {
        Guard.NotNull(outputStream, nameof(outputStream));
        if (!outputStream.CanWrite) {
            throw new ArgumentException("Stream must be writable.", nameof(outputStream));
        }
    }

    private static void WriteOutput(string outputPath, byte[] bytes) {
        string fullPath = ValidateOutputPath(outputPath);
        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllBytes(fullPath, bytes);
    }

    private static string ValidateOutputPath(string outputPath) {
        Guard.NotNull(outputPath, nameof(outputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path cannot be empty or whitespace.", nameof(outputPath));
        }

        string fullPath;
        try {
            fullPath = Path.GetFullPath(outputPath);
        } catch (Exception ex) {
            throw new ArgumentException("Output path is invalid.", nameof(outputPath), ex);
        }

        if (Directory.Exists(fullPath) && (File.GetAttributes(fullPath) & FileAttributes.Directory) == FileAttributes.Directory) {
            throw new ArgumentException("Output path refers to a directory; a file path is required.", nameof(outputPath));
        }

        var fileName = Path.GetFileName(fullPath);
        if (string.IsNullOrEmpty(fileName)) {
            throw new ArgumentException("Output path must include a file name.", nameof(outputPath));
        }

        if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
            throw new ArgumentException("Output path contains invalid file name characters.", nameof(outputPath));
        }

        return fullPath;
    }
}
