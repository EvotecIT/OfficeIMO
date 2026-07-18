namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
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
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageContent);
        if (imageBytes.Length == 0) {
            throw new ArgumentException("Image bytes cannot be empty.", nameof(imageBytes));
        }

        var effectiveOptions = options ?? new PdfImageStampOptions();
        ValidateImageOptions(effectiveOptions);
        var imageInfo = PdfDocument.ValidateImageBytes(imageBytes);

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
        var document = PdfReadDocument.Open(pdf);
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

        PdfFileVersion fileVersion = PdfPageExtractor.GetSourceFileVersion(pdf);
        return PdfPageExtractor.ExtractPages(objects, document.Metadata, pageObjectNumbers, overrides, additionalObjects, PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw), fileVersion);
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
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static byte[] WatermarkImage(byte[] pdf, byte[] imageBytes, PdfImageStampOptions? options = null) {
        var effectiveOptions = BuildImageWatermarkOptions(options);
        return StampImageCore(pdf, imageBytes, effectiveOptions, watermarkDefaults: true);
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable image stream to selected pages, or every page when no page selection is supplied.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static byte[] WatermarkImage(byte[] pdf, Stream imageStream, PdfImageStampOptions? options = null) {
        return WatermarkImage(pdf, ReadStream(imageStream, nameof(imageStream)), options);
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages from the current position of a readable stream, or every page when no page selection is supplied.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static byte[] WatermarkImage(Stream stream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        return WatermarkImage(ReadStream(stream, nameof(stream)), imageBytes, options);
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of readable PDF and image streams to selected pages, or every page when no page selection is supplied.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static byte[] WatermarkImage(Stream stream, Stream imageStream, PdfImageStampOptions? options = null) {
        return WatermarkImage(ReadStream(stream, nameof(stream)), ReadStream(imageStream, nameof(imageStream)), options);
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(byte[] pdf, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        WriteOutput(outputStream, WatermarkImage(pdf, imageBytes, options));
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable image stream and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(byte[] pdf, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        WriteOutput(outputStream, WatermarkImage(pdf, imageStream, options));
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable PDF stream and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(Stream stream, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        WriteOutput(outputStream, WatermarkImage(stream, imageBytes, options));
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of readable PDF and image streams and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(Stream stream, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        WriteOutput(outputStream, WatermarkImage(stream, imageStream, options));
    }

    /// <summary>
    /// Writes a new PDF with a centered image watermark on selected pages, or every page when no page selection is supplied.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(string inputPath, string outputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkImage(File.ReadAllBytes(inputPath), imageBytes, options));
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(string inputPath, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, WatermarkImage(File.ReadAllBytes(inputPath), imageBytes, options));
    }

    /// <summary>
    /// Writes a new PDF with a centered image watermark from the current position of a readable image stream.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(string inputPath, string outputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkImage(File.ReadAllBytes(inputPath), imageStream, options));
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable image stream to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(string inputPath, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        WriteOutput(outputStream, WatermarkImage(File.ReadAllBytes(inputPath), imageStream, options));
    }

    /// <summary>
    /// Writes a new PDF with a centered image watermark from the current position of a readable PDF stream.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(Stream stream, string outputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkImage(stream, imageBytes, options));
    }

    /// <summary>
    /// Writes a new PDF with a centered image watermark from the current position of readable PDF and image streams.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(Stream stream, string outputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkImage(stream, imageStream, options));
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages from a PDF file and returns the watermarked PDF bytes.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static byte[] WatermarkImageToBytes(string inputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return WatermarkImage(File.ReadAllBytes(inputPath), imageBytes, options);
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable image stream to selected pages from a PDF file and returns the watermarked PDF bytes.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static byte[] WatermarkImageToBytes(string inputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return WatermarkImage(File.ReadAllBytes(inputPath), imageStream, options);
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

}
