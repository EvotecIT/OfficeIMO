namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    /// <summary>
    /// Adds an image stamp to selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] StampImage(byte[] pdf, byte[] imageBytes, PdfImageStampOptions? options = null, PdfLoadOptions? readOptions = null) {
        return StampImageCore(pdf, imageBytes, options, watermarkDefaults: false, readOptions);
    }

    /// <summary>
    /// Adds an image stamp from the current position of a readable image stream to selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] StampImage(byte[] pdf, Stream imageStream, PdfImageStampOptions? options = null) {
        return StampImage(pdf, ReadImageStream(imageStream, nameof(imageStream), options), options);
    }

    /// <summary>
    /// Adds an image stamp to selected pages from the current position of a readable stream, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] StampImage(Stream stream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        return StampImage(ReadStream(stream, nameof(stream), cancellationToken: options?.CancellationToken ?? default), imageBytes, options);
    }

    /// <summary>
    /// Adds an image stamp from the current position of readable PDF and image streams to selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static byte[] StampImage(Stream stream, Stream imageStream, PdfImageStampOptions? options = null) {
        return StampImage(
            ReadStream(stream, nameof(stream), cancellationToken: options?.CancellationToken ?? default),
            ReadImageStream(imageStream, nameof(imageStream), options),
            options);
    }

    /// <summary>
    /// Adds an image stamp to selected pages and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampImage(byte[] pdf, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, StampImage(pdf, imageBytes, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds an image stamp from the current position of a readable image stream and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampImage(byte[] pdf, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, StampImage(pdf, imageStream, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds an image stamp from the current position of a readable PDF stream and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampImage(Stream stream, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, StampImage(stream, imageBytes, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds an image stamp from the current position of readable PDF and image streams and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampImage(Stream stream, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, StampImage(stream, imageStream, options), options?.CancellationToken ?? default);
    }

    private static byte[] StampImageCore(byte[] pdf, byte[] imageBytes, PdfImageStampOptions? options, bool watermarkDefaults, PdfLoadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(imageBytes, nameof(imageBytes));
        var cancellationToken = options?.CancellationToken ?? default;
        cancellationToken.ThrowIfCancellationRequested();
        PdfMutationPlanner.RequireCatalogPreservingPageContentRewrite(pdf, readOptions, cancellationToken);
        if (imageBytes.Length == 0) {
            throw new ArgumentException("Image bytes cannot be empty.", nameof(imageBytes));
        }

        var effectiveOptions = options ?? new PdfImageStampOptions();
        PdfImageInput.EnsureWithinLimit(imageBytes.LongLength, effectiveOptions.MaximumEncodedImageBytes);
        ValidateImageOptions(effectiveOptions);
        PdfDocument.PreparedImage prepared = PdfDocument.PrepareImageDocumentSource(imageBytes, cancellationToken);

        if (!PdfWriter.TryBuildImageStream(
                prepared,
                effectiveOptions.Width ?? Math.Max(1, prepared.Info.Width),
                effectiveOptions.Height ?? Math.Max(1, prepared.Info.Height),
                out var imageStream,
                out string? unsupportedReason)) {
            throw new NotSupportedException(unsupportedReason ?? "Image format is not supported.");
        }

        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(pdf, readOptions, out _, out _, cancellationToken).Map;
        var document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        if (document.Pages.Count == 0) {
            throw new ArgumentException("PDF does not contain any pages.", nameof(pdf));
        }

        var selectedPages = NormalizePageNumbers(effectiveOptions.PageNumbers, document.Pages.Count);
        var selectedSet = new HashSet<int>(selectedPages);
        var pageObjectNumbers = document.Pages.Select(page => page.ObjectNumber).ToArray();
        string imageResourceName = GetAvailableXObjectResourceName(objects, pageObjectNumbers);
        return PdfDocumentObjectGraphRewriter.Rewrite(pdf, readOptions, null, (rewrittenObjects, security) => {
            int nextObjectNumber = rewrittenObjects.Count == 0 ? 1 : rewrittenObjects.Keys.Max() + 1;
            int imageObjectNumber = PdfWriter.AddImageXObjectResource(imageStream, stream => {
                int number = nextObjectNumber++;
                rewrittenObjects[number] = new PdfIndirectObject(number, 0, stream);
                return number;
            }, PdfFileAssembler.ResolveFileVersion(PdfSyntax.GetHeaderVersion(pdf), document.CatalogVersion));

            for (int i = 0; i < document.Pages.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                int pageNumber = i + 1;
                if (!selectedSet.Contains(pageNumber)) continue;
                var size = document.Pages[i].GetPageSize();
                if (!effectiveOptions.BehindContent) {
                    IsolateExistingContents(rewrittenObjects, pageObjectNumbers[i], ref nextObjectNumber);
                }
                int stampObjectNumber = nextObjectNumber++;
                rewrittenObjects[stampObjectNumber] = new PdfIndirectObject(
                    stampObjectNumber,
                    0,
                    BuildImageStampStream(
                        imageResourceName,
                        size.Width,
                        size.Height,
                        imageStream.LayoutPixelWidth,
                        imageStream.LayoutPixelHeight,
                        imageStream.Orientation,
                        effectiveOptions,
                        watermarkDefaults));
                Dictionary<string, PdfObject> pageOverrides = BuildImagePageOverrides(
                    rewrittenObjects,
                    pageObjectNumbers[i],
                    imageResourceName,
                    stampObjectNumber,
                    effectiveOptions.BehindContent,
                    imageObjectNumber);
                PdfDictionary pageDictionary = (PdfDictionary)rewrittenObjects[pageObjectNumbers[i]].Value;
                foreach (KeyValuePair<string, PdfObject> item in pageOverrides) pageDictionary.Items[item.Key] = item.Value;
            }

            return security.InfoObjectNumber.HasValue && rewrittenObjects.ContainsKey(security.InfoObjectNumber.Value)
                ? security.InfoObjectNumber
                : null;
        }, cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Writes a new PDF with an image stamp on selected pages, or every page when no page selection is supplied.
    /// </summary>
    public static void StampImage(string inputPath, string outputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, StampImage(ReadPath(inputPath, options?.CancellationToken ?? default), imageBytes, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds an image stamp to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampImage(string inputPath, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, StampImage(ReadPath(inputPath, options?.CancellationToken ?? default), imageBytes, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Writes a new PDF with an image stamp from the current position of a readable image stream.
    /// </summary>
    public static void StampImage(string inputPath, string outputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, StampImage(ReadPath(inputPath, options?.CancellationToken ?? default), imageStream, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds an image stamp from the current position of a readable image stream to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void StampImage(string inputPath, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, StampImage(ReadPath(inputPath, options?.CancellationToken ?? default), imageStream, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Writes a new PDF with an image stamp from the current position of a readable PDF stream.
    /// </summary>
    public static void StampImage(Stream stream, string outputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, StampImage(stream, imageBytes, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Writes a new PDF with an image stamp from the current position of readable PDF and image streams.
    /// </summary>
    public static void StampImage(Stream stream, string outputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, StampImage(stream, imageStream, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds an image stamp to selected pages from a PDF file and returns the stamped PDF bytes.
    /// </summary>
    public static byte[] StampImageToBytes(string inputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return StampImage(ReadPath(inputPath, options?.CancellationToken ?? default), imageBytes, options);
    }

    /// <summary>
    /// Adds an image stamp from the current position of a readable image stream to selected pages from a PDF file and returns the stamped PDF bytes.
    /// </summary>
    public static byte[] StampImageToBytes(string inputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return StampImage(ReadPath(inputPath, options?.CancellationToken ?? default), imageStream, options);
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages, or every page when no page selection is supplied.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static byte[] WatermarkImage(byte[] pdf, byte[] imageBytes, PdfImageStampOptions? options = null, PdfLoadOptions? readOptions = null) {
        var effectiveOptions = BuildImageWatermarkOptions(options);
        return StampImageCore(pdf, imageBytes, effectiveOptions, watermarkDefaults: true, readOptions);
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable image stream to selected pages, or every page when no page selection is supplied.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static byte[] WatermarkImage(byte[] pdf, Stream imageStream, PdfImageStampOptions? options = null) {
        return WatermarkImage(pdf, ReadImageStream(imageStream, nameof(imageStream), options), options);
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages from the current position of a readable stream, or every page when no page selection is supplied.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static byte[] WatermarkImage(Stream stream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        return WatermarkImage(ReadStream(stream, nameof(stream), cancellationToken: options?.CancellationToken ?? default), imageBytes, options);
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of readable PDF and image streams to selected pages, or every page when no page selection is supplied.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static byte[] WatermarkImage(Stream stream, Stream imageStream, PdfImageStampOptions? options = null) {
        return WatermarkImage(
            ReadStream(stream, nameof(stream), cancellationToken: options?.CancellationToken ?? default),
            ReadImageStream(imageStream, nameof(imageStream), options),
            options);
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(byte[] pdf, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, WatermarkImage(pdf, imageBytes, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable image stream and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(byte[] pdf, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, WatermarkImage(pdf, imageStream, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable PDF stream and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(Stream stream, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, WatermarkImage(stream, imageBytes, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of readable PDF and image streams and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(Stream stream, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, WatermarkImage(stream, imageStream, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Writes a new PDF with a centered image watermark on selected pages, or every page when no page selection is supplied.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(string inputPath, string outputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkImage(ReadPath(inputPath, options?.CancellationToken ?? default), imageBytes, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(string inputPath, Stream outputStream, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, WatermarkImage(ReadPath(inputPath, options?.CancellationToken ?? default), imageBytes, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Writes a new PDF with a centered image watermark from the current position of a readable image stream.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(string inputPath, string outputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkImage(ReadPath(inputPath, options?.CancellationToken ?? default), imageStream, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable image stream to selected pages from a PDF file and writes the result to <paramref name="outputStream"/>.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(string inputPath, Stream outputStream, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        ValidateWritableOutputStream(outputStream);
        WriteOutput(outputStream, WatermarkImage(ReadPath(inputPath, options?.CancellationToken ?? default), imageStream, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Writes a new PDF with a centered image watermark from the current position of a readable PDF stream.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(Stream stream, string outputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkImage(stream, imageBytes, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Writes a new PDF with a centered image watermark from the current position of readable PDF and image streams.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static void WatermarkImage(Stream stream, string outputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteOutput(fullOutputPath, WatermarkImage(stream, imageStream, options), options?.CancellationToken ?? default);
    }

    /// <summary>
    /// Adds a centered image watermark to selected pages from a PDF file and returns the watermarked PDF bytes.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static byte[] WatermarkImageToBytes(string inputPath, byte[] imageBytes, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return WatermarkImage(ReadPath(inputPath, options?.CancellationToken ?? default), imageBytes, options);
    }

    /// <summary>
    /// Adds a centered image watermark from the current position of a readable image stream to selected pages from a PDF file and returns the watermarked PDF bytes.
    /// Simple PNG alpha and transparency soft masks are supported for compatible PNG inputs.
    /// </summary>
    public static byte[] WatermarkImageToBytes(string inputPath, Stream imageStream, PdfImageStampOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));

        return WatermarkImage(ReadPath(inputPath, options?.CancellationToken ?? default), imageStream, options);
    }

    private static PdfImageStampOptions BuildImageWatermarkOptions(PdfImageStampOptions? options) {
        if (options is null) {
            return new PdfImageStampOptions {
                BehindContent = true
            };
        }

        var effective = new PdfImageStampOptions {
            CancellationToken = options.CancellationToken,
            PageNumbers = options.PageNumbers,
            X = options.X,
            Y = options.Y,
            Width = options.Width,
            Height = options.Height,
            RotationDegrees = options.RotationDegrees,
            MaximumEncodedImageBytes = options.MaximumEncodedImageBytes,
            BehindContent = true
        };

        ValidateImageOptions(effective);
        return effective;
    }

}
