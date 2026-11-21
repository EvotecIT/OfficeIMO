using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Dependency-free rendered PDF comparison with structural evidence and review artifacts.</summary>
public static class PdfVisualComparer {
    /// <summary>Compares whole documents, a shared page set, or independent ordered page selections.</summary>
    public static PdfVisualComparisonReport Compare(
        byte[] expectedPdf,
        byte[] actualPdf,
        PdfPageSelection? selection = null,
        PdfVisualComparisonOptions? options = null,
        PdfLoadOptions? expectedReadOptions = null,
        PdfLoadOptions? actualReadOptions = null) =>
        Compare(expectedPdf, actualPdf, CancellationToken.None, selection, options, expectedReadOptions, actualReadOptions);

    /// <summary>Compares whole documents, a shared page set, or independent ordered page selections with cooperative cancellation.</summary>
    public static PdfVisualComparisonReport Compare(
        byte[] expectedPdf,
        byte[] actualPdf,
        CancellationToken cancellationToken,
        PdfPageSelection? selection = null,
        PdfVisualComparisonOptions? options = null,
        PdfLoadOptions? expectedReadOptions = null,
        PdfLoadOptions? actualReadOptions = null) {
        Guard.NotNull(expectedPdf, nameof(expectedPdf));
        Guard.NotNull(actualPdf, nameof(actualPdf));
        PdfVisualComparisonOptions effectiveOptions = options ?? new PdfVisualComparisonOptions();
        effectiveOptions.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        PdfReadDocument expected = PdfReadDocument.Open(expectedPdf, expectedReadOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PdfReadDocument actual = PdfReadDocument.Open(actualPdf, actualReadOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        var structural = new List<string>();
        if (expected.Pages.Count != actual.Pages.Count) {
            structural.Add("PageCount: expected " + expected.Pages.Count + ", actual " + actual.Pages.Count + ".");
        }

        if (selection is not null && (effectiveOptions.ExpectedPages is not null || effectiveOptions.ActualPages is not null))
            throw new ArgumentException("Choose the shared page selection or independent expected/actual selectors, not both.", nameof(selection));
        bool independentSelection = effectiveOptions.ExpectedPages is not null || effectiveOptions.ActualPages is not null;
        int[] expectedPages = ResolvePages(effectiveOptions.ExpectedPages, selection, expected.Pages.Count, effectiveOptions.MaxPages);
        int[] actualPages = selection is null
            ? ResolvePages(effectiveOptions.ActualPages, null, actual.Pages.Count, effectiveOptions.MaxPages)
            : expectedPages.Where(page => page <= actual.Pages.Count).ToArray();
        if (independentSelection) structural.Clear(); // Document totals remain context; selected sequences define the comparison scope.
        int pairCount = selection is null ? Math.Min(expectedPages.Length, actualPages.Length) : actualPages.Length;
        int[] unmatchedExpected = selection is null ? expectedPages.Skip(pairCount).ToArray()
            : expectedPages.Where(page => page > actual.Pages.Count).ToArray();
        int[] unmatchedActual = selection is null ? actualPages.Skip(pairCount).ToArray() : Array.Empty<int>();
        foreach (int page in unmatchedExpected) structural.Add("Expected page " + page + " has no selected actual partner.");
        foreach (int page in unmatchedActual) structural.Add("Actual page " + page + " has no selected expected partner.");
        var pages = new List<PdfVisualPageComparison>(pairCount);
        long totalPixels = 0;
        long totalOutputBytes = 0;
        for (int i = 0; i < pairCount; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            int expectedPageNumber = selection is null ? expectedPages[i] : actualPages[i];
            PdfVisualPageComparison page = ComparePage(expected, actual, expectedPageNumber, actualPages[i],
                effectiveOptions, structural, ref totalPixels, cancellationToken);
            totalOutputBytes = checked(totalOutputBytes + page.OutputByteLength);
            if (totalOutputBytes > effectiveOptions.MaxTotalOutputBytes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.RenderBytes, effectiveOptions.MaxTotalOutputBytes, totalOutputBytes);
            }
            pages.Add(page);
        }
        return new PdfVisualComparisonReport(pages.AsReadOnly(), structural.AsReadOnly(), expected.Pages.Count, actual.Pages.Count,
            expectedPages, actualPages, unmatchedExpected, unmatchedActual,
            PdfArtifactInfo.CaptureKnownPageCount(expectedPdf, expected.Pages.Count, cancellationToken).Sha256,
            PdfArtifactInfo.CaptureKnownPageCount(actualPdf, actual.Pages.Count, cancellationToken).Sha256, independentSelection || selection is not null);
    }

    private static int[] ResolvePages(OfficePageSelector? selector, PdfPageSelection? selection, int count, int maximum) {
        if (selector is not null) return selector.Resolve(count, maximum).ToArray();
        long selectedCount = selection is null ? count : selection.Ranges.Sum(range => (long)range.LastPage - range.FirstPage + 1);
        if (selectedCount > maximum) throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, maximum, selectedCount);
        return selection?.ToPageNumbers(count, nameof(selection)) ?? Enumerable.Range(1, count).ToArray();
    }

    /// <summary>Compares any two one-based pages, including pages aligned after insertion or reordering.</summary>
    public static PdfVisualPageComparison ComparePages(
        byte[] expectedPdf,
        int expectedPageNumber,
        byte[] actualPdf,
        int actualPageNumber,
        PdfVisualComparisonOptions? options = null,
        PdfLoadOptions? expectedReadOptions = null,
        PdfLoadOptions? actualReadOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(expectedPdf, nameof(expectedPdf));
        Guard.NotNull(actualPdf, nameof(actualPdf));
        PdfVisualComparisonOptions effective = options ?? new PdfVisualComparisonOptions();
        effective.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        PdfReadDocument expected = PdfReadDocument.Open(expectedPdf, expectedReadOptions, cancellationToken);
        PdfReadDocument actual = PdfReadDocument.Open(actualPdf, actualReadOptions, cancellationToken);
        long totalPixels = 0;
        PdfVisualPageComparison page = ComparePages(expected, expectedPageNumber, actual, actualPageNumber, effective, ref totalPixels, cancellationToken);
        if (page.OutputByteLength > effective.MaxTotalOutputBytes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.RenderBytes, effective.MaxTotalOutputBytes, page.OutputByteLength);
        }
        return page;
    }

    internal static PdfVisualPageComparison ComparePages(
        PdfReadDocument expected, int expectedPageNumber,
        PdfReadDocument actual, int actualPageNumber,
        PdfVisualComparisonOptions options, ref long totalPixels,
        CancellationToken cancellationToken) {
        if (expectedPageNumber < 1 || expectedPageNumber > expected.Pages.Count) throw new ArgumentOutOfRangeException(nameof(expectedPageNumber));
        if (actualPageNumber < 1 || actualPageNumber > actual.Pages.Count) throw new ArgumentOutOfRangeException(nameof(actualPageNumber));
        var structural = new List<string>();
        PdfVisualPageComparison page = ComparePage(expected, actual, expectedPageNumber, actualPageNumber, options, structural, ref totalPixels, cancellationToken);
        return page;
    }

    private static PdfVisualPageComparison ComparePage(PdfReadDocument expectedDocument, PdfReadDocument actualDocument, int expectedPageNumber, int actualPageNumber, PdfVisualComparisonOptions options, List<string> structural, ref long totalPixels, CancellationToken cancellationToken) {
        IReadOnlyList<PdfRenderCapabilityDiagnostic> expectedDiagnostics = expectedDocument.Pages[expectedPageNumber - 1]
            .GetRenderCapabilityDiagnostics(cancellationToken);
        IReadOnlyList<PdfRenderCapabilityDiagnostic> actualDiagnostics = actualDocument.Pages[actualPageNumber - 1]
            .GetRenderCapabilityDiagnostics(cancellationToken);
        bool incomplete = PdfRenderCapabilities.HasIncompleteVisualProjection(expectedDiagnostics) ||
            PdfRenderCapabilities.HasIncompleteVisualProjection(actualDiagnostics);
        OfficeDrawing expectedDrawing = PdfPageImageRenderer.RenderPage(expectedDocument, expectedPageNumber, cancellationToken);
        OfficeDrawing actualDrawing = PdfPageImageRenderer.RenderPage(actualDocument, actualPageNumber, cancellationToken);
        AddPixelBudget(expectedDrawing.Width, expectedDrawing.Height, options.Scale, options, ref totalPixels);
        AddPixelBudget(actualDrawing.Width, actualDrawing.Height, options.Scale, options, ref totalPixels);
        cancellationToken.ThrowIfCancellationRequested();
        var rasterOptions = new OfficeDrawingRasterRenderOptions {
            Scale = options.Scale,
            Background = options.Background,
            MaximumRasterPixels = options.MaxPixelsPerImage,
            CancellationToken = cancellationToken
        };
        OfficeRasterImage expectedImage = OfficeDrawingRasterRenderer.Render(expectedDrawing, rasterOptions);
        OfficeRasterImage actualImage = OfficeDrawingRasterRenderer.Render(actualDrawing, rasterOptions);
        return CompareRasters(expectedImage, actualImage, null, null, expectedPageNumber, actualPageNumber, incomplete,
            options, structural, ref totalPixels, expectedDiagnostics, actualDiagnostics, cancellationToken);
    }

    /// <summary>Compares page images rendered earlier, such as approved snapshot baselines, with newly rendered page images.</summary>
    /// <remarks>
    /// Images pair by position and page numbers follow that order. Tolerances, ignored regions, alignment, background
    /// and limits apply; <see cref="PdfVisualComparisonOptions.Scale"/> does not, because the inputs are already rasterized.
    /// The report's SHA-256 identities describe each ordered PNG sequence rather than a PDF file.
    /// </remarks>
    public static PdfVisualComparisonReport CompareImages(
        IReadOnlyList<byte[]> expectedPagePngs,
        IReadOnlyList<byte[]> actualPagePngs,
        PdfVisualComparisonOptions? options = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(expectedPagePngs, nameof(expectedPagePngs));
        Guard.NotNull(actualPagePngs, nameof(actualPagePngs));
        PdfVisualComparisonOptions effective = options ?? new PdfVisualComparisonOptions();
        effective.Validate();
        if (effective.ExpectedPages is not null || effective.ActualPages is not null)
            throw new ArgumentException("Page selectors apply to PDF inputs; pass only the page images to compare.", nameof(options));
        int largest = Math.Max(expectedPagePngs.Count, actualPagePngs.Count);
        if (largest > effective.MaxPages) throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, effective.MaxPages, largest);
        var structural = new List<string>();
        if (expectedPagePngs.Count != actualPagePngs.Count) {
            structural.Add("PageCount: expected " + expectedPagePngs.Count + ", actual " + actualPagePngs.Count + ".");
        }

        int pairCount = Math.Min(expectedPagePngs.Count, actualPagePngs.Count);
        int[] expectedPages = Enumerable.Range(1, expectedPagePngs.Count).ToArray();
        int[] actualPages = Enumerable.Range(1, actualPagePngs.Count).ToArray();
        int[] unmatchedExpected = expectedPages.Skip(pairCount).ToArray();
        int[] unmatchedActual = actualPages.Skip(pairCount).ToArray();
        foreach (int page in unmatchedExpected) structural.Add("Expected page " + page + " has no selected actual partner.");
        foreach (int page in unmatchedActual) structural.Add("Actual page " + page + " has no selected expected partner.");
        var pages = new List<PdfVisualPageComparison>(pairCount);
        long totalPixels = 0;
        long totalOutputBytes = 0;
        for (int index = 0; index < pairCount; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            byte[] expectedPng = expectedPagePngs[index] ?? throw new ArgumentException("Page images cannot be null.", nameof(expectedPagePngs));
            byte[] actualPng = actualPagePngs[index] ?? throw new ArgumentException("Page images cannot be null.", nameof(actualPagePngs));
            OfficeRasterImage expectedImage = DecodePagePng(expectedPng, index + 1, nameof(expectedPagePngs), effective, ref totalPixels);
            OfficeRasterImage actualImage = DecodePagePng(actualPng, index + 1, nameof(actualPagePngs), effective, ref totalPixels);
            PdfVisualPageComparison page = CompareRasters(expectedImage, actualImage, expectedPng, actualPng, index + 1, index + 1,
                incomplete: false, effective, structural, ref totalPixels,
                Array.Empty<PdfRenderCapabilityDiagnostic>(), Array.Empty<PdfRenderCapabilityDiagnostic>(), cancellationToken);
            totalOutputBytes = checked(totalOutputBytes + page.OutputByteLength);
            if (totalOutputBytes > effective.MaxTotalOutputBytes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.RenderBytes, effective.MaxTotalOutputBytes, totalOutputBytes);
            }
            pages.Add(page);
        }

        return new PdfVisualComparisonReport(pages.AsReadOnly(), structural.AsReadOnly(), expectedPagePngs.Count, actualPagePngs.Count,
            expectedPages, actualPages, unmatchedExpected, unmatchedActual,
            ComputeSequenceSha256(expectedPagePngs, cancellationToken), ComputeSequenceSha256(actualPagePngs, cancellationToken), selectedScope: false);
    }

    private static OfficeRasterImage DecodePagePng(byte[] png, int pageNumber, string parameterName, PdfVisualComparisonOptions options, ref long totalPixels) {
        // Charge the declared size before decoding so the comparison budgets bound the raster allocation.
        if (!TryReadPngSize(png, out int width, out int height)) {
            throw new ArgumentException("Page image " + pageNumber + " is not a readable PNG.", parameterName);
        }
        AddPixelBudget(width, height, 1D, options, ref totalPixels);
        if (!OfficePngReader.TryDecode(png, out OfficeRasterImage? image) || image is null
            || image.Width != width || image.Height != height) {
            throw new ArgumentException("Page image " + pageNumber + " is not a readable PNG.", parameterName);
        }
        return image;
    }

    private static bool TryReadPngSize(byte[] png, out int width, out int height) {
        width = 0;
        height = 0;
        // PNG signature, then the IHDR chunk: length, type, width, height.
        if (png.Length < 24 || png[0] != 137 || png[1] != 80 || png[2] != 78 || png[3] != 71
            || png[12] != (byte)'I' || png[13] != (byte)'H' || png[14] != (byte)'D' || png[15] != (byte)'R') {
            return false;
        }
        width = png[16] << 24 | png[17] << 16 | png[18] << 8 | png[19];
        height = png[20] << 24 | png[21] << 16 | png[22] << 8 | png[23];
        return width > 0 && height > 0;
    }

    private static string ComputeSequenceSha256(IReadOnlyList<byte[]> images, CancellationToken cancellationToken) {
        using var sha256 = System.Security.Cryptography.IncrementalHash.CreateHash(System.Security.Cryptography.HashAlgorithmName.SHA256);
        foreach (byte[] image in images) {
            cancellationToken.ThrowIfCancellationRequested();
            sha256.AppendData(BitConverter.GetBytes((long)image.Length));
            sha256.AppendData(image);
        }
        return PdfArtifactInfo.ToLowerHex(sha256.GetHashAndReset());
    }

    private static PdfVisualPageComparison CompareRasters(
        OfficeRasterImage expectedImage, OfficeRasterImage actualImage, byte[]? expectedPng, byte[]? actualPng,
        int expectedPageNumber, int actualPageNumber, bool incomplete, PdfVisualComparisonOptions options,
        List<string> structural, ref long totalPixels,
        IReadOnlyList<PdfRenderCapabilityDiagnostic> expectedDiagnostics, IReadOnlyList<PdfRenderCapabilityDiagnostic> actualDiagnostics,
        CancellationToken cancellationToken) {
        bool hasSizeDifference = expectedImage.Width != actualImage.Width || expectedImage.Height != actualImage.Height;
        if (hasSizeDifference) {
            structural.Add("Page " + expectedPageNumber + " dimensions: expected " + expectedImage.Width + "x" + expectedImage.Height + ", actual " + actualImage.Width + "x" + actualImage.Height + ".");
        }

        int width = Math.Max(expectedImage.Width, actualImage.Width);
        int height = Math.Max(expectedImage.Height, actualImage.Height);
        AddPixelBudget(width, height, 1D, options, ref totalPixels);
        PdfVisualComparisonOptions.EnsureIgnoredRegionWork(options.IgnoredRegions.Count, (long)width * height);
        (int ExpectedX, int ExpectedY) = GetOffset(width, height, expectedImage.Width, expectedImage.Height, options.Alignment);
        (int ActualX, int ActualY) = GetOffset(width, height, actualImage.Width, actualImage.Height, options.Alignment);
        var diff = new OfficeRasterImage(width, height, OfficeColor.White);
        var changedPixels = new System.Collections.BitArray(checked(width * height));
        long compared = 0;
        long different = 0;
        long channelDifferenceTotal = 0;
        int maximumDifference = 0;
        int left = width, top = height, right = -1, bottom = -1;
        for (int y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                if (options.IgnoredRegions.Any(region => region.Contains(x, y))) {
                    diff.SetPixel(x, y, OfficeColor.FromRgb(224, 224, 224));
                    continue;
                }

                OfficeColor expected = GetPixel(expectedImage, x - ExpectedX, y - ExpectedY, options.Background);
                OfficeColor actual = GetPixel(actualImage, x - ActualX, y - ActualY, options.Background);
                int pixelMax = 0;
                int pixelTotal = 0;
                AddDifference(expected.R, actual.R, ref pixelMax, ref pixelTotal);
                AddDifference(expected.G, actual.G, ref pixelMax, ref pixelTotal);
                AddDifference(expected.B, actual.B, ref pixelMax, ref pixelTotal);
                AddDifference(expected.A, actual.A, ref pixelMax, ref pixelTotal);
                compared++;
                channelDifferenceTotal += pixelTotal;
                maximumDifference = Math.Max(maximumDifference, pixelMax);
                if (pixelMax > options.ChannelTolerance && options.IgnoreAntialiasing && !hasSizeDifference
                    && PdfVisualAntialiasing.IsEdgeShift(expectedImage, actualImage, x, y, options.ChannelTolerance)) {
                    diff.SetPixel(x, y, OfficeColor.FromRgb(255, 224, 96));
                } else if (pixelMax > options.ChannelTolerance) {
                    different++;
                    changedPixels[checked(y * width + x)] = true;
                    left = Math.Min(left, x); top = Math.Min(top, y);
                    right = Math.Max(right, x); bottom = Math.Max(bottom, y);
                    diff.SetPixel(x, y, OfficeColor.FromRgb(255, (byte)Math.Max(0, 160 - pixelMax / 2), (byte)Math.Max(0, 160 - pixelMax / 2)));
                } else {
                    byte gray = (byte)Math.Round((expected.R + expected.G + expected.B) / 3D);
                    diff.SetPixel(x, y, OfficeColor.FromRgb(gray, gray, gray));
                }
            }
        }

        double ratio = compared == 0 ? 0D : different / (double)compared;
        double mean = compared == 0 ? 0D : channelDifferenceTotal / (double)(compared * 4L);
        expectedPng ??= OfficePngWriter.Encode(expectedImage, cancellationToken);
        actualPng ??= OfficePngWriter.Encode(actualImage, cancellationToken);
        byte[] diffPng = OfficePngWriter.Encode(diff, cancellationToken);
        return new PdfVisualPageComparison(
            expectedPageNumber,
            actualPageNumber,
            !incomplete && !hasSizeDifference && ratio <= options.AllowedDifferenceRatio,
            width,
            height,
            compared,
            different,
            maximumDifference,
            mean,
            expectedPng,
            actualPng,
            diffPng,
            hasSizeDifference,
            different == 0 ? null : new PdfPixelRegion(left, top, right - left + 1, bottom - top + 1),
            changedPixels, expectedDiagnostics, actualDiagnostics);
    }

    private static void AddPixelBudget(double width, double height, double scale, PdfVisualComparisonOptions options, ref long totalPixels) {
        int pixelWidth = checked((int)Math.Ceiling(width * scale));
        int pixelHeight = checked((int)Math.Ceiling(height * scale));
        long pixels = checked((long)pixelWidth * pixelHeight);
        if (pixels > options.MaxPixelsPerImage) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPixels, options.MaxPixelsPerImage, pixels);
        }
        totalPixels = checked(totalPixels + pixels);
        if (totalPixels > options.MaxTotalPixels) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPixels, options.MaxTotalPixels, totalPixels);
        }
    }

    private static (int X, int Y) GetOffset(int canvasWidth, int canvasHeight, int imageWidth, int imageHeight, PdfVisualPageAlignment alignment) =>
        alignment == PdfVisualPageAlignment.Center
            ? ((canvasWidth - imageWidth) / 2, (canvasHeight - imageHeight) / 2)
            : (0, 0);

    private static OfficeColor GetPixel(OfficeRasterImage image, int x, int y, OfficeColor outside) =>
        x >= 0 && y >= 0 && x < image.Width && y < image.Height ? image.GetPixel(x, y) : outside;

    private static void AddDifference(byte left, byte right, ref int maximum, ref int total) {
        int difference = Math.Abs(left - right);
        maximum = Math.Max(maximum, difference);
        total += difference;
    }
}
