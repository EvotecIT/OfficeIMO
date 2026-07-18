using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Dependency-free rendered PDF comparison with structural evidence and review artifacts.</summary>
public static class PdfVisualComparer {
    /// <summary>Compares all common pages or a selected page set.</summary>
    public static PdfVisualComparisonReport Compare(
        byte[] expectedPdf,
        byte[] actualPdf,
        PdfPageSelection? selection = null,
        PdfVisualComparisonOptions? options = null,
        PdfReadOptions? expectedReadOptions = null,
        PdfReadOptions? actualReadOptions = null) {
        Guard.NotNull(expectedPdf, nameof(expectedPdf));
        Guard.NotNull(actualPdf, nameof(actualPdf));
        PdfVisualComparisonOptions effectiveOptions = options ?? new PdfVisualComparisonOptions();
        effectiveOptions.Validate();
        PdfReadDocument expected = PdfReadDocument.Open(expectedPdf, expectedReadOptions);
        PdfReadDocument actual = PdfReadDocument.Open(actualPdf, actualReadOptions);
        var structural = new List<string>();
        if (expected.Pages.Count != actual.Pages.Count) {
            structural.Add("PageCount: expected " + expected.Pages.Count + ", actual " + actual.Pages.Count + ".");
        }

        int commonPageCount = Math.Min(expected.Pages.Count, actual.Pages.Count);
        int[] pageNumbers = selection?.ToPageNumbers(expected.Pages.Count, nameof(selection)) ?? Enumerable.Range(1, commonPageCount).ToArray();
        var pages = new List<PdfVisualPageComparison>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            if (pageNumber > actual.Pages.Count) {
                structural.Add("Page " + pageNumber + " is missing from the actual document.");
                continue;
            }

            pages.Add(ComparePage(expected, actual, pageNumber, effectiveOptions, structural));
        }

        return new PdfVisualComparisonReport(pages.AsReadOnly(), structural.AsReadOnly());
    }

    private static PdfVisualPageComparison ComparePage(PdfReadDocument expectedDocument, PdfReadDocument actualDocument, int pageNumber, PdfVisualComparisonOptions options, List<string> structural) {
        OfficeDrawing expectedDrawing = PdfPageImageRenderer.RenderPage(expectedDocument, pageNumber);
        OfficeDrawing actualDrawing = PdfPageImageRenderer.RenderPage(actualDocument, pageNumber);
        byte[] expectedPng = OfficeDrawingRasterRenderer.ToPng(expectedDrawing, options.Scale, options.Background);
        byte[] actualPng = OfficeDrawingRasterRenderer.ToPng(actualDrawing, options.Scale, options.Background);
        if (!OfficeRasterImageDecoder.TryDecode(expectedPng, out OfficeRasterImage? expectedImage) || expectedImage is null ||
            !OfficeRasterImageDecoder.TryDecode(actualPng, out OfficeRasterImage? actualImage) || actualImage is null) {
            throw new InvalidOperationException("Managed PDF comparison could not decode its rendered PNG output.");
        }

        if (expectedImage.Width != actualImage.Width || expectedImage.Height != actualImage.Height) {
            structural.Add("Page " + pageNumber + " dimensions: expected " + expectedImage.Width + "x" + expectedImage.Height + ", actual " + actualImage.Width + "x" + actualImage.Height + ".");
        }

        int width = Math.Max(expectedImage.Width, actualImage.Width);
        int height = Math.Max(expectedImage.Height, actualImage.Height);
        (int ExpectedX, int ExpectedY) = GetOffset(width, height, expectedImage.Width, expectedImage.Height, options.Alignment);
        (int ActualX, int ActualY) = GetOffset(width, height, actualImage.Width, actualImage.Height, options.Alignment);
        var diff = new OfficeRasterImage(width, height, OfficeColor.White);
        long compared = 0;
        long different = 0;
        long channelDifferenceTotal = 0;
        int maximumDifference = 0;
        for (int y = 0; y < height; y++) {
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
                if (pixelMax > options.ChannelTolerance) {
                    different++;
                    diff.SetPixel(x, y, OfficeColor.FromRgb(255, (byte)Math.Max(0, 160 - pixelMax / 2), (byte)Math.Max(0, 160 - pixelMax / 2)));
                } else {
                    byte gray = (byte)Math.Round((expected.R + expected.G + expected.B) / 3D);
                    diff.SetPixel(x, y, OfficeColor.FromRgb(gray, gray, gray));
                }
            }
        }

        double ratio = compared == 0 ? 0D : different / (double)compared;
        double mean = compared == 0 ? 0D : channelDifferenceTotal / (double)(compared * 4L);
        return new PdfVisualPageComparison(
            pageNumber,
            ratio <= options.AllowedDifferenceRatio,
            width,
            height,
            compared,
            different,
            maximumDifference,
            mean,
            expectedPng,
            actualPng,
            OfficePngWriter.Encode(diff));
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
