using System.Globalization;
using System.Xml.Linq;
using HtmlTinkerX;
using OfficeIMO.Drawing;
using OfficeIMO.Tests;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.ConversionConsistency;

internal static class BundleVerifier {
    internal static async Task<GateReport> VerifyAsync(string output, string rasterizer, CancellationToken cancellationToken) {
        EvidenceBundle bundle = GateJson.Read<EvidenceBundle>(Path.Combine(output, "bundle.json"));
        if (bundle.SchemaVersion != 2 || bundle.Cases.Count == 0) throw new InvalidDataException("Unsupported or empty bundle.");
        string version = await ArtifactPaths.RunAsync(rasterizer, new[] { "-v" }, cancellationToken: cancellationToken);
        await using var renderer = new HtmlBrowserPdfRenderer(new HtmlBrowserPdfRendererOptions(
            maximumBrowserInstances: 1, maximumQueuedCaptures: 4, networkPolicy: HtmlBrowserNetworkPolicy.Offline,
            setupTimeout: TimeSpan.FromSeconds(45)));
        string browserVersion = "";
        var reports = new List<CaseReport>();
        foreach (CaseBundle item in bundle.Cases) {
            try {
            var errors = new List<string>();
            if ((!item.Contract.ComparePdfPixels || !item.Contract.RequireSearchablePdf) && item.Contract.Limitations.Count == 0)
                errors.Add("Reduced PDF coverage requires an explicit limitation in the source contract.");
            foreach (string diagnostic in item.Diagnostics.Except(item.Contract.AllowedDiagnostics))
                errors.Add("Unaccepted diagnostic: " + diagnostic);
            string pdf = CheckedArtifact(output, item.PdfPath, item.PdfSha256);
            using var parsed = PdfPigDocument.Open(pdf);
            if (item.Contract.CaptureBrowserReference != (item.BrowserReference != null))
                errors.Add("Source-browser reference presence differs from the source contract.");
            string? browserPdf = item.BrowserReference == null
                ? null
                : CheckedArtifact(output, item.BrowserReference.PdfPath, item.BrowserReference.PdfSha256);
            using var browserParsed = browserPdf == null ? null : PdfPigDocument.Open(browserPdf);
            if (item.BrowserReference != null) {
                if (browserParsed!.NumberOfPages != item.BrowserReference.PageCount ||
                    browserParsed.NumberOfPages != item.Contract.Pages.Count)
                    errors.Add("Source-browser reference page count differs from the bundle or page contract.");
                if (string.IsNullOrWhiteSpace(browserVersion)) browserVersion = item.BrowserReference.BrowserVersion;
                else if (!string.Equals(browserVersion, item.BrowserReference.BrowserVersion, StringComparison.Ordinal))
                    errors.Add("Source-browser references were captured with different browser versions.");
            }
            if (parsed.NumberOfPages != item.Contract.Pages.Count)
                errors.Add($"PDF page count: expected {item.Contract.Pages.Count}, got {parsed.NumberOfPages}.");
            foreach (OfficeImageExportFormat format in Enum.GetValues<OfficeImageExportFormat>()) {
                var sequence = item.Images.Where(image => image.Format == format.ToString()).Select(image => image.Page).Order().ToArray();
                if (!sequence.SequenceEqual(Enumerable.Range(1, item.Contract.Pages.Count)))
                    errors.Add(format + " page sequence is missing, duplicated, or unexpected.");
            }
            var pages = new List<PageReport>();
            for (int index = 0; index < item.Contract.Pages.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                int pageNumber = index + 1;
                PageExpectation expected = item.Contract.Pages[index];
                var pageErrors = new List<string>();
                var comparisons = new List<ComparisonReport>();
                if (pageNumber > parsed.NumberOfPages) {
                    pages.Add(new PageReport(pageNumber, false, new List<string> { "Missing PDF page." }, comparisons));
                    continue;
                }
                var pdfPage = parsed.GetPage(pageNumber);
                if (item.Contract.RequireSearchablePdf) {
                    CheckText("PDF", pdfPage.Text, expected, pageErrors);
                    CheckTextPositions("PDF", pdfPage, expected, bundle.Dpi, pageErrors);
                }
                if (Math.Abs((double)pdfPage.Width * bundle.Dpi / 72D - (expected.PdfWidth ?? expected.Width)) > 1 ||
                    Math.Abs((double)pdfPage.Height * bundle.Dpi / 72D - (expected.PdfHeight ?? expected.Height)) > 1)
                    pageErrors.Add("PDF physical dimensions differ from the page contract.");
                string prefix = item.Contract.Id + "/verified-" + pageNumber.ToString("D3");
                string pdfPng = ArtifactPaths.Resolve(output, prefix + "-pdf.png");
                await RasterizeAsync(rasterizer, pdf, pageNumber, bundle.Dpi, pdfPng, cancellationToken);
                OfficeRasterImage pdfRaster = Decode(File.ReadAllBytes(pdfPng), "Png");
                if (item.Contract.ComparePdfPixels) CheckRegions("PDF", pdfRaster, expected, pageErrors);
                ImageArtifact? directArtifact = item.Images.SingleOrDefault(image => image.Page == pageNumber && image.Format == "Png");
                OfficeRasterImage? direct = directArtifact == null
                    ? null
                    : Decode(File.ReadAllBytes(CheckedArtifact(output, directArtifact.Path, directArtifact.Sha256)), "Png");
                if (direct == null) pageErrors.Add("Missing direct PNG reference.");
                if (browserPdf != null && browserParsed != null && pageNumber <= browserParsed.NumberOfPages) {
                    var browserPage = browserParsed.GetPage(pageNumber);
                    PageExpectation browserExpected = expected with {
                        TextPositions = expected.BrowserTextPositions,
                        Regions = expected.BrowserRegions
                    };
                    CheckText("Browser PDF", browserPage.Text, browserExpected, pageErrors);
                    CheckTextPositions("Browser PDF", browserPage, browserExpected, bundle.Dpi, pageErrors);
                    if (Math.Abs((double)browserPage.Width * bundle.Dpi / 72D - expected.Width) > 1 ||
                        Math.Abs((double)browserPage.Height * bundle.Dpi / 72D - expected.Height) > 1)
                        pageErrors.Add("Source-browser PDF physical dimensions differ from the page contract.");
                    string browserPng = ArtifactPaths.Resolve(output, prefix + "-browser.png");
                    await RasterizeAsync(rasterizer, browserPdf, pageNumber, bundle.Dpi, browserPng, cancellationToken);
                    OfficeRasterImage browserRaster = Decode(File.ReadAllBytes(browserPng), "Png");
                    CheckRegions("Browser PDF", browserRaster, browserExpected, pageErrors);
                    if (direct != null) {
                        PixelTolerance tolerance = item.Contract.BrowserVisualTolerance;
                        VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(direct,
                            browserRaster, tolerance.Channel,
                            (int)(expected.Width * (double)expected.Height * tolerance.DifferentRatio), tolerance.MeanAbsoluteError);
                        string diff = prefix + "-browser-vs-png-diff.png";
                        File.WriteAllBytes(ArtifactPaths.Resolve(output, diff), comparison.DiffPng);
                        comparisons.Add(new ComparisonReport("browser-vs-png", comparison.Passed, comparison.DifferentPixels,
                            comparison.TotalPixels, double.IsFinite(comparison.MeanAbsoluteError) ? comparison.MeanAbsoluteError : null, diff));
                        if (!comparison.Passed) pageErrors.Add("browser-vs-png exceeds the visual tolerance.");
                    }
                }
                foreach (ImageArtifact artifact in item.Images.Where(image => image.Page == pageNumber).OrderBy(image => image.Format == "Png" ? 0 : 1)) {
                    string path = CheckedArtifact(output, artifact.Path, artifact.Sha256);
                    OfficeRasterImage raster;
                    if (artifact.Format == "Svg") {
                        XDocument svg = XDocument.Load(path);
                        CheckText("SVG", string.Concat(svg.Descendants().Where(node => node.Name.LocalName == "text").Select(node => node.Value)), expected, pageErrors);
                        double cssWidth = expected.Width * 96D / bundle.Dpi;
                        double cssHeight = expected.Height * 96D / bundle.Dpi;
                        string width = cssWidth.ToString(CultureInfo.InvariantCulture) + "px";
                        string height = cssHeight.ToString(CultureInfo.InvariantCulture) + "px";
                        string html = "<!doctype html><style>html,body{margin:0;padding:0}svg{display:block;width:" + width + ";height:" + height + "}</style>" + svg.Root;
                        var capture = await renderer.CaptureAsync(new HtmlBrowserPdfRequest(HtmlBrowserPdfSource.FromHtml(html),
                            new HtmlBrowserPdfOptions(width: width, height: height, marginTop: "0", marginRight: "0", marginBottom: "0", marginLeft: "0", printBackground: true),
                            readiness: new HtmlBrowserPdfReadiness(loadState: HtmlBrowserLoadState.Load, stable: true, stableMilliseconds: 250, timeout: 15000)), cancellationToken);
                        browserVersion = capture.Diagnostics.BrowserVersion;
                        if (capture.Diagnostics.BlockedRequestCount > 0 || capture.Diagnostics.Warnings.Count > 0)
                            pageErrors.Add("Independent SVG renderer reported blocked resources or warnings.");
                        string svgPdf = ArtifactPaths.Resolve(output, prefix + "-svg.pdf");
                        File.WriteAllBytes(svgPdf, capture.PdfBytes);
                        using var svgParsed = PdfPigDocument.Open(capture.PdfBytes);
                        if (svgParsed.NumberOfPages != 1) pageErrors.Add("Independent SVG rendering produced extra pages.");
                        CheckTextPositions("SVG", svgParsed.GetPage(1), expected, bundle.Dpi, pageErrors);
                        CheckTextPositions("SVG", svgParsed.GetPage(1), expected with { TextPositions = expected.SvgTextPositions }, bundle.Dpi, pageErrors);
                        string svgPng = ArtifactPaths.Resolve(output, prefix + "-svg.png");
                        await RasterizeAsync(rasterizer, svgPdf, 1, bundle.Dpi, svgPng, cancellationToken);
                        raster = Decode(File.ReadAllBytes(svgPng), "Png");
                    } else raster = Decode(File.ReadAllBytes(path), artifact.Format);
                    CheckRegions(artifact.Format, raster, expected, pageErrors);
                    if (artifact.Format == "Png" && !item.Contract.ComparePdfPixels) continue;
                    if (direct == null) continue;
                    string route = artifact.Format == "Png" ? "png-vs-pdf" : artifact.Format.ToLowerInvariant() + "-vs-png";
                    var tolerance = item.Contract.VisualTolerance;
                    VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(direct,
                        artifact.Format == "Png" ? pdfRaster : raster, tolerance.Channel,
                        (int)(expected.Width * (double)expected.Height * tolerance.DifferentRatio), tolerance.MeanAbsoluteError);
                    string diff = prefix + "-" + route + "-diff.png";
                    File.WriteAllBytes(ArtifactPaths.Resolve(output, diff), comparison.DiffPng);
                    comparisons.Add(new ComparisonReport(route, comparison.Passed, comparison.DifferentPixels, comparison.TotalPixels, double.IsFinite(comparison.MeanAbsoluteError) ? comparison.MeanAbsoluteError : null, diff));
                    if (!comparison.Passed) pageErrors.Add(route + " exceeds the visual tolerance.");
                }
                pages.Add(new PageReport(pageNumber, pageErrors.Count == 0, pageErrors, comparisons));
            }
            reports.Add(new CaseReport(item.Contract.Id, errors.Count == 0 && pages.All(page => page.Passed), errors, pages,
                item.PdfRoute, item.Contract.ComparePdfPixels,
                item.Contract.RequireSearchablePdf && errors.Count == 0 && pages.All(page => page.Passed), item.Contract.Limitations));
            } catch (Exception error) when (error is not OperationCanceledException) {
                reports.Add(new CaseReport(item.Contract.Id, false, new() { error.Message }, new(),
                    item.PdfRoute, item.Contract.ComparePdfPixels, false, item.Contract.Limitations));
            }
        }
        return new GateReport(2, bundle.Commit, bundle.FontSha256, version, browserVersion, reports.All(item => item.Passed), reports);
    }

    private static string CheckedArtifact(string output, string relative, string hash) {
        string path = ArtifactPaths.Resolve(output, relative);
        if (ArtifactPaths.HashFile(path) != hash) throw new InvalidDataException("Artifact hash mismatch: " + relative);
        return path;
    }

    private static Task<string> RasterizeAsync(string executable, string pdf, int page, int dpi, string png, CancellationToken token) =>
        ArtifactPaths.RunAsync(executable, new[] { "-f", page.ToString(CultureInfo.InvariantCulture), "-l", page.ToString(CultureInfo.InvariantCulture),
            "-singlefile", "-r", dpi.ToString(CultureInfo.InvariantCulture), "-png", pdf, Path.ChangeExtension(png, null) }, cancellationToken: token);

    private static OfficeRasterImage Decode(byte[] bytes, string format) {
        OfficeRasterImage? image;
        bool success = format switch {
            "Png" => OfficePngReader.TryDecode(bytes, out image),
            "Jpeg" => OfficeJpegCodec.TryDecode(bytes, out image),
            "Tiff" => OfficeTiffCodec.TryDecode(bytes, out image),
            "Webp" => OfficeWebpCodec.TryDecode(bytes, out image),
            _ => throw new InvalidDataException("Unsupported raster format: " + format)
        };
        return success && image != null ? image : throw new InvalidDataException("Cannot decode " + format);
    }

    private static void CheckText(string route, string actual, PageExpectation expected, List<string> errors) {
        foreach (string label in expected.Text)
            if (!actual.Contains(label, StringComparison.Ordinal)) errors.Add(route + " is missing labelled text: " + label);
    }

    private static void CheckTextPositions(string route, UglyToad.PdfPig.Content.Page page, PageExpectation expected, int dpi, List<string> errors) {
        string text = string.Concat(page.Letters.Select(letter => letter.Value));
        foreach (TextPositionExpectation position in expected.TextPositions) {
            int offset = text.IndexOf(position.Text, StringComparison.Ordinal);
            if (offset < 0) { errors.Add(route + " is missing positioned text: " + position.Text); continue; }
            int cursor = 0;
            var letter = page.Letters.First(item => { bool contains = cursor <= offset && cursor + item.Value.Length > offset; cursor += item.Value.Length; return contains; });
            double x = letter.StartBaseLine.X * dpi / 72D;
            double baseline = ((double)page.Height - letter.StartBaseLine.Y) * dpi / 72D;
            if (Math.Abs(x - position.X) > position.Tolerance || Math.Abs(baseline - position.Baseline) > position.Tolerance)
                errors.Add($"{route} text {position.Text}: baseline ({x:F3},{baseline:F3}), expected ({position.X},{position.Baseline}) within {position.Tolerance}px.");
        }
    }

    private static void CheckRegions(string route, OfficeRasterImage image, PageExpectation expected, List<string> errors) {
        if (image.Width != expected.Width || image.Height != expected.Height)
            errors.Add($"{route} dimensions: expected {expected.Width}x{expected.Height}, got {image.Width}x{image.Height}.");
        foreach (RegionExpectation region in expected.Regions) {
            string hex = region.Color.TrimStart('#');
            int red = int.Parse(hex[..2], NumberStyles.HexNumber), green = int.Parse(hex[2..4], NumberStyles.HexNumber), blue = int.Parse(hex[4..6], NumberStyles.HexNumber);
            if (region.X < 0 || region.Y < 0 || region.Width <= 0 || region.Height <= 0 || region.X + region.Width > image.Width || region.Y + region.Height > image.Height)
                { errors.Add(route + " region is outside the page: " + region.Id); continue; }
            int count = 0;
            for (int y = region.Y; y < region.Y + region.Height; y++)
                for (int x = region.X; x < region.X + region.Width; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (Math.Abs(pixel.R - red) <= region.ColorTolerance && Math.Abs(pixel.G - green) <= region.ColorTolerance && Math.Abs(pixel.B - blue) <= region.ColorTolerance) count++;
                }
            if (count < region.MinimumPixels || count > region.MaximumPixels)
                errors.Add($"{route} region {region.Id}: {count} matching pixels, expected {region.MinimumPixels}..{region.MaximumPixels}.");
        }
    }
}
