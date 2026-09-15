using HtmlTinkerX;
using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using OfficeIMO.Tests;

namespace OfficeIMO.ConversionConsistency;

internal static class SvgContentSafetyBrowserEvidenceRunner {
    internal static async Task<SvgContentSafetyBrowserEvidence> RunAsync(
        string repository,
        string output,
        string rasterizer,
        string? browserExecutablePath,
        CancellationToken cancellationToken) {
        if (Directory.Exists(output) && Directory.EnumerateFileSystemEntries(output).Any()) {
            throw new IOException("SVG content-safety evidence output must be empty.");
        }
        Directory.CreateDirectory(output);
        string fixturePath = Path.Combine(
            repository,
            "OfficeIMO.TestAssets",
            "ContentSafety",
            "svg-concealed-adversarial.svg");
        byte[] source = File.ReadAllBytes(fixturePath);
        var readerOptions = new OfficeSvgDrawingReaderOptions();
        readerOptions.Fonts.AddRange(PortableVisualFontAssets.CreateSpreadsheetFonts());
        OfficeContentSafetyReport inspection = OfficeSvgDrawingReader.InspectContentSafety(source, readerOptions: readerOptions);
        ValidateExpectedFindingInventory(inspection);
        string[] selected = inspection.Findings
            .Where(finding => finding.CleanupCapability != OfficeContentCleanupCapability.ReportOnly)
            .Select(finding => finding.Id)
            .ToArray();
        if (selected.Length == 0) throw new InvalidDataException("The SVG adversarial fixture produced no cleanup-capable findings.");
        OfficeContentCleanupResult cleaned = OfficeSvgDrawingReader.RemoveSelectedContent(
            source,
            new OfficeContentCleanupSelection(selected),
            readerOptions: readerOptions);
        if (!cleaned.Changed || cleaned.After.Findings.Any(finding =>
                finding.CleanupCapability != OfficeContentCleanupCapability.ReportOnly)) {
            throw new InvalidDataException("The SVG adversarial fixture did not reach a stable cleanup state.");
        }
        ValidateExpectedReportOnlyInventory(cleaned.After);

        var rendererOptions = new HtmlBrowserPdfRendererOptions(
            browserExecutablePath: browserExecutablePath,
            networkPolicy: HtmlBrowserNetworkPolicy.Offline,
            setupTimeout: TimeSpan.FromSeconds(45));
        await using var renderer = new HtmlBrowserPdfRenderer(rendererOptions);
        HtmlBrowserPdfResult before = await CaptureAsync(renderer, source, cancellationToken);
        HtmlBrowserPdfResult after = await CaptureAsync(renderer, cleaned.Output, cancellationToken);
        ValidateBrowserDiagnostics(before, "before");
        ValidateBrowserDiagnostics(after, "after");

        string beforePdf = Path.Combine(output, "svg-content-safety-before.pdf");
        string afterPdf = Path.Combine(output, "svg-content-safety-after.pdf");
        File.WriteAllBytes(beforePdf, before.PdfBytes);
        File.WriteAllBytes(afterPdf, after.PdfBytes);
        string beforePng = Path.Combine(output, "svg-content-safety-before.png");
        string afterPng = Path.Combine(output, "svg-content-safety-after.png");
        await RasterizeAsync(rasterizer, beforePdf, beforePng, cancellationToken);
        await RasterizeAsync(rasterizer, afterPdf, afterPng, cancellationToken);

        OfficeRasterImage beforeRaster = VisualBaselineTestSupport.DecodePng(
            File.ReadAllBytes(beforePng),
            "The browser-rendered SVG before image is invalid.");
        OfficeRasterImage afterRaster = VisualBaselineTestSupport.DecodePng(
            File.ReadAllBytes(afterPng),
            "The browser-rendered SVG after image is invalid.");
        int beforeVisibleControlPixels = ValidateExpectedBrowserRaster(beforeRaster, "before");
        int afterVisibleControlPixels = ValidateExpectedBrowserRaster(afterRaster, "after");
        VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(
                beforeRaster,
                afterRaster,
                channelTolerance: 0,
                allowedDifferentPixels: 0,
                maximumMeanAbsoluteError: 0D);
        string diffPath = Path.Combine(output, "svg-content-safety-diff.png");
        File.WriteAllBytes(diffPath, comparison.DiffPng);

        string browserVersion = string.Equals(before.Diagnostics.BrowserVersion, after.Diagnostics.BrowserVersion, StringComparison.Ordinal)
            ? before.Diagnostics.BrowserVersion
            : before.Diagnostics.BrowserVersion + " / " + after.Diagnostics.BrowserVersion;
        var report = new SvgContentSafetyBrowserEvidence(
            SchemaVersion: 1,
            Passed: comparison.Passed,
            BrowserVersion: browserVersion,
            BeforeBlockedRequests: before.Diagnostics.BlockedRequestCount,
            AfterBlockedRequests: after.Diagnostics.BlockedRequestCount,
            RasterWidth: beforeRaster.Width,
            RasterHeight: beforeRaster.Height,
            BeforeVisibleControlPixels: beforeVisibleControlPixels,
            AfterVisibleControlPixels: afterVisibleControlPixels,
            FixtureSha256: ArtifactPaths.Hash(source),
            CleanedSha256: ArtifactPaths.Hash(cleaned.Output),
            RemovedFindings: cleaned.Changes.Count,
            RemainingReportOnlyFindings: cleaned.After.Findings.Count,
            DifferentPixels: comparison.DifferentPixels,
            TotalPixels: comparison.TotalPixels,
            MeanAbsoluteError: comparison.MeanAbsoluteError,
            BeforePngSha256: ArtifactPaths.HashFile(beforePng),
            AfterPngSha256: ArtifactPaths.HashFile(afterPng),
            DiffPngSha256: ArtifactPaths.HashFile(diffPath));
        GateJson.Write(Path.Combine(output, "svg-content-safety-browser-evidence.json"), report);
        return report;
    }

    private static void ValidateExpectedFindingInventory(OfficeContentSafetyReport report) {
        OfficeContentSafetyFinding[] removable = report.Findings
            .Where(finding => finding.CleanupCapability != OfficeContentCleanupCapability.ReportOnly)
            .ToArray();
        OfficeContentSafetyFinding[] reportOnly = report.Findings
            .Where(finding => finding.CleanupCapability == OfficeContentCleanupCapability.ReportOnly)
            .ToArray();
        if (removable.Length != 6 || reportOnly.Length != 3) {
            throw new InvalidDataException(
                "The SVG adversarial fixture produced an unexpected cleanup/report-only finding inventory (" +
                removable.Length + "/" + reportOnly.Length + ").");
        }
        RequireFinding(removable, "CSS DISPLAY PAYLOAD", OfficeContentConcealmentKind.HiddenByProperty);
        RequireFinding(removable, "DEEP SELECTOR PAYLOAD", OfficeContentConcealmentKind.HiddenByProperty);
        RequireFinding(removable, "PRESENTATION VISIBILITY PAYLOAD", OfficeContentConcealmentKind.HiddenByProperty);
        RequireFinding(removable, "OPACITY PAYLOAD", OfficeContentConcealmentKind.TransparentText);
        RequireFinding(removable, "CLIPPED PAYLOAD", OfficeContentConcealmentKind.ClippedContent);
        RequireFinding(removable, "OFF CANVAS PAYLOAD", OfficeContentConcealmentKind.OffCanvas);
        ValidateExpectedReportOnlyInventory(report);
    }

    private static void ValidateExpectedReportOnlyInventory(OfficeContentSafetyReport report) {
        OfficeContentSafetyFinding[] reportOnly = report.Findings
            .Where(finding => finding.CleanupCapability == OfficeContentCleanupCapability.ReportOnly)
            .ToArray();
        if (reportOnly.Length != 3) {
            throw new InvalidDataException(
                "The SVG adversarial fixture produced " + reportOnly.Length + " report-only findings instead of 3.");
        }
        RequireFinding(reportOnly, "LOW CONTRAST PAYLOAD", OfficeContentConcealmentKind.LowContrastText);
        if (!reportOnly.Any(finding => finding.TextPreview == "PAINT ORDER PAYLOAD" &&
                finding.Kind is OfficeContentConcealmentKind.LowContrastText or OfficeContentConcealmentKind.Other)) {
            throw new InvalidDataException("The SVG adversarial fixture did not report the expected paint-order finding.");
        }
        if (!reportOnly.Any(finding => finding.Kind == OfficeContentConcealmentKind.NonPrimaryContent &&
                finding.TextPreview.Contains(".css-hidden", StringComparison.Ordinal))) {
            throw new InvalidDataException("The SVG adversarial fixture did not report its stylesheet as non-primary content.");
        }
    }

    private static void RequireFinding(
        IEnumerable<OfficeContentSafetyFinding> findings,
        string text,
        OfficeContentConcealmentKind kind) {
        if (!findings.Any(finding => finding.TextPreview == text && finding.Kind == kind)) {
            throw new InvalidDataException(
                "The SVG adversarial fixture did not produce the expected " + kind + " finding for " + text + ".");
        }
    }

    private static async Task<HtmlBrowserPdfResult> CaptureAsync(
        HtmlBrowserPdfRenderer renderer,
        byte[] svg,
        CancellationToken cancellationToken) {
        string svgDataUrl = "data:image/svg+xml;base64," + Convert.ToBase64String(svg);
        string html = "<!doctype html><html><head><meta charset=\"utf-8\"><style>" +
            "@page{size:640px 430px;margin:0}html,body{width:640px;height:430px;margin:0;padding:0;overflow:hidden}" +
            "img{display:block;width:640px;height:430px}</style></head><body>" +
            "<img alt=\"\" src=\"" + svgDataUrl + "\"></body></html>";
        return await renderer.CaptureAsync(
            new HtmlBrowserPdfRequest(
                HtmlBrowserPdfSource.FromHtml(html),
                new HtmlBrowserPdfOptions(
                    width: "640px",
                    height: "430px",
                    marginTop: "0",
                    marginRight: "0",
                    marginBottom: "0",
                    marginLeft: "0",
                    printBackground: true),
                readiness: new HtmlBrowserPdfReadiness(
                    loadState: HtmlBrowserLoadState.Load,
                    stable: true,
                    stableMilliseconds: 250,
                    timeout: 15000)),
            cancellationToken);
    }

    private static void ValidateBrowserDiagnostics(HtmlBrowserPdfResult capture, string lane) {
        if (capture.Diagnostics.Warnings.Count != 0) {
            throw new InvalidDataException(
                "The " + lane + " SVG browser capture reported warnings: " +
                string.Join(" | ", capture.Diagnostics.Warnings));
        }
    }

    private static int ValidateExpectedBrowserRaster(OfficeRasterImage raster, string lane) {
        if (raster.Width != 640 || raster.Height is < 430 or > 431) {
            throw new InvalidDataException(
                "The " + lane + " SVG browser raster has unexpected dimensions " +
                raster.Width + "x" + raster.Height + ".");
        }
        OfficeColor background = raster.GetPixel(raster.Width - 1, raster.Height - 1);
        if (background.A != 255 || background.R < 248 || background.G < 248 || background.B < 248) {
            throw new InvalidDataException("The " + lane + " SVG browser raster does not contain the expected white canvas.");
        }
        int visibleControlPixels = 0;
        for (int y = 15; y < 50; y++) {
            for (int x = 20; x < 250; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.A == 255 && pixel.R < 192 && pixel.G < 192 && pixel.B < 192) visibleControlPixels++;
            }
        }
        if (visibleControlPixels < 100) {
            throw new InvalidDataException("The " + lane + " SVG browser raster did not paint the expected visible control text.");
        }
        return visibleControlPixels;
    }

    private static Task<string> RasterizeAsync(
        string executable,
        string pdf,
        string png,
        CancellationToken cancellationToken) =>
        ArtifactPaths.RunAsync(
            executable,
            new[] { "-f", "1", "-l", "1", "-singlefile", "-r", "96", "-png", pdf, Path.ChangeExtension(png, null) },
            cancellationToken: cancellationToken);
}

internal sealed record SvgContentSafetyBrowserEvidence(
    int SchemaVersion,
    bool Passed,
    string BrowserVersion,
    int BeforeBlockedRequests,
    int AfterBlockedRequests,
    int RasterWidth,
    int RasterHeight,
    int BeforeVisibleControlPixels,
    int AfterVisibleControlPixels,
    string FixtureSha256,
    string CleanedSha256,
    int RemovedFindings,
    int RemainingReportOnlyFindings,
    int DifferentPixels,
    int TotalPixels,
    double MeanAbsoluteError,
    string BeforePngSha256,
    string AfterPngSha256,
    string DiffPngSha256);
