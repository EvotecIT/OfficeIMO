using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using HtmlTinkerX;
using OfficeIMO.Html.Pdf.Browser;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Html.Pdf.Browser.Tests;

public sealed class HtmlBrowserPdfOfficeExtensionsTests {
    [Fact]
    public async Task CancelledCaptureConversionStopsBeforeReadingInvalidPdfBytes() {
        HtmlBrowserPdfResult capture = CreateCapture(new byte[] { 1, 2, 3 }, tagged: false);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.ThrowsAny<OperationCanceledException>(() =>
            capture.ToPdfDocumentResult(cancellationToken: cancellation.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            capture.ToPdfDocumentResultAsync(cancellationToken: cancellation.Token));
    }

    [Fact]
    public void CapturedPdfOpensIntoTheCanonicalDocumentModelWithBrowserDiagnostics() {
        byte[] bytes = PdfDocument.Create(pdf => pdf.Content(content => content
            .Paragraph(paragraph => paragraph.Text("Browser capture bridge"))))
            .ToBytes();
        HtmlBrowserPdfResult capture = CreateCapture(bytes, tagged: false);

        PdfDocumentConversionResult result = capture.ToPdfDocumentResult();

        Assert.True(result.Value.Preflight().CanExtractText);
        Assert.True(result.Value.PlanMutation(PdfMutationOperation.ModifyPageContent).CanExecute);
        HtmlBrowserPdfCaptureReport report = Assert.IsType<HtmlBrowserPdfCaptureReport>(
            Assert.Single(result.SourceConversionReports));
        Assert.False(report.Tagged);
        Assert.False(report.HasLoss);
        Assert.Same(capture.Diagnostics, report.Diagnostics);
    }

    [Fact]
    public void CaptureReportTreatsBlockedResourcesAndWarningsAsPotentialLoss() {
        HtmlBrowserPdfResult capture = CreateCapture(
            PdfDocument.Create(pdf => pdf.Content(content => content
                .Paragraph(paragraph => paragraph.Text("Incomplete capture")))).ToBytes(),
            tagged: true,
            blockedRequestCount: 1,
            blockedRequests: new[] { "https://blocked.example/font.woff2" },
            warnings: new[] { "A browser resource was unavailable." });

        PdfDocumentConversionResult result = capture.ToPdfDocumentResult();
        HtmlBrowserPdfCaptureReport report = Assert.IsType<HtmlBrowserPdfCaptureReport>(
            Assert.Single(result.SourceConversionReports));

        Assert.True(report.Tagged);
        Assert.True(report.HasLoss);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.True(result.Value.Preflight().CanExtractText);
    }

    [Fact]
    public async Task CapturedPdfCanBeOpenedAsynchronouslyWithoutLosingTaggedIntent() {
        HtmlBrowserPdfResult capture = CreateCapture(
            PdfDocument.Create(pdf => pdf.Content(content => content
                .Paragraph(paragraph => paragraph.Text("Async bridge")))).ToBytes(),
            tagged: true);

        PdfDocumentConversionResult result = await capture.ToPdfDocumentResultAsync();

        HtmlBrowserPdfCaptureReport report = Assert.IsType<HtmlBrowserPdfCaptureReport>(
            Assert.Single(result.SourceConversionReports));
        Assert.True(report.Tagged);
        Assert.True(result.Value.Preflight().CanRead);
    }

    private static HtmlBrowserPdfResult CreateCapture(
        byte[] bytes,
        bool tagged,
        int blockedRequestCount = 0,
        IReadOnlyList<string>? blockedRequests = null,
        IReadOnlyList<string>? warnings = null) {
        HtmlBrowserPdfDiagnostics diagnostics = (HtmlBrowserPdfDiagnostics)CreateNonPublic(
            typeof(HtmlBrowserPdfDiagnostics),
            HtmlBrowserPdfSourceKind.Html,
            1L,
            false,
            false,
            "about:blank",
            "test",
            TimeSpan.Zero,
            TimeSpan.Zero,
            TimeSpan.Zero,
            TimeSpan.Zero,
            TimeSpan.Zero,
            blockedRequestCount,
            blockedRequests ?? Array.Empty<string>(),
            warnings ?? Array.Empty<string>());

        return (HtmlBrowserPdfResult)CreateNonPublic(
            typeof(HtmlBrowserPdfResult),
            bytes,
            diagnostics,
            tagged);
    }

    private static object CreateNonPublic(Type type, params object[] arguments) =>
        Activator.CreateInstance(
            type,
            BindingFlags.Instance | BindingFlags.NonPublic,
            binder: null,
            args: arguments,
            culture: null) ?? throw new InvalidOperationException("Unable to construct " + type.FullName + ".");
}
