using System.Text;
using OfficeIMO.Pdf;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Workflows.Tests {
    public sealed class PdfComparisonGalleryBudgetTests(ITestOutputHelper output) {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task GalleryOutputChargesItsActualEncodedSizeIncludingUnmatchedPages(bool longerExpected) {
            using var scope = new TestDirectory();
            var fixture = CreateComparison(longerExpected);
            string expected = Path.Combine(scope.Path, "expected.pdf");
            string actual = Path.Combine(scope.Path, "actual.pdf");
            string gallery = Path.Combine(scope.Path, "comparison.html");
            File.WriteAllBytes(expected, fixture.Expected);
            File.WriteAllBytes(actual, fixture.Actual);
            string canonicalGallery = fixture.Report.ToHtmlGallery("OfficeIMO document comparison");
            long galleryBytes = Encoding.UTF8.GetByteCount(canonicalGallery);
            PdfVisualPageComparison pair = Assert.Single(fixture.Report.Pages);
            long retainedPngBytes = RetainedPngBytes(pair);
            long embeddedPngBytes = Convert.ToBase64String(pair.ExpectedPng).Length +
                Convert.ToBase64String(pair.ActualPng).Length + Convert.ToBase64String(pair.DiffPng).Length;
            output.WriteLine($"Canonical gallery: {galleryBytes} UTF-8 bytes; retained PNGs: {retainedPngBytes} bytes; embedded base64: {embeddedPngBytes} bytes; markup: {galleryBytes - embeddedPngBytes} bytes; unmatched selected pages: 99.");

            var request = new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Compare,
                InputPath = expected,
                ComparisonPath = actual,
                ComparisonExpectedPages = fixture.ExpectedPages,
                ComparisonActualPages = fixture.ActualPages,
                OutputPath = gallery,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                Limits = new OfficeWorkflowLimits { MaximumOutputBytes = long.MaxValue }
            };
            OfficeWorkflowResult generous = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.True(generous.Succeeded, generous.Summary);
            Assert.Equal(canonicalGallery, File.ReadAllText(gallery));
            Assert.Equal(galleryBytes, generous.OutputBytes);
            Assert.Equal("1", generous.HealthReport!.Metrics["pagesCompared"]);
            Assert.Equal("100", generous.HealthReport.Metrics[longerExpected ? "expectedSelectedPageCount" : "actualSelectedPageCount"]);
            Assert.False(generous.HealthReport.Verified);

            request.Limits.MaximumOutputBytes = 20_000;
            OfficeWorkflowResult fitting = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.True(fitting.Succeeded, fitting.Summary);
            Assert.Equal(canonicalGallery, File.ReadAllText(gallery));
            Assert.Equal(galleryBytes, fitting.OutputBytes);

            File.WriteAllText(gallery, "existing report");
            request.Limits.MaximumOutputBytes = galleryBytes - 1;
            OfficeWorkflowResult bounded = await new OfficeWorkflowRunner().RunAsync(request);

            Assert.Equal(OfficeWorkflowStatus.Failed, bounded.Status);
            Assert.Single(bounded.Diagnostics, diagnostic => diagnostic.Code == "WorkflowFailed");
            Assert.Equal("existing report", File.ReadAllText(gallery));
            Assert.Null(bounded.OutputPath);
            Assert.Equal(0, bounded.OutputBytes);
            Assert.Equal(fixture.Expected, File.ReadAllBytes(expected));
            Assert.Equal(fixture.Actual, File.ReadAllBytes(actual));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task ReportOnlyComparisonChargesRetainedImagesWithoutGalleryEncoding(bool longerExpected) {
            using var scope = new TestDirectory();
            var fixture = CreateComparison(longerExpected);
            string expected = Path.Combine(scope.Path, "expected.pdf");
            string actual = Path.Combine(scope.Path, "actual.pdf");
            File.WriteAllBytes(expected, fixture.Expected);
            File.WriteAllBytes(actual, fixture.Actual);
            long retainedPngBytes = RetainedPngBytes(Assert.Single(fixture.Report.Pages));
            output.WriteLine($"Report-only retained PNG budget: {retainedPngBytes} bytes; paired pages: 1; unmatched selected pages: 99.");
            var request = new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.Compare,
                InputPath = expected,
                ComparisonPath = actual,
                ComparisonExpectedPages = fixture.ExpectedPages,
                ComparisonActualPages = fixture.ActualPages,
                Limits = new OfficeWorkflowLimits { MaximumOutputBytes = retainedPngBytes }
            };

            OfficeWorkflowResult exact = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.True(exact.Succeeded, exact.Summary);
            Assert.Equal("1", exact.HealthReport!.Metrics["pagesCompared"]);
            Assert.Equal("100", exact.HealthReport.Metrics[longerExpected ? "expectedSelectedPageCount" : "actualSelectedPageCount"]);
            Assert.Null(exact.OutputPath);
            Assert.Equal(0, exact.OutputBytes);

            request.Limits.MaximumOutputBytes = retainedPngBytes - 1;
            OfficeWorkflowResult bounded = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.Equal(OfficeWorkflowStatus.Failed, bounded.Status);
            OfficeWorkflowDiagnostic failure = Assert.Single(bounded.Diagnostics, diagnostic => diagnostic.Code == "WorkflowFailed");
            Assert.Equal(nameof(PdfReadLimitException), failure.Details["exceptionType"]);
            Assert.Contains("RenderBytes", failure.Message, StringComparison.Ordinal);
            Assert.Null(bounded.OutputPath);
            Assert.Equal(0, bounded.OutputBytes);
            Assert.Equal(fixture.Expected, File.ReadAllBytes(expected));
            Assert.Equal(fixture.Actual, File.ReadAllBytes(actual));
            Assert.Equal(2, Directory.GetFiles(scope.Path).Length);
        }

        private static (byte[] Expected, byte[] Actual, PdfPageSelector ExpectedPages, PdfPageSelector ActualPages,
            PdfVisualComparisonReport Report) CreateComparison(bool longerExpected) {
            byte[] onePage = CreatePages(1);
            byte[] hundredPages = CreatePages(100);
            byte[] expected = longerExpected ? hundredPages : onePage;
            byte[] actual = longerExpected ? onePage : hundredPages;
            PdfPageSelector expectedPages = PdfPageSelector.Parse(longerExpected ? "1-100" : "1");
            PdfPageSelector actualPages = PdfPageSelector.Parse(longerExpected ? "1" : "1-100");
            PdfVisualComparisonReport report = PdfVisualComparer.Compare(expected, actual, options: new PdfVisualComparisonOptions {
                ExpectedPages = expectedPages,
                ActualPages = actualPages
            });
            Assert.Single(report.Pages);
            Assert.Equal(99, (longerExpected ? report.UnmatchedExpectedPageNumbers : report.UnmatchedActualPageNumbers).Count);
            return (expected, actual, expectedPages, actualPages, report);
        }

        private static byte[] CreatePages(int count) => PdfDocument.Create(document => {
            for (int index = 0; index < count; index++) document.Page(page => page.Size(24, 24).Margin(0, 0, 0, 0));
        }).ToBytes();

        private static long RetainedPngBytes(PdfVisualPageComparison pair) => checked(
            pair.ExpectedPng.LongLength + pair.ActualPng.LongLength + pair.DiffPng.LongLength);

        private sealed class TestDirectory : IDisposable {
            public TestDirectory() {
                Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeimo-comparison-gallery-budget-" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(Path);
            }
            public string Path { get; }
            public void Dispose() => Directory.Delete(Path, recursive: true);
        }
    }
}
