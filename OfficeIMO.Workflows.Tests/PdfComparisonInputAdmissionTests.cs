using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Workflows.Tests {
    public sealed class PdfComparisonInputAdmissionTests {
        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public async Task SelectedComparisonReportsReadFailureBeforeResolvingPages(bool unreadableActual, bool encrypted) {
            string root = Path.Combine(Path.GetTempPath(), "officeimo-comparison-admission-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try {
                string expected = Path.Combine(root, "expected.pdf");
                string actual = Path.Combine(root, "actual.pdf");
                string output = Path.Combine(root, "report.html");
                byte[] readable = PdfDocument.Create(document => document.Page(page => page.Size(150, 180))).ToBytes();
                byte[] unreadable = encrypted
                    ? PdfDocument.Load(readable).Security.Encrypt(new("reader") { OwnerPassword = "owner" }).Pdf
                    : Encoding.ASCII.GetBytes("%PDF-1.7\nbroken");
                byte[] expectedBytes = unreadableActual ? readable : unreadable;
                byte[] actualBytes = unreadableActual ? unreadable : readable;
                File.WriteAllBytes(expected, expectedBytes);
                File.WriteAllBytes(actual, actualBytes);
                File.WriteAllText(output, "existing report");

                OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                    Operation = OfficeWorkflowOperation.Compare,
                    InputPath = expected,
                    ComparisonPath = actual,
                    PdfPassword = encrypted && !unreadableActual ? "wrong" : null,
                    ComparisonPdfPassword = encrypted && unreadableActual ? "wrong" : null,
                    ComparisonExpectedPages = PdfPageSelector.Parse("1"),
                    ComparisonActualPages = PdfPageSelector.Parse("1"),
                    OutputPath = output,
                    ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
                });

                Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
                OfficeWorkflowDiagnostic failure = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "WorkflowFailed");
                Assert.Equal(encrypted ? nameof(PdfInvalidPasswordException) : nameof(PdfParseException), failure.Details["exceptionType"]);
                Assert.Equal("existing report", File.ReadAllText(output));
                Assert.Equal(expectedBytes, File.ReadAllBytes(expected));
                Assert.Equal(actualBytes, File.ReadAllBytes(actual));
                Assert.Null(result.OutputPath);
                Assert.Equal(0, result.OutputBytes);
            } finally {
                Directory.Delete(root, recursive: true);
            }
        }

        [Theory]
        [InlineData(null)]
        [InlineData(false)]
        [InlineData(true)]
        public async Task ZeroPageCatalogsCompareWithoutSelectorsAndRejectSelectedPages(bool? selectActual) {
            string root = Path.Combine(Path.GetTempPath(), "officeimo-zero-page-comparison-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try {
                string expected = Path.Combine(root, "expected.pdf");
                string actual = Path.Combine(root, "actual.pdf");
                string output = Path.Combine(root, "report.html");
                byte[] zeroPages = Encoding.ASCII.GetBytes(
                    "%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
                    "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n" +
                    "trailer\n<< /Root 1 0 R /Size 3 >>\n%%EOF\n");
                PdfDocumentPreflight preflight = PdfDocument.Preflight(zeroPages);
                Assert.False(preflight.CanRead);
                Assert.All(preflight.ReadBlockers, blocker => Assert.Equal(PdfReadBlockerKind.NoPages, blocker.Kind));
                File.WriteAllBytes(expected, zeroPages);
                File.WriteAllBytes(actual, zeroPages);
                File.WriteAllText(output, "existing report");

                OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                    Operation = OfficeWorkflowOperation.Compare,
                    InputPath = expected,
                    ComparisonPath = actual,
                    ComparisonExpectedPages = selectActual == false ? PdfPageSelector.Parse("1") : null,
                    ComparisonActualPages = selectActual == true ? PdfPageSelector.Parse("1") : null,
                    OutputPath = output,
                    ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
                });

                if (selectActual is null) {
                    Assert.True(result.Succeeded, result.Summary);
                    Assert.True(result.HealthReport!.Verified);
                    Assert.Equal("0", result.HealthReport.Metrics["pagesCompared"]);
                    Assert.True(File.Exists(output));
                    Assert.NotEqual("existing report", File.ReadAllText(output));
                } else {
                    Assert.Equal(OfficeWorkflowFailureKind.ValidationFailed, result.FailureKind);
                    Assert.Equal("existing report", File.ReadAllText(output));
                }
                Assert.Equal(zeroPages, File.ReadAllBytes(expected));
                Assert.Equal(zeroPages, File.ReadAllBytes(actual));
            } finally {
                Directory.Delete(root, recursive: true);
            }
        }
    }
}
