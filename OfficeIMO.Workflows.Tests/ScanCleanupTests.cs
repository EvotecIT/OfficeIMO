using System.Security.Cryptography;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Workflows.Tests;

public sealed class ScanCleanupTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProviderScanCopyNeverOpensTheSourceForWriting(bool deferred) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-scan-provider-" + Guid.NewGuid().ToString("N"));
        const string sourceLocation = "content://scan/source";
        byte[] source = PdfDocument.Create(c => c.Page(p => p.Size(200, 100).Margin(0))).ToBytes();
        int writes = 0;
        try {
            var recovery = new OfficeWorkflowOutputRecoveryStore(root);
            var request = new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.ScanCleanup,
                InputPath = sourceLocation,
                InputStream = new("source.pdf", _ => Task.FromResult<Stream>(new MemoryStream(source))),
                OutputPath = deferred ? "content://scan/folder" : sourceLocation,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                OutputStream = new("copy.pdf", _ => Task.FromResult<Stream>(new MemoryStream(source)),
                    _ => { writes++; return Task.FromResult<Stream>(new MemoryStream()); }, recovery,
                    deferred ? _ => Task.FromResult(sourceLocation) : null),
                ScanCleanup = new() { AcknowledgeRasterOutput = true, Preparation = new() { Dpi = 72 } }
            };
            OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.Equal(deferred ? OfficeWorkflowStatus.Unconfirmed : OfficeWorkflowStatus.Failed, result.Status);
            Assert.Equal(0, writes);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [Fact]
    public async Task ReviewedRasterCopyKeepsSelectedPageAndRequiresMatchingSource() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-scan-workflow-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "cleaned.pdf");
            var document = PdfDocument.Create(compose => {
                compose.Page(page => page.Size(200, 300).Content(c => c.Item(i => i.Paragraph(p => p.Text("First page")))));
                compose.Page(page => page.Size(300, 200).Content(c => c.Item(i => i.Paragraph(p => p.Text("Second page")))));
            });
            byte[] source = document.ToBytes(); await File.WriteAllBytesAsync(input, source);
            var settings = new OfficeScanCleanupOptions {
                AcknowledgeRasterOutput = true,
                ExpectedSourceSha256 = Convert.ToHexString(SHA256.HashData(source)),
                Preparation = new() { Dpi = 72, ReadOptions = new() { PageSelection = PdfPageSelection.From(2) }, ScanProcessing = new() { Deskew = false, NormalizeBackground = false } }
            };
            var request = new OfficeWorkflowRequest { Operation = OfficeWorkflowOperation.ScanCleanup, InputPath = input, OutputPath = output, ScanCleanup = settings };
            var result = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.Equal(OfficeWorkflowStatus.Completed, result.Status);
            var cleaned = PdfDocument.Load(await File.ReadAllBytesAsync(output));
            Assert.Equal(1, cleaned.Inspect().PageCount); Assert.Equal(300, cleaned.Render.Drawing(1).Width);
            Assert.True(string.IsNullOrWhiteSpace(cleaned.Reader.Text())); Assert.Single(cleaned.Images.Placements());
            Assert.Equal(source, await File.ReadAllBytesAsync(input));
            File.Delete(output);
            var overwrite = new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.ScanCleanup,
                InputPath = input,
                OutputPath = input,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                ScanCleanup = settings
            };
            Assert.NotEqual(OfficeWorkflowStatus.Completed, (await new OfficeWorkflowRunner().RunAsync(overwrite)).Status);
            Assert.Equal(source, await File.ReadAllBytesAsync(input));
            settings.Preparation.ReadOptions = new() { PageSelection = PdfPageSelection.From(1, 2) };
            settings.Preparation.Regions = new[] { new PdfOcrPageRegion(2, 0, 0, 0.5, 1) };
            var cropped = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.Equal(OfficeWorkflowStatus.Completed, cropped.Status);
            var cropDocument = PdfDocument.Load(await File.ReadAllBytesAsync(output));
            Assert.Equal(1, cropDocument.Inspect().PageCount);
            Assert.Equal(150, cropDocument.Render.Drawing(1).Width);
            File.Delete(output);
            settings.ExpectedSourceSha256 = new string('0', 64);
            var stale = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.Equal(OfficeWorkflowStatus.Failed, stale.Status); Assert.Contains("changed since scan review", stale.Summary); Assert.False(File.Exists(output));
            settings.AcknowledgeRasterOutput = false;
            Assert.Equal(OfficeWorkflowStatus.Failed, (await new OfficeWorkflowRunner().RunAsync(request)).Status);
            Assert.False(File.Exists(output));
        } finally { Directory.Delete(root, true); }
    }
}
