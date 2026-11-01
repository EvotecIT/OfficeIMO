using System.Security.Cryptography;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Workflows.Tests;

public sealed class ScanCleanupTests {
    [Fact]
    public void ScanDiagnosticBudgetsAreSnapshottedAndMustBePositive() {
        var options = new OfficeScanCleanupOptions {
            AcknowledgeRasterOutput = true,
            MaximumDiagnostics = 20_000,
            MaximumDiagnosticCharacters = 16L * 1024L * 1024L
        };
        OfficeScanCleanupOptions snapshot = options.Snapshot();
        options.MaximumDiagnostics = 1;
        options.MaximumDiagnosticCharacters = 1;
        Assert.Equal(20_000, snapshot.MaximumDiagnostics);
        Assert.Equal(16L * 1024L * 1024L, snapshot.MaximumDiagnosticCharacters);
        options.MaximumDiagnostics = 0;
        Assert.Throws<ArgumentOutOfRangeException>(() => options.Snapshot());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProviderScanCopyNeverOpensTheSourceForWriting(bool deferred) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-scan-provider-" + Guid.NewGuid().ToString("N"));
        const string sourceLocation = "content://scan/source";
        byte[] source = PdfDocument.Create(c => c.Page(p => p.Size(global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(100)).Margin(global::OfficeIMO.OfficeLength.Points(0)))).ToBytes();
        int writes = 0;
        try {
            var recovery = new OfficeWorkflowOutputRecoveryStore(root);
            var request = new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.ScanCleanup,
                InputPath = sourceLocation,
                InputStream = new("source.pdf", _ => Task.FromResult<Stream>(new MemoryStream(source))),
                OutputPath = deferred ? "content://scan/folder" : sourceLocation,
                ConflictPolicy = OfficePublicationPolicy.Replace,
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
                compose.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(300)).Content(c => c.Item(i => i.Paragraph(p => p.Text("First page")))));
                compose.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(200)).Content(c => c.Item(i => i.Paragraph(p => p.Text("Second page")))));
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
                ConflictPolicy = OfficePublicationPolicy.Replace,
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

    [Theory]
    [InlineData("user-open", PdfStandardPermissions.All, "user-open")]
    [InlineData("user-open", PdfStandardPermissions.Print, "owner-only")]
    [InlineData("", PdfStandardPermissions.Print, "owner-only")]
    public async Task ProtectedSourceKeepsItsSecurityInTheRasterCopy(string userPassword, PdfStandardPermissions allowed, string suppliedPassword) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-scan-protected-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            const string ownerPassword = "owner-only";
            PdfStandardEncryptionOptions encryption = userPassword.Length == 0
                ? PdfStandardEncryptionOptions.ForPermissions(ownerPassword, allowed)
                : new PdfStandardEncryptionOptions(userPassword) { OwnerPassword = ownerPassword, AllowedPermissions = allowed };
            byte[] source = PdfDocument.Create(compose => {
                compose.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(300)));
                compose.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(200)));
            }).Security.Encrypt(encryption).Pdf;
            string input = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "cleaned.pdf");
            await File.WriteAllBytesAsync(input, source);

            OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.ScanCleanup, InputPath = input, OutputPath = output, PdfPassword = suppliedPassword,
                ScanCleanup = new() { AcknowledgeRasterOutput = true, Preparation = new() { Dpi = 72 } }
            });

            Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary);
            byte[] cleaned = await File.ReadAllBytesAsync(output);
            var ownerOptions = new PdfLoadOptions { Password = ownerPassword };
            PdfDocumentInfo owner = PdfDocument.Load(cleaned, ownerOptions).Inspect(ownerOptions);
            Assert.True(owner.Security.HasEncryption);
            Assert.True(owner.Security.HasOwnerAuthorization);
            Assert.Equal(allowed, owner.Security.AllowedStandardPermissions);
            Assert.Equal(2, owner.PageCount);
            if (userPassword.Length > 0) {
                Assert.Throws<PdfPasswordRequiredException>(() => PdfDocument.Load(cleaned).Inspect());
                var userOptions = new PdfLoadOptions { Password = userPassword, PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions };
                Assert.Equal(OfficeIMO.Pdf.PdfPasswordAuthenticationRole.User, PdfDocument.Load(cleaned, userOptions).Inspect(userOptions).Security.PasswordAuthenticationRole);
            }
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task ProtectedSourceWhoseSecurityCannotBeRetainedPublishesNothing() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-scan-unretained-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            byte[] source = PdfDocument.Create(compose => compose.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(300))))
                .Security.Encrypt(new PdfStandardEncryptionOptions("user-open")).Pdf;
            byte[] marker = System.Text.Encoding.ASCII.GetBytes("/ID");
            int identifier = source.AsSpan().LastIndexOf(marker);
            Assert.True(identifier > 0);
            source[identifier + 1] = (byte)'X';
            string input = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "cleaned.pdf");
            await File.WriteAllBytesAsync(input, source);

            OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.ScanCleanup, InputPath = input, OutputPath = output, PdfPassword = "user-open",
                ScanCleanup = new() { AcknowledgeRasterOutput = true, Preparation = new() { Dpi = 72 } }
            });

            Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
            Assert.Contains("retain encryption", result.Summary);
            Assert.False(File.Exists(output));
        } finally { Directory.Delete(root, true); }
    }
}
