using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.Pdf;
using OpenMcdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class LegacyDocPdfStageTests {
    [Fact]
    public void ExplicitLegacyLossPolicyBlocksOutputAndAcceptedLossStaysVisible() {
        byte[] source = CreateLegacySource();
        using var input = new MemoryStream(source);
        using var destination = new MemoryStream();
        var settings = new OfficeConversionOptions {
            SourceName = "source.doc", WordLoad = new WordLoadOptions { ReportUnsupportedContent = false, FailOnLoss = true }
        };
        OfficeConversionException rejection = Assert.Throws<OfficeConversionException>(() =>
            OfficeConverter.Convert(input, OfficeDocumentFormat.Word, destination, OfficeDocumentFormat.Pdf, settings));
        Assert.True(rejection.Report.HasLoss);
        Assert.Equal(0, destination.Length);
        settings.WordLoad.FailOnLoss = false;
        OfficeConversionReport accepted = OfficeConverter.Convert(input, OfficeDocumentFormat.Word,
            destination, OfficeDocumentFormat.Pdf, settings);
        Assert.True(accepted.HasLoss);
        Assert.Contains(accepted.Diagnostics, finding => finding.Format == "Word" && finding.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains(accepted.Diagnostics, finding => finding.Code == OfficeDiagnosticCodes.FontFamilySubstitution);
        using var pdf = PdfDocument.Load(new MemoryStream(destination.ToArray()));
        Assert.Contains("Preserved DOC text", pdf.ExtractText());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RejectedImportAndFailedSerializationRetainCanonicalImportEvidence(bool allowImport) {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-doc-failure-" + Guid.NewGuid().ToString("N"))).FullName;
        try {
            string input = Path.Combine(root, "source.doc"), output = Path.Combine(root, "result.pdf");
            File.WriteAllBytes(input, CreateLegacySource());
            byte[] sentinel = [1, 2, 3]; File.WriteAllBytes(output, sentinel);
            OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new() {
                Operation = OfficeWorkflowOperation.Convert, InputPath = input, OutputPath = output, ConversionRouteId = "doc-pdf",
                ConflictPolicy = OfficePublicationPolicy.Replace,
                ConversionOptions = new() {
                WordLoad = new OfficeIMO.Word.WordLoadOptions { FailOnLoss = (allowImport ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block) == OfficeConversionLossPolicy.Block }
            },
                Limits = new() { MaximumOutputBytes = allowImport ? 1 : 10_000_000 }
            });
            Assert.False(result.Succeeded);
            Assert.Equal(sentinel, File.ReadAllBytes(output));
            var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            var imports = evidence.FidelityDiagnostics.Where(finding => finding.Source == "Word").ToArray();
            Assert.NotEmpty(imports);
            foreach (var finding in imports)
                Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == finding.Code && diagnostic.Message == finding.Message
                    && diagnostic.Stage == "import" && diagnostic.Details["source"] == finding.Source
                    && diagnostic.Details["lossKind"] == finding.LossKind.ToString()
                    && (diagnostic.Details.TryGetValue("location", out string? location) ? location : null) == finding.Location);
        } finally { Directory.Delete(root, true); }
    }

    internal static byte[] CreateLegacySource() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-legacy-pdf-" + Guid.NewGuid().ToString("N") + ".doc");
        byte[] source;
        try {
            using var word = WordDocument.Create();
            // Force a rendering-stage approximation independently of installed Office fonts.
            word.AddParagraph("Preserved DOC text").FontFamily = "OfficeIMO Missing Workflow Stage Font";
            word.Save(path);
            using var package = new MemoryStream();
            package.Write(File.ReadAllBytes(path)); package.Position = 0;
            using (RootStorage root = RootStorage.Open(package, StorageModeFlags.LeaveOpen)) {
                root.CreateStorage("_VBA_PROJECT_CUR"); root.CreateStorage("ObjectPool"); root.Flush(true);
            }
            source = package.ToArray();
        } finally { File.Delete(path); }
        return source;
    }
}
