using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.Pdf;
using OpenMcdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class LegacyDocPdfStageTests {
    [Fact]
    public void UnsupportedLegacyStorageBlocksByDefaultAndRemainsVisibleWhenAccepted() {
        byte[] source = CreateLegacySource();
        var settings = new LegacyDocImportOptions { ReportUnsupportedContent = false };
        var rejection = Assert.Throws<OfficeConversionException>(() => LegacyDocPdfConverter.ToPdfDocumentResult(new MemoryStream(source), importOptions: settings));
        Assert.True(rejection.Report.HasLoss);
        var accepted = LegacyDocPdfConverter.ToPdfDocumentResult(new MemoryStream(source), importOptions: settings,
            lossPolicy: OfficeConversionLossPolicy.Allow);
        Assert.True(Assert.Single(accepted.SourceConversionReports).HasLoss);
        using var destination = new MemoryStream();
        PdfSaveResult output = accepted.SaveResult(destination);
        output.RequireSuccess();
        Assert.True(output.HasLoss);
        Assert.Equal(2, output.ConversionReports.Count);
        Assert.Contains(output.Report.Warnings, warning => warning.Code == "NativeFontFamilySubstituted");
        Assert.Contains(accepted.FidelityDiagnostics, finding => finding.LossKind != OfficeConversionLossKind.None);
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
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                ConversionOptions = new() { LegacyDocLossPolicy = allowImport ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block },
                Limits = new() { MaximumOutputBytes = allowImport ? 1 : 10_000_000 }
            });
            Assert.False(result.Succeeded);
            Assert.Equal(sentinel, File.ReadAllBytes(output));
            var evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            var imports = evidence.FidelityDiagnostics.Where(finding => finding.Source.StartsWith("OfficeIMO.Word.LegacyDoc", StringComparison.Ordinal)).ToArray();
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
            using var word = WordDocument.Create(); word.AddParagraph("Preserved DOC text"); word.Save(path);
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
