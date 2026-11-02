using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterNativeRouteTests {
    [Theory]
    [InlineData("source.docm", OfficeDocumentFormat.Word, ".docx")]
    [InlineData("source.xlsb", OfficeDocumentFormat.Excel, ".xlsx")]
    [InlineData("source.ppt", OfficeDocumentFormat.PowerPoint, ".pptx")]
    [InlineData("source.docx", OfficeDocumentFormat.Word, ".docm")]
    [InlineData("source.xlsx", OfficeDocumentFormat.Excel, ".xlsb")]
    [InlineData("source.pptx", OfficeDocumentFormat.PowerPoint, ".pptm")]
    public async Task BatchSelectionReachesNativeVariantCodecs(string name, OfficeDocumentFormat format, string extension) {
        string root = Directory.CreateTempSubdirectory("officeimo-native-route-test-").FullName;
        try {
            string source = Path.Combine(root, name), output = Path.Combine(root, "output");
            switch (format) {
                case OfficeDocumentFormat.Word:
                    using (var document = WordDocument.Create()) { document.AddParagraph("Native variant content"); document.Save(source); }
                    break;
                case OfficeDocumentFormat.Excel:
                    using (var document = ExcelDocument.Create()) { document.AddSheet("Data").CellValue(1, 1, "Native variant content"); document.Save(source); }
                    break;
                case OfficeDocumentFormat.PowerPoint:
                    using (var document = PowerPointDocument.Create()) { document.AddSlide().AddTextBox("Native variant content"); document.Save(source); }
                    break;
            }
            byte[] original = File.ReadAllBytes(source);
            var settings = new OfficeBatchOptions { InputPaths = [source], OutputDirectory = output, TargetExtension = extension };
            OfficeBatchPlan plan = await OfficeConverter.PlanAsync(settings);
            Assert.True(plan.CanRun);
            Assert.Equal(Path.Combine(output, "source" + extension), Assert.Single(plan.Items).OutputPath);
            OfficeBatchReport report = await OfficeConverter.ConvertAsync([source], output, format, settings);
            Assert.True(report.Succeeded, string.Join("; ", report.Items.Select(item => item.Summary)));
            string artifact = Assert.Single(report.Items).OutputPath!;
            Assert.Equal(Path.Combine(output, "source" + extension), artifact);
            Assert.Equal(original, File.ReadAllBytes(source));
            Assert.True(new FileInfo(artifact).Length > 0);
            switch (format) {
                case OfficeDocumentFormat.Word:
                    using (var document = WordDocument.Load(artifact)) Assert.Contains(document.Paragraphs, paragraph => paragraph.Text == "Native variant content");
                    break;
                case OfficeDocumentFormat.Excel:
                    using (var document = ExcelDocument.Load(artifact)) { Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string text)); Assert.Equal("Native variant content", text); }
                    break;
                case OfficeDocumentFormat.PowerPoint:
                    using (var document = PowerPointDocument.Load(artifact)) Assert.Single(document.Slides);
                    break;
            }
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task NativeAnalysisRetainsExistingDestinationAndReportsNoWrites() {
        string root = Directory.CreateTempSubdirectory("officeimo-native-names-test-").FullName;
        try {
            string source = Path.Combine(root, "Report.docx");
            string target = Path.Combine(Directory.CreateDirectory(Path.Combine(root, "output")).FullName, "Report.docm");
            using (var document = WordDocument.Create()) {
                document.AddParagraph("Native analysis content");
                document.Save(source);
            }
            File.WriteAllText(target, "Keep this destination during analysis");
            OfficeConversionReport assessment = await OfficeConverter.AnalyzeAsync(source, target);
            Assert.Equal(0, assessment.ItemsWritten);
            Assert.Equal("Keep this destination during analysis", File.ReadAllText(target));
            OfficeConversionReport result = await OfficeConverter.ConvertAsync(source, target,
                new OfficeConversionOptions { FileConflictPolicy = OfficePublicationPolicy.Replace });
            Assert.Equal(1, result.ItemsWritten);
            using var reopened = WordDocument.Load(target);
            Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text.Contains("Native analysis content", StringComparison.Ordinal));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void SourcePasswordAndInputLimitApplyToEncryptedPdfReconstruction() {
        using var pdf = PdfDocument.Create(new PdfCreateOptions().SetEncryption(new PdfStandardEncryptionOptions("synthetic-open")));
        pdf.Compose(document => document.Content(content => content.Text("Encrypted PDF input")));
        byte[] bytes = pdf.ToBytes(new PdfSaveOptions { AllowLoss = true });
        using var source = new MemoryStream(bytes);
        using var target = new MemoryStream();
        OfficeConverter.Convert(source, OfficeDocumentFormat.Pdf, target, OfficeDocumentFormat.Word,
            new OfficeConversionOptions { SourcePassword = "synthetic-open" });
        using var word = WordDocument.Load(new MemoryStream(target.ToArray()));
        Assert.Contains(word.Paragraphs, paragraph => paragraph.Text.Contains("Encrypted PDF input", StringComparison.Ordinal));
        using var rejectedTarget = new MemoryStream();
        Assert.Throws<InvalidDataException>(() => OfficeConverter.Convert(source, OfficeDocumentFormat.Pdf, rejectedTarget, OfficeDocumentFormat.Word,
            new OfficeConversionOptions { SourcePassword = "synthetic-open", MaximumInputBytes = bytes.Length - 1 }));
        Assert.Equal(0, rejectedTarget.Length);
        using var policyRejectedTarget = new MemoryStream();
        var loadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = 32 } };
        PdfReadLimitException policyError = Assert.Throws<PdfReadLimitException>(() => OfficeConverter.Convert(source,
            OfficeDocumentFormat.Pdf, policyRejectedTarget, OfficeDocumentFormat.Word,
            new OfficeConversionOptions { SourcePassword = "synthetic-open", PdfLoad = loadOptions }));
        Assert.Equal(PdfReadLimitKind.InputBytes, policyError.Kind);
        Assert.Equal(32, policyError.Limit);
        Assert.Equal(bytes.LongLength, policyError.Actual);
        Assert.Equal(0, policyRejectedTarget.Length);
        Assert.Equal(32, loadOptions.Limits.MaxInputBytes);
        Assert.Null(loadOptions.Password);
    }
}
