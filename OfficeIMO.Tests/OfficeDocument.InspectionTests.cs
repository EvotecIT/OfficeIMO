using OfficeIMO.Excel;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeDocumentInspectionTests {
    [Theory]
    [InlineData("source.docx", ReaderInputKind.Word)]
    [InlineData("source.xlsx", ReaderInputKind.Excel)]
    [InlineData("source.pptx", ReaderInputKind.PowerPoint)]
    [InlineData("source.pdf", ReaderInputKind.Pdf)]
    [InlineData("source.md", ReaderInputKind.Markdown)]
    public void InspectionUsesTheNativeHandlerAndPreservesItsSource(string name, ReaderInputKind expected) {
        string root = Directory.CreateTempSubdirectory("officeimo-inspection-test-").FullName;
        try {
            string path = Path.Combine(root, name);
            switch (expected) {
                case ReaderInputKind.Word:
                    using (var document = WordDocument.Create()) { document.AddParagraph("Inspection content"); document.Save(path); }
                    break;
                case ReaderInputKind.Excel:
                    using (var document = ExcelDocument.Create()) { document.AddWorksheet("Data").CellValue(1, 1, "Inspection content"); document.Save(path); }
                    break;
                case ReaderInputKind.PowerPoint:
                    using (var document = PowerPointPresentation.Create()) { document.AddSlide().AddTextBox("Inspection content"); document.Save(path); }
                    break;
                case ReaderInputKind.Pdf:
                    using (var document = PdfDocument.Create()) {
                        document.Compose(builder => builder.Content(content => content.Text("Inspection content")));
                        document.Save(path, new PdfSaveOptions { AllowLoss = true });
                    }
                    break;
                case ReaderInputKind.Markdown:
                    using (var document = MarkdownDocument.Parse("# Inspection content")) document.Save(path);
                    break;
            }
            byte[] original = File.ReadAllBytes(path);
            OfficeDocumentInfo inspection = OfficeDocument.Inspect(path);
            OfficeDocumentReadResult result = Assert.IsType<OfficeDocumentReadResult>(inspection.Content);
            Assert.Equal(expected, result.Kind);
            Assert.NotEmpty(result.Blocks);
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally { Directory.Delete(root, true); }
    }
    [Fact]
    public void SelectedSafetyInspectionDoesNotLoadContentOrRewriteTheSource() {
        string root = Directory.CreateTempSubdirectory("officeimo-safety-inspection-").FullName;
        try {
            string path = Path.Combine(root, "source.docx");
            using (var document = WordDocument.Create()) {
                document.AddParagraph("Inspection evidence");
                document.Save(path);
            }
            byte[] original = File.ReadAllBytes(path);
            var options = new OfficeInspectOptions {
                IncludeContent = false,
                ContentSafety = new OfficeIMO.ContentSafety.OfficeContentSafetyOptions(),
                Provenance = new OfficeIMO.Provenance.OfficeProvenanceOptions(),
                PackageSignatures = new OfficeIMO.Security.OfficePackageSignatureInspectionOptions()
            };
            options.Reader.MaxInputBytes = 1;
            OfficeDocumentInfo result = OfficeDocument.Inspect(path, options);
            Assert.Null(result.Content);
            Assert.NotNull(result.ContentSafety);
            Assert.NotNull(result.Provenance);
            Assert.NotNull(result.PackageSignatures);
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void UploadedPackageInspectionPreservesStreamPositionAndUsesTheConfiguredLimit() {
        using var word = WordDocument.Create();
        word.AddParagraph("Uploaded inspection");
        byte[] bytes = word.ToBytes();
        using var source = new MemoryStream(bytes);
        source.Position = 7;
        var options = new OfficeInspectOptions {
            IncludeContent = false,
            Provenance = new OfficeIMO.Provenance.OfficeProvenanceOptions(),
            PackageSignatures = new OfficeIMO.Security.OfficePackageSignatureInspectionOptions()
        };
        OfficeDocumentInfo result = OfficeDocument.Inspect(source, "upload.docx", options);
        Assert.NotNull(result.Provenance);
        Assert.NotNull(result.PackageSignatures);
        Assert.Null(result.Content);
        Assert.Equal(7, source.Position);
        Assert.True(source.CanRead);
        options.MaximumInputBytes = bytes.Length - 1;
        Assert.Throws<InvalidDataException>(() => OfficeDocument.Inspect(source, "upload.docx", options));
        Assert.Equal(7, source.Position);
        Assert.Equal(bytes, source.ToArray());
    }

    [Fact]
    public void UnsupportedPdfPreflightRetainsBlockersWithoutDemandingContentExtraction() {
        using var source = new MemoryStream(Encoding.ASCII.GetBytes("%PDF-1.7\n"));
        OfficeDocumentInfo result = OfficeDocument.Inspect(source, "unsupported.pdf", new OfficeInspectOptions {
            IncludeContent = false, PdfPreflight = PdfLoadOptions.UntrustedDefaults
        });
        Assert.Null(result.Content);
        Assert.NotNull(result.PdfPreflight);
        Assert.False(result.PdfPreflight.CanRead);
        Assert.NotEmpty(result.PdfPreflight.Diagnostics);
        Assert.True(source.CanRead);
        Assert.Equal(0, source.Position);
    }

}
