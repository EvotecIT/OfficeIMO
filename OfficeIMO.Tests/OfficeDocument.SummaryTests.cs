using OfficeIMO.Excel;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeDocumentSummaryTests {
    private static OfficeInspectOptions Summary(bool detailed = false) => new() {
        IncludeSummary = true, IncludeContent = false, Detailed = detailed
    };

    [Fact]
    public void WorkbookInventoryCountsEmptySheetsAndRetainsNativeDetails() {
        string root = Directory.CreateTempSubdirectory("officeimo-summary-").FullName;
        try {
            string path = Path.Combine(root, "inventory.xlsx");
            using (var workbook = ExcelDocument.Create()) {
                workbook.AddSheet("Data").CellValue(1, 1, "Inventory");
                workbook.AddSheet("Empty");
                workbook.Save(path);
            }
            byte[] original = File.ReadAllBytes(path);
            OfficeDocumentInfo info = OfficeDocument.Inspect(path, Summary(detailed: true));
            Assert.Equal(ReaderInputKind.Excel, info.Format);
            Assert.Equal(ReaderInputKind.Excel, info.DetectedFormat);
            Assert.Equal(2, info.SheetCount);
            Assert.Null(info.PageCount);
            Assert.Null(info.SlideCount);
            Assert.False(info.IsEncrypted);
            Assert.NotNull(info.Modified);
            Assert.Null(info.Content);
            Assert.Equal(2, Assert.IsType<ExcelWorkbookInfo>(info.Details).Worksheets.Count);
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void LoadedWorkbookInspectionSeesUnsavedSheetsWithoutReopeningItsPath() {
        string root = Directory.CreateTempSubdirectory("officeimo-live-summary-").FullName;
        try {
            string path = Path.Combine(root, "removed.xlsx");
            using var workbook = ExcelDocument.Create();
            workbook.AddSheet("Saved");
            workbook.Save(path);
            workbook.AddSheet("Unsaved");
            File.Delete(path);
            OfficeDocumentInfo info = OfficeDocument.Inspect((object)workbook, Summary(detailed: true));
            Assert.Equal(2, info.SheetCount);
            Assert.Equal(new[] { "Saved", "Unsaved" }, Assert.IsType<ExcelWorkbookInfo>(info.Details).Worksheets.Select(sheet => sheet.Name));
            Assert.Null(info.Modified);
            workbook.AddSheet("Still open");
            Assert.Equal(3, workbook.Sheets.Count);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void PresentationSummaryIncludesUnsavedSlidesAndKeepsTheModelOpen() {
        using var presentation = PowerPointDocument.Create();
        presentation.AddSlide().AddTextBox("Opening");
        presentation.AddSlide().AddTextBox("Unsaved slide");
        OfficeDocumentInfo info = OfficeDocument.Inspect((object)presentation, Summary(detailed: true));
        Assert.Equal(ReaderInputKind.PowerPoint, info.Format);
        Assert.Equal(2, info.SlideCount);
        Assert.IsType<PowerPointInspectionReport>(info.Details);
        presentation.AddSlide();
        Assert.Equal(3, presentation.Slides.Count);
    }

    [Fact]
    public void ContentDetectionIdentifiesAnOfficePackageWithTheWrongExtension() {
        using var word = WordDocument.Create();
        word.AddParagraph("Word content");
        using var input = new MemoryStream(word.ToBytes());
        input.Position = 7;
        OfficeDocumentInfo info = OfficeDocument.Inspect(input, "attachment.bin", Summary(detailed: true));
        Assert.Equal(ReaderInputKind.Word, info.Format);
        Assert.Equal(ReaderInputKind.Word, info.DetectedFormat);
        Assert.Equal(ReaderDetectionConfidence.High, info.Confidence);
        Assert.IsType<WordDocumentInfo>(info.Details);
        Assert.Equal(7, input.Position);
        Assert.True(input.CanRead);
    }

    [Fact]
    public void EncryptedOfficeInventoryDoesNotDemandAPasswordOrPretendToKnowCounts() {
        string root = Directory.CreateTempSubdirectory("officeimo-encrypted-summary-").FullName;
        try {
            string path = Path.Combine(root, "protected.docx");
            using (var word = WordDocument.Create()) {
                word.AddParagraph("Protected content");
                word.Save(path, new WordSaveOptions { Password = "Synthetic-inspection-password" });
            }
            byte[] original = File.ReadAllBytes(path);
            OfficeDocumentInfo info = OfficeDocument.Inspect(path, Summary(detailed: true));
            Assert.True(info.IsEncrypted);
            Assert.Null(info.PageCount);
            Assert.Null(info.Content);
            Assert.Null(info.Details);
            Assert.Equal(original, File.ReadAllBytes(path));
            var authenticated = Summary(detailed: true);
            authenticated.Reader.OpenPassword = "Synthetic-inspection-password";
            OfficeDocumentInfo opened = OfficeDocument.Inspect(path, authenticated);
            Assert.True(opened.IsEncrypted);
            Assert.IsType<WordDocumentInfo>(opened.Details);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void SummaryHonorsInputAndNativePackageAdmissionLimitsWithoutChangingOptions() {
        using var word = WordDocument.Create();
        word.AddParagraph("Bounded native inspection");
        byte[] bytes = word.ToBytes();
        using var source = new MemoryStream(bytes);
        var options = Summary();
        options.MaximumInputBytes = bytes.Length - 1;
        Assert.Throws<InvalidDataException>(() => OfficeDocument.Inspect(source, "source.docx", options));
        options.MaximumInputBytes = bytes.Length;
        options.Reader.OpenXmlMaxPartUncompressedBytes = 32;
        Assert.Throws<ReaderResourceLimitException>(() => OfficeDocument.Inspect(source, "source.docx", options));
        Assert.Equal(32, options.Reader.OpenXmlMaxPartUncompressedBytes);
        Assert.Equal(0, source.Position);
    }

    [Fact]
    public void PdfSummaryUsesNativePreflightAndDoesNotRequireContentProjection() {
        using var document = PdfDocument.Create();
        document.Compose(builder => builder.Content(content => content.Text("PDF inspection")));
        OfficeDocumentInfo info = OfficeDocument.Inspect((object)document, Summary(detailed: true));
        Assert.Equal(1, info.PageCount);
        Assert.False(info.IsEncrypted);
        Assert.IsType<PdfDocumentInfo>(info.Details);
        Assert.Null(info.Content);
    }

    [Fact]
    public void LockedPdfInventoryRetainsItsNativeBlockersUntilAuthentication() {
        string root = Directory.CreateTempSubdirectory("officeimo-locked-pdf-summary-").FullName;
        try {
            string path = Path.Combine(root, "protected.pdf");
            using (var document = PdfDocument.Create(new PdfCreateOptions()
                .SetEncryption(new PdfStandardEncryptionOptions("synthetic-open")))) {
                document.Compose(builder => builder.Content(content => content.Text("Protected inspection")));
                document.Save(path, new PdfSaveOptions { AllowLoss = true });
            }
            byte[] original = File.ReadAllBytes(path);
            OfficeDocumentInfo locked = OfficeDocument.Inspect(path, Summary(detailed: true));
            Assert.True(locked.IsEncrypted);
            Assert.Null(locked.PageCount);
            Assert.False(Assert.IsType<PdfDocumentPreflight>(locked.Details).CanReadLogicalObjects);
            var options = Summary(detailed: true);
            options.Reader.OpenPassword = "synthetic-open";
            OfficeDocumentInfo opened = OfficeDocument.Inspect(path, options);
            Assert.True(opened.IsEncrypted);
            Assert.Equal(1, opened.PageCount);
            Assert.IsType<PdfDocumentInfo>(opened.Details);
            Assert.Null(opened.Content);
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void LoadedInspectionRejectsSourceReportsInsteadOfIgnoringThem() {
        using var markdown = MarkdownDocument.Parse("# Live content");
        var options = Summary();
        options.PackageSignatures = new OfficeIMO.Security.OfficePackageSignatureInspectionOptions();
        Assert.Throws<NotSupportedException>(() => OfficeDocument.Inspect((object)markdown, options));
        options.PackageSignatures = null;
        options.Detailed = true;
        Assert.Contains("Live content", Assert.IsType<string>(OfficeDocument.Inspect((object)markdown, options).Details));
    }

    [Fact]
    public void PathOnlyCustomHandlersRemainTheOwnerOfTheirInspectionProjection() {
        string root = Directory.CreateTempSubdirectory("officeimo-custom-summary-").FullName;
        try {
            string path = Path.Combine(root, "source.custom");
            File.WriteAllText(path, "Custom inventory");
            int reads = 0;
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
                Id = "summary.custom", DisplayName = "Custom inventory", Kind = ReaderInputKind.Text,
                Extensions = new[] { ".custom" },
                ReadPath = (name, options, token) => {
                    reads++;
                    token.ThrowIfCancellationRequested();
                    return new[] { new ReaderChunk { Text = File.ReadAllText(name), Kind = ReaderInputKind.Text } };
                }
            }).Build();
            var settings = Summary(detailed: true);
            settings.DocumentReader = reader;
            OfficeDocumentInfo info = OfficeDocument.Inspect(path, settings);
            Assert.Equal(1, reads);
            Assert.Equal(ReaderInputKind.Text, info.Format);
            Assert.Contains("Custom inventory", Assert.IsType<OfficeDocumentReadResult>(info.Details).GetText());
            Assert.Null(info.Content);
        } finally { Directory.Delete(root, true); }
    }
}
