using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Word;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderNativeLeanProjectionTests {
    private static byte[] Pixel => Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Theory]
    [InlineData("word", "source.docx")]
    [InlineData("excel", "source.xlsx")]
    [InlineData("powerpoint", "source.pptx")]
    public void LeanProjectionAvoidsUnrequestedImageReadsAndMarkdown(string format, string name) {
        byte[] bytes = Create(format);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler(new ReaderWordOptions { IncludePageLocations = true })
            .AddExcelHandler().AddPowerPointHandler().Build();
        var options = new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata, MaxOpenXmlImageAssetBytes = 1,
            Inspection = new ReaderInspectionOptions { IncludeContentSafety = false, IncludePackageContent = false }
        };
        OfficeDocumentReadResult result = reader.ReadDocument(bytes, name, options);
        Assert.Contains(result.Chunks, chunk => chunk.Text!.Contains("Body evidence", StringComparison.Ordinal));
        Assert.All(result.Chunks, chunk => Assert.True(string.IsNullOrEmpty(chunk.Markdown)));
        Assert.Empty(result.Assets);
        Assert.Empty(result.Visuals);
        Assert.Contains("officeimo.reader.projection.text-and-metadata", result.CapabilitiesUsed);
        if (format == "word") Assert.Empty(result.Pages);
        else Assert.NotEmpty(result.Pages);
        Assert.Equal(ReaderProjection.TextAndMetadata, options.Projection);
        Assert.Equal(1, options.MaxOpenXmlImageAssetBytes);
        options.Projection = ReaderProjection.Full;
        IOException limit = Assert.Throws<IOException>(() => reader.ReadDocument(bytes, name, options));
        Assert.Contains("MaxOpenXmlImageAssetBytes", limit.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void WordNativeSnapshotRetainsImageDescriptorsWithoutCopyingBytes() {
        using WordDocument document = WordDocument.Create();
        using var image = new MemoryStream(Pixel, writable: false);
        document.AddParagraph("Body evidence").AddImage(image, "pixel.png", OfficeIMO.OfficeLength.Pixels(1), OfficeIMO.OfficeLength.Pixels(1), options: new OfficeIMO.Word.WordImageOptions { Description = "Source alt text" });
        WordDocumentInfo lean = document.Inspect(includeImageBytes: false);
        WordRunInfo descriptor = lean.Sections.SelectMany(section => section.Elements).OfType<WordParagraphInfo>()
            .SelectMany(paragraph => paragraph.Runs).Single(run => run.InlineImage != null);
        Assert.Null(descriptor.InlineImage!.Bytes);
        Assert.Equal("Source alt text", descriptor.InlineImage.Description);
        Assert.Equal("pixel.png", descriptor.InlineImage.FileName);
        WordDocumentInfo full = document.Inspect();
        Assert.Equal(Pixel, full.Sections.SelectMany(section => section.Elements).OfType<WordParagraphInfo>()
            .SelectMany(paragraph => paragraph.Runs).Single(run => run.InlineImage != null).InlineImage!.Bytes);
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.Inspect(false, canceled.Token));
    }

    [Fact]
    public void PowerPointPlainExtractionSnapshotsOptionsAndKeepsCompleteText() {
        using PowerPointDocument document = PowerPointDocument.Create();
        document.AddSlide().AddTextBox(new string('x', 600), 100000, 100000, 2000000, 500000);
        var extraction = new PowerPointExtractionExtensions.PowerPointExtractOptions { IncludeMarkdown = false };
        var chunking = new PowerPointExtractChunkingOptions { MaxChars = 256 };
        var pending = document.ExtractMarkdownChunks(extraction, chunking);
        extraction.IncludeMarkdown = true; chunking.MaxChars = 1000;
        PowerPointExtractChunk[] chunks = pending.ToArray();
        Assert.Equal(3, chunks.Length);
        Assert.Equal(new string('x', 600), string.Concat(chunks.Select(chunk => chunk.Text)));
        Assert.All(chunks, chunk => Assert.Null(chunk.Markdown));
        Assert.Equal(1000, chunking.MaxChars);
        Assert.NotNull(document.ExtractMarkdownChunks().Single().Markdown);
    }

    [Theory]
    [InlineData("word", "source.docx")]
    [InlineData("excel", "source.xlsx")]
    [InlineData("powerpoint", "source.pptx")]
    public void LeanPackageMetadataDoesNotDecodeTheVbaPayload(string format, string name) {
        using var bytes = new MemoryStream();
        byte[] original = Create(format); bytes.Write(original, 0, original.Length); bytes.Position = 0;
        using var rawVba = new MemoryStream(new byte[] { 1, 2, 3, 4 }, writable: false);
        if (format == "word") {
            using WordprocessingDocument package = WordprocessingDocument.Open(bytes, true);
            package.MainDocumentPart!.AddNewPart<VbaProjectPart>().FeedData(rawVba);
            package.ChangeDocumentType(WordprocessingDocumentType.MacroEnabledDocument);
        } else if (format == "excel") {
            using SpreadsheetDocument package = SpreadsheetDocument.Open(bytes, true);
            package.WorkbookPart!.AddNewPart<VbaProjectPart>().FeedData(rawVba);
            package.ChangeDocumentType(SpreadsheetDocumentType.MacroEnabledWorkbook);
        } else {
            using PresentationDocument package = PresentationDocument.Open(bytes, true);
            package.PresentationPart!.AddNewPart<VbaProjectPart>().FeedData(rawVba);
            package.ChangeDocumentType(PresentationDocumentType.MacroEnabledPresentation);
        }
        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddWordHandler().AddExcelHandler().AddPowerPointHandler().Build()
            .ReadDocument(bytes.ToArray(), name, new ReaderOptions {
                Projection = ReaderProjection.TextAndMetadata, Inspection = new ReaderInspectionOptions { IncludeContentSafety = false }
            });
        OfficeDocumentMetadataEntry vba = Assert.Single(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.VbaProject);
        Assert.Equal("false", vba.Attributes!["moduleInspectionRequested"]);
        Assert.Equal("false", vba.Attributes["lengthKnown"]);
        Assert.DoesNotContain("lengthBytes", vba.Attributes.Keys);
        Assert.Contains(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Coverage && item.Value == "module-payload-unrequested");
    }

    [Theory]
    [InlineData("word", "mislabeled.docx")]
    [InlineData("excel", "mislabeled.xlsx")]
    [InlineData("powerpoint", "mislabeled.pptx")]
    public void ModernLeanReaderRefusesNativeBinarySourceBeforeProjection(string format, string name) {
        byte[] source;
        if (format == "word") {
            using WordDocument document = WordDocument.Create(); document.AddParagraph("Legacy body");
            source = document.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc });
            using var input = new MemoryStream(source, writable: false);
            Assert.Throws<NotSupportedException>(() => WordDocument.Load(input, new WordLoadOptions { AllowLegacyProjection = false }));
            using WordDocument ordinary = WordDocument.Load(new MemoryStream(source, writable: false));
            Assert.Contains(ordinary.Paragraphs, paragraph => paragraph.Text == "Legacy body");
        } else if (format == "excel") {
            using ExcelDocument document = ExcelDocument.Create(); document.AddWorksheet("Legacy").CellValue(1, 1, "Legacy body");
            source = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls });
            using var input = new MemoryStream(source, writable: false);
            Assert.Throws<NotSupportedException>(() => ExcelDocument.Load(input, new ExcelLoadOptions { AllowBinaryProjection = false }));
            using ExcelDocument ordinary = ExcelDocument.Load(new MemoryStream(source, writable: false)); Assert.Single(ordinary.Sheets);
        } else {
            using PowerPointDocument document = PowerPointDocument.Create(); document.AddSlide().AddTextBox("Legacy body", 100000, 100000, 2000000, 500000);
            source = document.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt });
            using var input = new MemoryStream(source, writable: false);
            Assert.Throws<NotSupportedException>(() => PowerPointDocument.Load(input, new PowerPointLoadOptions { AllowLegacyProjection = false, Password = "unused-wrong-password" }));
            using PowerPointDocument ordinary = PowerPointDocument.Load(new MemoryStream(source, writable: false)); Assert.Single(ordinary.Slides);
        }
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().AddExcelHandler().AddPowerPointHandler().Build();
        Assert.Throws<NotSupportedException>(() => reader.ReadDocument(source, name, new ReaderOptions { Projection = ReaderProjection.TextAndMetadata }));
    }

    [Theory]
    [InlineData("word", "source.docx", true)]
    [InlineData("excel", "source.xlsx", true)]
    [InlineData("powerpoint", "source.pptx", true)]
    [InlineData("word", "source.docx", false)]
    [InlineData("excel", "source.xlsx", false)]
    [InlineData("powerpoint", "source.pptx", false)]
    public void NativeExpansionAdmissionIsTerminalWithoutInspectionOrPasswordRetry(string format, string name, bool perPart) {
        byte[] source = Create(format);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().AddExcelHandler().AddPowerPointHandler().Build();
        var options = new ReaderOptions {
            Projection = ReaderProjection.TextAndMetadata, OpenPassword = "must-not-be-used-for-an-admission-failure",
            OpenXmlMaxPartUncompressedBytes = perPart ? 1 : null,
            OpenXmlMaxTotalUncompressedBytes = perPart ? null : 1
        };
        ReaderResourceLimitException failure = Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(source, name, options));
        Assert.Equal(perPart ? "OpenXmlMaxPartUncompressedBytes" : "OpenXmlMaxTotalUncompressedBytes", failure.LimitName);
        Assert.Equal(1, failure.Maximum);
        OfficePackageSecurityException native = Assert.IsType<OfficePackageSecurityException>(failure.InnerException);
        Assert.Equal(perPart ? OfficePackageSecurityRule.PartSize : OfficePackageSecurityRule.TotalUncompressedSize, native.Rule);
        Assert.Null(options.Inspection);
        Assert.Equal("must-not-be-used-for-an-admission-failure", options.OpenPassword);
        Assert.Equal(ReaderProjection.TextAndMetadata, options.Projection);
    }

    [Theory]
    [InlineData("word", "encrypted.docx")]
    [InlineData("excel", "encrypted.xlsx")]
    [InlineData("powerpoint", "encrypted.pptx")]
    public void ModernEncryptedLeanReadsRetainTextAndOriginalContainerEvidence(string format, string name) {
        const string password = "lean-encrypted-public-proof";
        byte[] original = Create(format);
        using var input = new MemoryStream(original, writable: false);
        using var encrypted = new MemoryStream();
        if (format == "word") { using WordDocument document = WordDocument.Load(input); document.Save(encrypted,new OfficeIMO.Word.WordSaveOptions{Password=password}); }
        else if (format == "excel") { using ExcelDocument document = ExcelDocument.Load(input); document.Save(encrypted,new OfficeIMO.Excel.ExcelSaveOptions{Password=password}); }
        else { using PowerPointDocument document = PowerPointDocument.Load(input); document.Save(encrypted,new OfficeIMO.PowerPoint.PowerPointSaveOptions{Password=password}); }
        byte[] source = encrypted.ToArray();
        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddWordHandler().AddExcelHandler().AddPowerPointHandler().Build()
            .ReadDocument(source, name, new ReaderOptions {
                Projection = ReaderProjection.TextAndMetadata, OpenPassword = password,
                OpenXmlMaxPartUncompressedBytes = 1024 * 1024, OpenXmlMaxTotalUncompressedBytes = 8 * 1024 * 1024,
                Inspection = new ReaderInspectionOptions { IncludeContentSafety = false }
            });
        Assert.Contains(result.Chunks, chunk => chunk.Text!.Contains("Body evidence", StringComparison.Ordinal));
        Assert.Empty(result.Assets);
        Assert.Contains(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.PackageSecurity
            && item.Attributes != null && item.Attributes.TryGetValue("scope", out string? scope) && scope == "encrypted-source-container");
    }

    [Theory]
    [InlineData("path")]
    [InlineData("stream")]
    [InlineData("path-async")]
    [InlineData("stream-async")]
    public async Task NativeEncryptedBinaryPowerPointRefusalPrecedesPasswordAndPayloadProjection(string route) {
        using PowerPointDocument source = PowerPointDocument.Create(); source.AddSlide().AddTextBox("Legacy body", 100000, 100000, 2000000, 500000);
        byte[] bytes = source.ToBytes(new OfficeIMO.PowerPoint.PowerPointSaveOptions{Password="source-password",Format=PowerPointFileFormat.Ppt});
        string path = Path.Combine(Path.GetTempPath(), "lean-legacy-" + Guid.NewGuid().ToString("N") + ".ppt");
        File.WriteAllBytes(path, bytes);
        var options = new PowerPointLoadOptions { AllowLegacyProjection = false, Password = "unused-wrong-password" };
        using var input = new MemoryStream(bytes, writable: false);
        try {
            if (route == "path") Assert.Throws<NotSupportedException>(() => PowerPointDocument.Load(path, options));
            else if (route == "stream") Assert.Throws<NotSupportedException>(() => PowerPointDocument.Load(input, options));
            else if (route == "path-async") await Assert.ThrowsAsync<NotSupportedException>(() => PowerPointDocument.LoadAsync(path, options));
            else await Assert.ThrowsAsync<NotSupportedException>(() => PowerPointDocument.LoadAsync(input, options));
            Assert.False(options.AllowLegacyProjection);
            Assert.True(input.CanRead);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void XmlOnlyExcelLoadRefusesXlsbBeforeProjection() {
        using ExcelDocument source = ExcelDocument.Create(); source.AddWorksheet("Binary").CellValue(1, 1, "Saved value");
        byte[] bytes = source.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        using var input = new MemoryStream(bytes, writable: false);
        Assert.Throws<NotSupportedException>(() => ExcelDocument.Load(input, new ExcelLoadOptions { AllowBinaryProjection = false }));
        using ExcelDocument ordinary = ExcelDocument.Load(new MemoryStream(bytes, writable: false)); Assert.Single(ordinary.Sheets);
    }

    [Theory]
    [InlineData("word", ".docx")]
    [InlineData("excel", ".xlsx")]
    [InlineData("powerpoint", ".pptx")]
    public async Task NativeAdmissionRemainsTerminalInBatchAndFolderWhileMalformedInputIsPerFile(string format, string extension) {
        string directory = Path.Combine(Path.GetTempPath(), "lean-admission-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "source" + extension);
        File.WriteAllBytes(path, Create(format));
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().AddExcelHandler().AddPowerPointHandler().Build();
        var options = new ReaderOptions { Projection = ReaderProjection.TextAndMetadata, OpenXmlMaxPartUncompressedBytes = 1 };
        try {
            await Assert.ThrowsAsync<ReaderResourceLimitException>(() => reader.ReadDocumentsDetailedAsync(new[] { path }, options));
            Assert.Throws<ReaderResourceLimitException>(() => reader.ReadFolder(directory, options: options).ToArray());
            File.WriteAllBytes(path, new byte[] { 0, 1, 2, 3 });
            ReaderDocumentReadOutcome malformed = Assert.Single(await reader.ReadDocumentsDetailedAsync(new[] { path }, new ReaderOptions {
                Projection = ReaderProjection.TextAndMetadata, OpenXmlMaxPartUncompressedBytes = 1024 * 1024
            }));
            Assert.False(malformed.Succeeded);
            Assert.NotNull(malformed.Error);
            Assert.IsNotType<ReaderResourceLimitException>(malformed.Error);
        } finally {
            File.Delete(path); Directory.Delete(directory);
        }
    }

    private static byte[] Create(string format) {
        if (format == "word") {
            using WordDocument document = WordDocument.Create();
            using var image = new MemoryStream(Pixel, writable: false);
            document.AddParagraph("Body evidence").AddImage(image, "pixel.png", OfficeIMO.OfficeLength.Pixels(1), OfficeIMO.OfficeLength.Pixels(1));
            return document.ToBytes();
        }
        if (format == "excel") {
            using ExcelDocument document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Evidence"); sheet.CellValue(1, 1, "Header"); sheet.CellValue(2, 1, "Body evidence");
            sheet.AddImage(1, 1, Pixel, widthPixels: 1, heightPixels: 1);
            return document.ToBytes();
        }
        using PowerPointDocument presentation = PowerPointDocument.Create();
        PowerPointSlide slide = presentation.AddSlide(); slide.AddTextBox("Body evidence", 100000, 100000, 2000000, 500000);
        using var picture = new MemoryStream(Pixel, writable: false); slide.AddImage(picture, OfficeImageFormat.Png);
        return presentation.ToBytes();
    }
}
