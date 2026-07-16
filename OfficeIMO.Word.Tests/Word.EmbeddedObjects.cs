using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Packaging;
using V = DocumentFormat.OpenXml.Vml;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_CreatingWordDocumentWithEmbeddedObjects() {
        var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithEmbeddedObjects.docx");
        string excelFilePath = Path.Combine(_directoryDocuments, "SampleFileExcel.xlsx");
        string imageFilePath = Path.Combine(_directoryDocuments, "SampleExcelIcon.png");

        using (var document = WordDocument.Create(filePath)) {
            document.AddParagraph("Add excel object");
            document.AddEmbeddedObject(excelFilePath, imageFilePath);

            Assert.Single(document.EmbeddedObjects);
            Assert.Single(document.Sections[0].EmbeddedObjects);
            document.Save();
        }

        using (var document = WordDocument.Load(filePath)) {
            Assert.Single(document.EmbeddedObjects);
            Assert.Single(document.Sections[0].EmbeddedObjects);
        }
    }

    [Fact]
    public void Test_AddEmbeddedObjectWithOptions() {
        var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithEmbeddedOptions.docx");
        string excelFilePath = Path.Combine(_directoryDocuments, "SampleFileExcel.xlsx");
        string iconPath = Path.Combine(_directoryDocuments, "SampleExcelIcon.png");

        using (var document = WordDocument.Create(filePath)) {
            document.AddParagraph("Add excel object");
            var options = WordEmbeddedObjectOptions.Icon(iconPath, width: 32, height: 32);
            document.AddEmbeddedObject(excelFilePath, options);

            Assert.Single(document.EmbeddedObjects);
            document.Save();
        }

        using (var document = WordDocument.Load(filePath)) {
            Assert.Single(document.EmbeddedObjects);
        }
    }

    [Fact]
    public void Test_EmbeddedObjectCustomSize() {
        var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithEmbeddedCustomSize.docx");
        string excelFilePath = Path.Combine(_directoryDocuments, "SampleFileExcel.xlsx");
        string iconPath = Path.Combine(_directoryDocuments, "SampleExcelIcon.png");

        using (var document = WordDocument.Create(filePath)) {
            document.AddParagraph("Add excel object");
            var options = WordEmbeddedObjectOptions.Icon(iconPath, width: 32, height: 32);
            document.AddEmbeddedObject(excelFilePath, options);
            document.Save();
        }

        using var word = WordprocessingDocument.Open(filePath, false);
        var shape = word.MainDocumentPart?.Document?.Descendants<V.Shape>().FirstOrDefault();
        Assert.NotNull(shape);
        Assert.Contains("width:32pt", shape!.Style);
        Assert.Contains("height:32pt", shape!.Style);
    }

    [Fact]
    public void Test_LoadAndOpenEmbeddedExcelObject() {
        var filePath = Path.Combine(_directoryWithFiles, "DocumentWithEmbeddedExcel.docx");
        string excelFilePath = Path.Combine(_directoryDocuments, "SampleFileExcel.xlsx");
        string imageFilePath = Path.Combine(_directoryDocuments, "SampleExcelIcon.png");

        using (var document = WordDocument.Create(filePath)) {
            document.AddParagraph("Add excel object");
            document.AddEmbeddedObject(excelFilePath, imageFilePath);
            document.Save();
        }

        using (var word = WordprocessingDocument.Open(filePath, false)) {
            var embeddedPart = word.MainDocumentPart?.EmbeddedPackageParts.FirstOrDefault();
            Assert.NotNull(embeddedPart);
            using var stream = embeddedPart!.GetStream(FileMode.Open, FileAccess.Read);
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            ms.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(ms, false);
            Assert.NotNull(spreadsheet.WorkbookPart);
            Assert.True(spreadsheet.WorkbookPart.WorksheetParts.Any());
        }
    }

    [Fact]
    public void Test_EmbeddedPayloadCanBeHashedReplacedAndRemoved() {
        string filePath = Path.Combine(_directoryWithFiles, "EmbeddedPayloadManagement.docx");
        string excelFilePath = Path.Combine(_directoryDocuments, "SampleFileExcel.xlsx");
        string imageFilePath = Path.Combine(_directoryDocuments, "SampleExcelIcon.png");
        byte[] original = File.ReadAllBytes(excelFilePath);
        byte[] replacement = System.Text.Encoding.UTF8.GetBytes("replacement package payload");
        File.Delete(filePath);

        string payloadId;
        using (WordDocument document = WordDocument.Create(filePath)) {
            document.AddEmbeddedObject(excelFilePath, imageFilePath);
            OfficeEmbeddedPayloadInfo info = Assert.Single(document.GetEmbeddedPayloads(includeSha256: true));
            payloadId = info.Id;
            Assert.Equal(OfficeEmbeddedPayloadKind.EmbeddedPackage, info.Kind);
            Assert.Equal(original.Length, info.Length);
            Assert.Equal(64, info.Sha256!.Length);
            Assert.Equal(original, document.ExtractEmbeddedPayload(info.Id));
            Assert.Throws<InvalidDataException>(() => document.ExtractEmbeddedPayload(info.Id, original.Length - 1));
            document.ReplaceEmbeddedPayload(info.Id, replacement);
            document.Save();
        }

        using (WordDocument document = WordDocument.Load(filePath)) {
            Assert.Equal(replacement, document.ExtractEmbeddedPayload(payloadId));
            Assert.Equal(WordFeatureSupportLevel.PartiallyEditable, Assert.Single(document.InspectFeatures().FindFeatures("Embedded packages")).SupportLevel);
            Assert.True(document.RemoveEmbeddedPayload(payloadId));
            Assert.False(document.RemoveEmbeddedPayload(payloadId));
            Assert.Empty(document.GetEmbeddedPayloads());
            Assert.Empty(document.EmbeddedObjects);
            document.Save();
        }

        using WordprocessingDocument saved = WordprocessingDocument.Open(filePath, false);
        Assert.Empty(saved.MainDocumentPart!.EmbeddedPackageParts);
        File.Delete(filePath);
    }
}
