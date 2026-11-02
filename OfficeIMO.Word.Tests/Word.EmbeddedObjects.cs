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

        using var package = WordprocessingDocument.Open(filePath, false);
        var shape = Assert.Single(package.MainDocumentPart!.Document.Descendants<V.Shape>());
        Assert.Equal("width:64.8pt;height:64.8pt", shape.Style!.Value);
    }

    [Fact]
    public void Test_AddEmbeddedObjectWithOptions() {
        var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithEmbeddedOptions.docx");
        string excelFilePath = Path.Combine(_directoryDocuments, "SampleFileExcel.xlsx");
        string iconPath = Path.Combine(_directoryDocuments, "SampleExcelIcon.png");

        using (var document = WordDocument.Create(filePath)) {
            document.AddParagraph("Add excel object");
            var options = WordEmbeddedObjectOptions.Icon(iconPath, width: global::OfficeIMO.OfficeLength.Points(32), height: global::OfficeIMO.OfficeLength.Points(32));
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
        var currentCulture = System.Globalization.CultureInfo.CurrentCulture;
        try {
            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("fr-FR");
            using (var document = WordDocument.Create(filePath)) {
                document.AddParagraph("Add excel object");
                var options = WordEmbeddedObjectOptions.Icon(iconPath,
                    width: global::OfficeIMO.OfficeLength.Centimeters(2.54), height: global::OfficeIMO.OfficeLength.Pixels(48));
                document.AddEmbeddedObject(excelFilePath, options);
                document.AddParagraph().AddEmbeddedObject(excelFilePath, iconPath,
                    width: global::OfficeIMO.OfficeLength.Millimeters(12.7));
                document.Save();
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.EmbeddedObjects.Count);
                document.Save();
            }

            using var package = WordprocessingDocument.Open(filePath, false);
            var shapes = package.MainDocumentPart!.Document.Descendants<V.Shape>().ToArray();
            Assert.Equal(2, shapes.Length);
            Assert.Equal("width:72pt;height:36pt", shapes[0].Style!.Value);
            Assert.Equal("width:36pt;height:64.8pt", shapes[1].Style!.Value);
            Assert.Equal(2, package.MainDocumentPart.EmbeddedPackageParts.Count());
            foreach (var embedded in package.MainDocumentPart.EmbeddedPackageParts) {
                using var stream = embedded.GetStream(FileMode.Open, FileAccess.Read);
                using var spreadsheet = SpreadsheetDocument.Open(stream, false);
                Assert.NotEmpty(spreadsheet.WorkbookPart!.WorksheetParts);
            }
        } finally {
            System.Globalization.CultureInfo.CurrentCulture = currentCulture;
        }
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
            Assert.Equal(OfficeFeatureSupportLevel.PartiallyEditable, Assert.Single(document.InspectFeatures().FindFeatures("Embedded packages")).SupportLevel);
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
