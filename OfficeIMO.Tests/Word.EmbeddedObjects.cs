using System.IO;
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
        var shape = word.MainDocumentPart.Document.Descendants<V.Shape>().FirstOrDefault();
        Assert.NotNull(shape);
        Assert.Contains("width:32pt", shape.Style);
        Assert.Contains("height:32pt", shape.Style);
    }
}
