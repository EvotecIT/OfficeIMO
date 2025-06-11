using System.IO;
using OfficeIMO.Word;
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

            Assert.Equal(1, document.EmbeddedObjects.Count);
            Assert.Equal(1, document.Sections[0].EmbeddedObjects.Count);
            document.Save();
        }

        using (var document = WordDocument.Load(filePath)) {
            Assert.Equal(1, document.EmbeddedObjects.Count);
            Assert.Equal(1, document.Sections[0].EmbeddedObjects.Count);
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

            Assert.Equal(1, document.EmbeddedObjects.Count);
            document.Save();
        }

        using (var document = WordDocument.Load(filePath)) {
            Assert.Equal(1, document.EmbeddedObjects.Count);
        }
    }
}

