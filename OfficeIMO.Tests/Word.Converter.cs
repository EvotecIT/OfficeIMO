using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_ConvertDotXtoDocX() {
        string filePath = Path.Combine(_directoryDocuments, "ExampleTemplate.dotx");
        string outFilePath = Path.Combine(_directoryWithFiles, "ExampleTemplate.docx");
        WordHelpers.ConvertDotXtoDocX(filePath, outFilePath);

        Assert.True(File.Exists(outFilePath));

        using (WordDocument document = WordDocument.Load(filePath)) {
            Assert.True(document.Paragraphs.Count == 57);
            Assert.Contains(document.Paragraphs[56].Text, "Feel free to use and share the file according to the license above.");
        }
    }

    [Fact]
    public void Test_ConvertDotXtoDocX_ReleasesDocumentStream() {
        string templatePath = Path.Combine(_directoryDocuments, "ExampleTemplate.dotx");
        string outFilePath = Path.Combine(_directoryWithFiles, "ExampleTemplate_Cleanup.docx");

        WordHelpers.ConvertDotXtoDocX(templatePath, outFilePath);

        Assert.False(templatePath.IsFileLocked());
        Assert.True(File.Exists(outFilePath));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Test_SaveDotXAsDocX_ChangesPackageType(bool useSaveAs) {
        string templatePath = Path.Combine(_directoryDocuments, "ExampleTemplate.dotx");
        string outFilePath = Path.Combine(_directoryWithFiles, useSaveAs ? "ExampleTemplate_SaveAs.docx" : "ExampleTemplate_Save.docx");

        using (WordDocument document = WordDocument.Load(templatePath)) {
            if (useSaveAs) {
                using WordDocument savedDocument = document.SaveAs(outFilePath);
            } else {
                document.Save(outFilePath);
            }
        }

        Assert.True(File.Exists(outFilePath));
        using WordprocessingDocument saved = WordprocessingDocument.Open(outFilePath, false);
        Assert.Equal(WordprocessingDocumentType.Document, saved.DocumentType);
        Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", saved.MainDocumentPart?.ContentType);
    }
}
