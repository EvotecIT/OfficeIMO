using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_ConvertDotxToDocx() {
        string filePath = Path.Combine(_directoryDocuments, "ExampleTemplate.dotx");
        string outFilePath = Path.Combine(_directoryWithFiles, "ExampleTemplate.docx");
        WordHelpers.ConvertDotxToDocx(filePath, outFilePath);

        Assert.True(File.Exists(outFilePath));

        using (WordDocument document = WordDocument.Load(filePath)) {
            Assert.True(document.Paragraphs.Count == 57);
            Assert.Contains(document.Paragraphs[56].Text, "Feel free to use and share the file according to the license above.");
        }
    }

    [Fact]
    public void Test_ConvertDotxToDocx_ReleasesDocumentStream() {
        string templatePath = Path.Combine(_directoryDocuments, "ExampleTemplate.dotx");
        string outFilePath = Path.Combine(_directoryWithFiles, "ExampleTemplate_Cleanup.docx");

        WordHelpers.ConvertDotxToDocx(templatePath, outFilePath);

        Assert.False(templatePath.IsFileLocked());
        Assert.True(File.Exists(outFilePath));
    }

    [Fact]
    public void Test_ConvertDotxToDocx_AcceptsRelativeTemplatePath() {
        string templatePath = Path.Combine(_directoryDocuments, "ExampleTemplate.dotx");
        string relativeTemplatePath = GetConverterRelativePath(Environment.CurrentDirectory, templatePath);
        string outFilePath = Path.Combine(_directoryWithFiles, "ExampleTemplate_Relative.docx");

        Assert.False(Path.IsPathRooted(relativeTemplatePath));
        WordHelpers.ConvertDotxToDocx(relativeTemplatePath, outFilePath);

        Assert.True(File.Exists(outFilePath));
        using WordprocessingDocument saved = WordprocessingDocument.Open(outFilePath, false);
        Assert.Equal(WordprocessingDocumentType.Document, saved.DocumentType);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Test_SaveDotXAsDocX_ChangesPackageType(bool useSaveAs) {
        const string templateFileName = "ExampleTemplate.dotx";
        string outputFileName = useSaveAs ? "ExampleTemplate_SaveAs.docx" : "ExampleTemplate_Save.docx";
        Assert.False(Path.IsPathRooted(templateFileName));
        Assert.False(Path.IsPathRooted(outputFileName));

        string templatePath = JoinPath(_directoryDocuments, templateFileName);
        string outFilePath = JoinPath(_directoryWithFiles, outputFileName);

        using (WordDocument document = WordDocument.Load(templatePath)) {
            if (useSaveAs) {
                document.SaveCopy(outFilePath);
            } else {
                document.Save(outFilePath);
            }
        }

        Assert.True(File.Exists(outFilePath));
        using WordprocessingDocument saved = WordprocessingDocument.Open(outFilePath, false);
        Assert.Equal(WordprocessingDocumentType.Document, saved.DocumentType);
        Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", saved.MainDocumentPart?.ContentType);
    }

    private static string JoinPath(string basePath, string fileName) {
#if NET472
        char separator = Path.DirectorySeparatorChar;
        return basePath.EndsWith(separator.ToString(), StringComparison.Ordinal) ||
               basePath.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            ? basePath + fileName
            : basePath + separator + fileName;
#else
        return Path.Join(basePath, fileName);
#endif
    }

    private static string GetConverterRelativePath(string baseDirectory, string path) {
#if NET472
        string fullBaseDirectory = Path.GetFullPath(baseDirectory);
        if (!fullBaseDirectory.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)) {
            fullBaseDirectory += Path.DirectorySeparatorChar;
        }

        var baseUri = new Uri(fullBaseDirectory);
        var pathUri = new Uri(Path.GetFullPath(path));
        return Uri.UnescapeDataString(baseUri.MakeRelativeUri(pathUri).ToString())
            .Replace('/', Path.DirectorySeparatorChar);
#else
        return Path.GetRelativePath(baseDirectory, path);
#endif
    }
}
