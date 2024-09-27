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
}
