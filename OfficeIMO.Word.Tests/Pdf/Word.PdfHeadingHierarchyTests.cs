using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "docx")]
    [InlineData(true, "docx")]
    [InlineData(false, "doc")]
    [InlineData(true, "doc")]
    public void SaveAsPdf_PreservesNineHeadingLevelsInBodyAndColumns(bool columns, string extension) {
        string source = Path.Combine(_directoryWithFiles, "HeadingHierarchy" + columns + "." + extension);
        string path = source + ".pdf";
        using var document = WordDocument.Create(source);
        if (columns) document.Sections[0].ColumnCount = 2;
        WordParagraphStyles[] styles = {
            WordParagraphStyles.Heading1, WordParagraphStyles.Heading2, WordParagraphStyles.Heading3,
            WordParagraphStyles.Heading4, WordParagraphStyles.Heading5, WordParagraphStyles.Heading6,
            WordParagraphStyles.Heading7, WordParagraphStyles.Heading8, WordParagraphStyles.Heading9
        };
        for (int index = 0; index < styles.Length; index++) {
            WordParagraph heading = document.AddParagraph("AuthoredHeading" + (index + 1));
            heading.SetStyle(styles[index]);
            heading.FontSize = 12;
        }
        document.Save();
        using var reopened = WordDocument.Load(source);
        reopened.SaveAsPdf(path, new WordToPdfOptions { IncludePageNumbers = false });
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(File.ReadAllBytes(path));
        for (int level = 1; level <= 9; level++) {
            var heading = Assert.Single(logical.Headings, h => h.Text == "AuthoredHeading" + level);
            Assert.Equal(level, heading.Level);
        }
        using WordDocument imported = PdfDocument.Load(path).ToWordDocumentResult(new PdfToWordOptions()).Value;
        for (int level = 1; level <= 9; level++) {
            var heading = Assert.Single(imported.Paragraphs, p => p.Text == "AuthoredHeading" + level);
            Assert.Equal(styles[level - 1], heading.Style);
        }
    }
}
