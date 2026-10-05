using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "docx", true)]
    [InlineData(true, "docx", true)]
    [InlineData(false, "doc", true)]
    [InlineData(true, "doc", true)]
    [InlineData(false, "docx", false)]
    [InlineData(true, "docx", false)]
    [InlineData(false, "doc", false)]
    [InlineData(true, "doc", false)]
    public void SaveAsPdf_SparseHeadingLevelsUseTagsAndBookmarkFallback(bool columns, string extension, bool tagged) {
        string source = Path.Combine(_directoryWithFiles, $"SparseHeadingHierarchy{columns}{tagged}.{extension}");
        using var document = WordDocument.Create(source);
        if (columns) document.Sections[0].ColumnCount = 2;
        WordParagraphStyles[] styles = { WordParagraphStyles.Heading1, WordParagraphStyles.Heading7, WordParagraphStyles.Heading9 };
        int[] levels = { 1, 7, 9 };
        for (int index = 0; index < styles.Length; index++) {
            WordParagraph heading = document.AddParagraph("SparseHeading" + levels[index]).SetStyle(styles[index]);
            heading.FontSize = 12;
            heading.AddText(" InlineItalic" + levels[index]).Italic = true;
            heading.AddFootNote("SparseNote" + levels[index]);
        }
        document.Save();
        using var reopened = WordDocument.Load(source);
        string output = source + ".pdf";
        reopened.SaveAsPdf(output, new WordToPdfOptions {
            IncludePageNumbers = false,
            PdfOptions = new PdfOptions { TaggedStructureMode = tagged ? PdfTaggedStructureMode.CatalogMarkers : PdfTaggedStructureMode.None }
        });
        var logical = PdfDocumentReadResult.Load(File.ReadAllBytes(output));
        using WordDocument imported = PdfDocument.Load(output).ToWordDocumentResult(new PdfToWordOptions()).Value;
        WordParagraphStyles[] fallback = { WordParagraphStyles.Heading1, WordParagraphStyles.Heading2, WordParagraphStyles.Heading3 };
        for (int index = 0; index < levels.Length; index++) {
            string marker = "SparseHeading" + levels[index];
            Assert.Equal(tagged ? levels[index] : index + 1,
                Assert.Single(logical.Headings, heading => heading.Text.Contains(marker)).Level);
            Assert.Equal(tagged ? styles[index] : fallback[index],
                Assert.Single(imported.Paragraphs, paragraph => paragraph.Text.Contains(marker)).Style);
        }
    }

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
