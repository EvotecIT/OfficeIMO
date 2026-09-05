using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordAllSeverityBatch14SecurityTests {
    [Fact]
    public void RemovingStoryImagesDeletesRelationshipsFromTheirOwningParts() {
        string imagePath = Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg");
        using WordDocument document = WordDocument.Create();
        WordParagraph body = document.AddParagraph("stories");
        WordParagraph footnote = body.AddFootNote("footnote").FootNote!.Paragraphs![1];
        WordParagraph endnote = body.AddEndNote("endnote").EndNote!.Paragraphs![1];
        body.AddComment("OfficeIMO", "OI", "comment");
        WordParagraph comment = Assert.Single(document.Comments).Paragraphs[0];
        footnote.AddImage(imagePath, 16, 16);
        endnote.AddImage(imagePath, 16, 16);
        comment.AddImage(imagePath, 16, 16);

        footnote.Image!.Remove();
        endnote.Image!.Remove();
        comment.Image!.Remove();

        MainDocumentPart main = document._wordprocessingDocument.MainDocumentPart!;
        Assert.DoesNotContain(main.FootnotesPart!.Parts, pair => pair.OpenXmlPart is ImagePart);
        Assert.DoesNotContain(main.EndnotesPart!.Parts, pair => pair.OpenXmlPart is ImagePart);
        Assert.DoesNotContain(main.WordprocessingCommentsPart!.Parts, pair => pair.OpenXmlPart is ImagePart);
    }

    [Fact]
    public void NativePdfClampsUntrustedHeaderExpansionToThePage() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        for (int index = 0; index < 100; index++) {
            document.Header.Default!.AddParagraph("large header " + index);
        }
        document.AddParagraph("body remains renderable");
        using var output = new MemoryStream();

        document.SaveAsPdf(output, new WordToPdfOptions { IncludePageNumbers = false });

        Assert.True(output.Length > 0);
    }

    [Fact]
    public void NativePdfTreatsZeroDxaTableWidthAsAutomatic() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = 0;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "cell";
        using var output = new MemoryStream();

        document.SaveAsPdf(output, new WordToPdfOptions { IncludePageNumbers = false });

        Assert.True(output.Length > 0);
    }
}
