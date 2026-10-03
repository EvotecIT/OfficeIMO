using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_PaginationFragmentsRetainMultipleColumnBreaks() {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].ColumnCount = 3;
        document.AddParagraph("FirstColumn").AddBreak(WordBreakType.Column).AddText("SecondColumn")
            .AddBreak(WordBreakType.Column).AddText("ThirdColumn");
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(1, pdf.NumberOfPages);
        var words = pdf.GetPage(1).GetWords().ToDictionary(word => word.Text);
        Assert.True(words["FirstColumn"].BoundingBox.Left < words["SecondColumn"].BoundingBox.Left);
        Assert.True(words["SecondColumn"].BoundingBox.Left < words["ThirdColumn"].BoundingBox.Left);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_PaginationFragmentsRetainExplicitBookmark(bool columnBreak, bool bookmarkOnlyPrefix) {
        using WordDocument document = WordDocument.Create();
        if (columnBreak) document.Sections[0].ColumnCount = 2;
        WordParagraph target = document.AddParagraph(bookmarkOnlyPrefix ? "" : "Before");
        target.AddBookmark("FragmentTarget");
        target.AddBreak(columnBreak ? WordBreakType.Column : WordBreakType.Page).AddText("After");
        document.AddHyperLink("Go to target", "FragmentTarget");
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false });
        var logical = OfficeIMO.Pdf.PdfDocumentReadResult.Load(bytes);
        Assert.Contains(logical.NamedDestinations, destination => destination.Name == "FragmentTarget");
        Assert.Contains(logical.GetLinksByDestinationName("FragmentTarget"), link => link.IsNamedDestinationLink);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("After", string.Join(" ", pdf.GetPages().Select(page => page.Text)));
    }

    [Theory]
    [InlineData(false, 2)]
    [InlineData(true, 1)]
    public void SaveAsPdf_PaginationFragmentsRetainCrossParagraphFieldVisibility(bool columnBreak, int expectedPages) {
        using WordDocument document = WordDocument.Create();
        if (columnBreak) document.Sections[0].ColumnCount = 2;
        document.AddParagraph()._paragraph.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }, new W.FieldCode(" PRIVATE ")));
        W.BreakValues breakType = columnBreak ? W.BreakValues.Column : W.BreakValues.Page;
        document.AddParagraph()._paragraph.Append(
            new W.Run(new W.Text("HiddenInstructionPlain"), new W.Break { Type = breakType }),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
            new W.Run(new W.Text("Before"), new W.Break { Type = breakType }, new W.Text("After")),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        string text = string.Join(" ", pdf.GetPages().Select(page => page.Text));
        Assert.Equal(expectedPages, pdf.NumberOfPages);
        Assert.Contains("Before", text);
        Assert.Contains("After", text);
        Assert.DoesNotContain("HiddenInstructionPlain", text);
        Assert.DoesNotContain("PRIVATE", text);
    }
}
