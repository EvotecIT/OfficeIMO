using System;
using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "none", 3)]
    [InlineData(true, "none", 3)]
    [InlineData(false, "unconfigured", 3)]
    [InlineData(true, "unconfigured", 3)]
    [InlineData(false, "collapse", 1)]
    [InlineData(true, "collapse", 1)]
    [InlineData(false, "preformatted", 3)]
    [InlineData(true, "preformatted", 3)]
    public void SaveAsPdf_PreservesWordSpacesByDefaultAndHonorsExplicitPdfPolicy(bool table, string policy, int separatorCount) {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph;
        if (table) paragraph = word.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        else paragraph = word.AddParagraph();
        paragraph.Text = "  ALPHA   BETA"; paragraph.FontSize = 12; paragraph.FontFamily = "Arial";
        PdfOptions? pdfOptions = policy == "none" ? null : new PdfOptions { DefaultFontSize = 12D };
        if (policy == "collapse") pdfOptions!.PreserveTextWhitespace = false;
        if (policy == "preformatted") pdfOptions!.PreserveTextWhitespace = true;
        bool? callerValue = pdfOptions?.PreserveTextWhitespace;
        string source = word._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = PdfPigDocument.Open(word.ToPdfDocument(new WordToPdfOptions {
            IncludePageNumbers = false, FontFamily = "Arial", ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = pdfOptions
        }).ToBytes());
        var letters = pdf.GetPage(1).Letters.Where(l => !string.IsNullOrWhiteSpace(l.Value)).ToArray();
        Assert.Equal("ALPHABETA", string.Concat(letters.Select(l => l.Value)));
        Assert.InRange(Math.Abs(letters[5].StartBaseLine.X - letters[0].StartBaseLine.X - 39.348D - separatorCount * 3.336D), 0, .02D);
        Assert.Equal(callerValue, pdfOptions?.PreserveTextWhitespace);
        Assert.Equal(source, word._wordprocessingDocument.MainDocumentPart!.Document.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_DefaultSpaceFlowWrapsOnceAndCloneRetainsExplicitOverride() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("ALPHA" + new string(' ', 60) + "BETA");
        paragraph.FontSize = 12; paragraph.FontFamily = "Arial";
        PdfOptions unconfigured = new PdfOptions { DefaultFontSize = 12D }.Clone();
        var options = new WordToPdfOptions { IncludePageNumbers = false, FontFamily = "Arial",
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = unconfigured,
            PageSize = new PageSize(240D, 792D), Margins = PageMargins.Uniform(72D) };
        using var normal = PdfPigDocument.Open(word.ToPdfDocument(options).ToBytes());
        var letters = normal.GetPage(1).Letters.Where(l => !string.IsNullOrWhiteSpace(l.Value)).ToArray();
        Assert.InRange(letters[0].StartBaseLine.Y - letters[5].StartBaseLine.Y, 12D, 16D);
        Assert.InRange(Math.Abs(letters[5].StartBaseLine.X - 72D), 0, .02D);
        PdfOptions explicitOptions = new PdfOptions { DefaultFontSize = 12D, PreserveTextWhitespace = false }.Clone();
        options.PdfOptions = explicitOptions;
        using var collapsed = PdfPigDocument.Open(word.ToPdfDocument(options).ToBytes());
        var collapsedLetters = collapsed.GetPage(1).Letters.Where(l => !string.IsNullOrWhiteSpace(l.Value)).ToArray();
        Assert.Equal(collapsedLetters[0].StartBaseLine.Y, collapsedLetters[5].StartBaseLine.Y, 3);
        Assert.False(explicitOptions.PreserveTextWhitespace);
    }
}
