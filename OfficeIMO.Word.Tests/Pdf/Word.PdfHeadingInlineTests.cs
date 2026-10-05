using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordParagraphStyles.Heading1, false, "docx")]
    [InlineData(WordParagraphStyles.Heading7, false, "docx")]
    [InlineData(WordParagraphStyles.Heading9, true, "docx")]
    [InlineData(WordParagraphStyles.Heading7, false, "doc")]
    [InlineData(WordParagraphStyles.Heading9, true, "doc")]
    public void SaveAsPdf_HeadingPreservesInlineStylesAndLinks(WordParagraphStyles style, bool columns, string extension) {
        string source = Path.Combine(_directoryWithFiles, $"InlineHeading{style}{columns}.{extension}");
        const string uri = "https://example.com/heading-reference";
        using var document = WordDocument.Create(source);
        if (columns) document.Sections[0].ColumnCount = 2;
        WordParagraph heading = document.AddParagraph("PlainMarker ").SetStyle(style);
        heading.FontSize = 12;
        heading.AddText("ItalicMarker ").Italic = true;
        heading.AddHyperLink("LinkedMarker", new Uri(uri), tooltip: "Heading link metadata");
        document.Save();
        using var reopened = WordDocument.Load(source);
        WordHyperLink sourceLink = Assert.Single(reopened.Paragraphs[0].GetRuns(), run => run.IsHyperLink).Hyperlink!;
        string expectedContents = string.IsNullOrWhiteSpace(sourceLink.Tooltip) ? sourceLink.Text : sourceLink.Tooltip;
        string output = source + ".pdf";
        reopened.SaveAsPdf(output, new WordToPdfOptions { IncludePageNumbers = false });
        var logical = PdfDocumentReadResult.Load(File.ReadAllBytes(output));
        PdfLogicalLinkAnnotation link = Assert.Single(logical.GetLinksByUri(uri));
        Assert.Equal(expectedContents, link.Contents);
        var spans = PdfReadDocument.Open(File.ReadAllBytes(output)).Pages.SelectMany(page => page.GetTextSpans());
        Assert.Contains(spans, span => span.Text.Contains("ItalicMarker") && span.IsItalic);
        Assert.Contains(logical.Headings, item => item.Text.Contains("ItalicMarker"));
    }

    [Theory]
    [InlineData(WordParagraphStyles.Heading1, false, "docx")]
    [InlineData(WordParagraphStyles.Heading7, false, "docx")]
    [InlineData(WordParagraphStyles.Heading9, true, "docx")]
    [InlineData(WordParagraphStyles.Heading7, false, "doc")]
    [InlineData(WordParagraphStyles.Heading9, true, "doc")]
    public void SaveAsPdf_HeadingPreservesNoteReferences(WordParagraphStyles style, bool columns, string extension) {
        string source = Path.Combine(_directoryWithFiles, $"NoteHeading{style}{columns}.{extension}");
        using var document = WordDocument.Create(source);
        if (columns) document.Sections[0].ColumnCount = 2;
        WordParagraph heading = document.AddParagraph("HeadingNoteMarker").SetStyle(style);
        heading.FontSize = 12;
        heading.AddFootNote("FootnoteBodyMarker");
        heading.AddEndNote("EndnoteBodyMarker");
        document.Save();
        using var reopened = WordDocument.Load(source);
        string output = source + ".pdf";
        reopened.SaveAsPdf(output, new WordToPdfOptions { IncludePageNumbers = false });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(output);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("HeadingNoteMarker11", text);
        Assert.Contains("FootnoteBodyMarker", text);
        Assert.Contains("EndnoteBodyMarker", text);
    }
}
