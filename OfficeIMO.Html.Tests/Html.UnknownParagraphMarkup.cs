using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Word markup inside a wrapper the package doesn't know (a vendor element) is held untyped and isn't part of the
/// document model: HTML export reads every story as if the wrapper weren't there.
/// </summary>
public sealed class HtmlUnknownParagraphMarkupTests {
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string VendorNamespace = "urn:officeimo:tests:vendor";

    // A paragraph, a footnote reference, a bookmark, a field and deleted text, all inside the wrapper.
    private const string WrappedMarkup =
        "<w:p><w:r><w:t>Wrapped</w:t></w:r></w:p>" +
        "<w:r><w:footnoteReference w:id=\"1\"/></w:r>" +
        "<w:bookmarkStart w:id=\"90\" w:name=\"WrappedMark\"/>" +
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r>" +
        "<w:r><w:delText>Gone</w:delText></w:r>";

    [Theory]
    [InlineData("body")]
    [InlineData("run")]
    [InlineData("heading")]
    [InlineData("code")]
    [InlineData("table cell")]
    [InlineData("header")]
    [InlineData("footnote")]
    [InlineData("text box")]
    public void MarkupInsideAnUnknownWrapperDoesNotChangeTheExport(string story) {
        string expected = ExportHtml(wrappedStory: null);

        string actual = ExportHtml(story);

        Assert.Equal(expected, actual);
        Assert.DoesNotContain("Wrapped", actual, StringComparison.Ordinal);
    }

    private static string ExportHtml(string? wrappedStory) {
        using var document = WordDocument.Create();
        WordParagraph header = document.PageSetup.Header.AddParagraph("Header text");
        WordParagraph heading = document.AddParagraph("Heading text").SetStyle(WordParagraphStyles.Heading1);
        WordParagraph code = document.AddParagraph("Code text");
        code.SetStyleId("HTMLPreformatted");
        WordParagraph body = document.AddParagraph("Body text");
        WordFootNote note = body.AddFootNote("Note text");
        WordTable table = document.AddTable(1, 1);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Cell text";
        WordTextBox box = document.AddTextBox("Box text");

        OpenXmlElement? target = wrappedStory switch {
            "body" => body._paragraph,
            "run" => body._paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().First(),
            "heading" => heading._paragraph,
            "code" => code._paragraph,
            "table cell" => table.Rows[0].Cells[0]._tableCell,
            "header" => header._paragraph,
            "footnote" => note.Paragraphs!.Last()._paragraph,
            "text box" => box.Paragraphs.Last()._paragraph,
            _ => null
        };
        target?.AppendChild(Wrapper());

        return document.ToHtml(new WordHtmlExportOptions {
            ExportHeadersAndFooters = true
        });
    }

    // Built as unknown markup with inner XML, the way the package holds it when loaded: the Word elements inside stay untyped.
    private static OpenXmlUnknownElement Wrapper() {
        var wrapper = new OpenXmlUnknownElement("x", "wrapper", VendorNamespace);
        wrapper.InnerXml = $"<x:group xmlns:x=\"{VendorNamespace}\" xmlns:w=\"{WordNamespace}\">{WrappedMarkup}</x:group>";
        return wrapper;
    }
}
