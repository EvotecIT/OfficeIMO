using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Word markup inside a wrapper the package doesn't know (a vendor element) is held untyped and isn't part of the
/// document model: Markdown export and its diagnostics read every story as if the wrapper weren't there.
/// </summary>
public sealed class WordMarkdownUnknownParagraphMarkupTests {
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string VendorNamespace = "urn:officeimo:tests:vendor";

    // A paragraph, a bookmark, a page break, a VML shape and a field, all inside the wrapper.
    private const string WrappedMarkup =
        "<w:p><w:r><w:t>Wrapped</w:t></w:r></w:p>" +
        "<w:bookmarkStart w:id=\"90\" w:name=\"WrappedMark\"/>" +
        "<w:r><w:br w:type=\"page\"/></w:r>" +
        "<w:r><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\"/></w:pict></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>";

    [Theory]
    [InlineData("body")]
    [InlineData("heading")]
    [InlineData("run")]
    [InlineData("header")]
    [InlineData("text box")]
    public void MarkupInsideAnUnknownWrapperDoesNotChangeTheExport(string story) {
        (string expectedMarkdown, string expectedDiagnostics) = Export(wrappedStory: null);

        (string actualMarkdown, string actualDiagnostics) = Export(story);

        Assert.Equal(expectedMarkdown, actualMarkdown);
        Assert.Equal(expectedDiagnostics, actualDiagnostics);
        Assert.DoesNotContain("Wrapped", actualMarkdown, StringComparison.Ordinal);
    }

    private static (string Markdown, string Diagnostics) Export(string? wrappedStory) {
        using var document = WordDocument.Create();
        // The header holds no visible content of its own, so the omitted-header check reads only the wrapper.
        WordParagraph header = document.PageSetup.Header.AddParagraph();
        WordParagraph heading = document.AddParagraph("Heading text").SetStyle(WordParagraphStyles.Heading1);
        WordParagraph body = document.AddParagraph("Body text");
        body.AddFootNote("Note text");
        WordParagraph rule = document.AddParagraph();
        rule.Borders.BottomStyle = WordBorderStyle.Single;
        WordTable table = document.AddTable(1, 1);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Cell text";
        WordTextBox box = document.AddTextBox("Box text");

        OpenXmlElement? target = wrappedStory switch {
            "body" => body._paragraph,
            "heading" => heading._paragraph,
            "run" => body._paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().First(),
            "header" => header._paragraph,
            "text box" => box.Paragraphs.Last()._paragraph,
            _ => null
        };
        target?.AppendChild(Wrapper());

        var diagnostics = new List<string>();
        var options = new WordMarkdownExportOptions {
            DiagnosticSink = diagnostic => diagnostics.Add(diagnostic.Code + " " + diagnostic.Location + " " + diagnostic.Message)
        };
        string markdown = document.ToMarkdown(options).ToMarkdown();
        return (markdown, string.Join("\n", diagnostics));
    }

    // Built as unknown markup with inner XML, the way the package holds it when loaded: the Word elements inside stay untyped.
    private static OpenXmlUnknownElement Wrapper() {
        var wrapper = new OpenXmlUnknownElement("x", "wrapper", VendorNamespace);
        wrapper.InnerXml = $"<x:group xmlns:x=\"{VendorNamespace}\" xmlns:w=\"{WordNamespace}\">{WrappedMarkup}</x:group>";
        return wrapper;
    }
}
