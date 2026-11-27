using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// A <c>w:p</c> inside markup the package doesn't know (a vendor wrapper) is held untyped and isn't a paragraph of the
/// document: PDF export reads headers, notes and text boxes as if the wrapper weren't there.
/// </summary>
public sealed class WordPdfUnknownParagraphMarkupTests {
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Theory]
    [InlineData("header")]
    [InlineData("footnote")]
    [InlineData("text box")]
    public void AParagraphInsideUnknownMarkupDoesNotStopTheExport(string story) {
        string expected = ExportText(wrappedStory: null);

        string actual = ExportText(story);

        Assert.Equal(expected, actual);
        Assert.Contains("Header text", actual, StringComparison.Ordinal);
        Assert.Contains("Note text", actual, StringComparison.Ordinal);
        Assert.Contains("Box text", actual, StringComparison.Ordinal);
    }

    private static string ExportText(string? wrappedStory) {
        using var document = WordDocument.Create();
        WordParagraph header = document.PageSetup.Header.AddParagraph("Header text");
        WordFootNote note = document.AddParagraph("Body text").AddFootNote("Note text");
        WordTextBox box = document.AddTextBox("Box text");
        WordParagraph? wrapped = wrappedStory switch {
            "header" => header,
            "footnote" => note.Paragraphs!.Last(),
            "text box" => box.Paragraphs.Last(),
            _ => null
        };
        wrapped?._paragraph.AppendChild(Wrapper());

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions {
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        return PdfReadDocument.Open(result.ToBytes(), new PdfLoadOptions { IncludeArtifactText = true }).ExtractText();
    }

    // Built as unknown markup with inner XML, the way the package holds it when loaded: the paragraph inside stays untyped.
    private static OpenXmlUnknownElement Wrapper() {
        var wrapper = new OpenXmlUnknownElement("x", "wrapper", "urn:officeimo:tests:vendor");
        wrapper.InnerXml = $"<w:p xmlns:w=\"{WordNamespace}\"><w:r><w:t>Wrapped</w:t></w:r></w:p>";
        return wrapper;
    }
}
