using DocumentFormat.OpenXml;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

/// <summary>
/// Word markup inside a wrapper the package doesn't know (a vendor element) is held untyped and isn't part of the
/// document model: RTF export and its report read every story as if the wrapper weren't there. Word run markup the
/// package holds untyped beside a paragraph's runs is skipped the same way.
/// </summary>
public sealed class WordRtfUnknownParagraphMarkupTests {
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string VendorNamespace = "urn:officeimo:tests:vendor";

    // A framed paragraph, a footnote reference, a comment anchor and a run outside any paragraph, all inside the wrapper.
    private const string WrappedMarkup =
        "<w:p><w:pPr><w:framePr w:w=\"2000\" w:h=\"400\"/></w:pPr><w:r><w:t>Wrapped</w:t></w:r></w:p>" +
        "<w:r><w:footnoteReference w:id=\"1\"/></w:r>" +
        "<w:commentRangeStart w:id=\"7\"/><w:r><w:commentReference w:id=\"7\"/></w:r>";

    [Theory]
    [InlineData("body")]
    [InlineData("run")]
    [InlineData("table cell")]
    [InlineData("header")]
    [InlineData("footnote")]
    [InlineData("text box")]
    public void MarkupInsideAnUnknownWrapperDoesNotChangeTheExport(string story) {
        (string expectedRtf, string expectedReport) = Export(wrappedStory: null);

        (string actualRtf, string actualReport) = Export(story);

        Assert.Equal(expectedRtf, actualRtf);
        Assert.Equal(expectedReport, actualReport);
        Assert.DoesNotContain("Wrapped", actualRtf, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("paragraph", "sdtContent")]
    [InlineData("paragraph", "sdt")]
    [InlineData("paragraph", "r")]
    [InlineData("hyperlink", "sdtContent")]
    [InlineData("hyperlink", "r")]
    public void UntypedRunMarkupBesideTheRunsIsSkipped(string parent, string element) {
        (string expectedRtf, string expectedReport) = ExportWithUntypedChild(parent, element: null);

        (string actualRtf, string actualReport) = ExportWithUntypedChild(parent, element);

        Assert.Equal(expectedRtf, actualRtf);
        Assert.Equal(expectedReport, actualReport);
        Assert.DoesNotContain("Hidden", actualRtf, StringComparison.Ordinal);
    }

    private static (string Rtf, string Report) ExportWithUntypedChild(string parent, string? element) {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Body text");
        paragraph.AddHyperLink("Link text", new Uri("https://example.test"));
        if (element != null) {
            string run = "<w:r><w:t>Hidden</w:t></w:r>";
            string markup = element switch {
                "sdtContent" => $"<w:sdtContent xmlns:w=\"{WordNamespace}\">{run}</w:sdtContent>",
                "sdt" => $"<w:sdt xmlns:w=\"{WordNamespace}\"><w:sdtContent>{run}</w:sdtContent></w:sdt>",
                _ => $"<w:r xmlns:w=\"{WordNamespace}\"><w:t>Hidden</w:t></w:r>"
            };
            OpenXmlElement target = parent == "hyperlink"
                ? paragraph._paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Hyperlink>().Single()
                : paragraph._paragraph;
            target.AppendChild(UntypedElement(markup));
        }

        RtfDocument rtf = document.ToRtf(null, out OfficeConversionReport report);
        string diagnostics = string.Join("\n", report.Diagnostics.Select(diagnostic =>
            diagnostic.Code + " " + diagnostic.Message));
        return (rtf.ToRtf(new RtfSaveOptions { IncludeGenerator = false }), diagnostics);
    }

    // Word markup parsed inside an unknown wrapper stays untyped; detached, it keeps that.
    private static OpenXmlElement UntypedElement(string outerXml) {
        var wrapper = new OpenXmlUnknownElement("x", "wrapper", VendorNamespace) { InnerXml = outerXml };
        OpenXmlElement element = wrapper.FirstChild!;
        element.Remove();
        return element;
    }

    private static (string Rtf, string Report) Export(string? wrappedStory) {
        using var document = WordDocument.Create();
        WordParagraph header = document.PageSetup.Header.AddParagraph("Header text");
        WordParagraph body = document.AddParagraph("Body text");
        WordFootNote note = body.AddFootNote("Note text");
        WordTable table = document.AddTable(1, 1);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Cell text";
        WordTextBox box = document.AddTextBox("Box text");

        OpenXmlElement? target = wrappedStory switch {
            "body" => body._paragraph,
            "run" => body._paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().First(),
            "table cell" => table.Rows[0].Cells[0]._tableCell,
            "header" => header._paragraph,
            "footnote" => note.Paragraphs!.Last()._paragraph,
            "text box" => box.Paragraphs.Last()._paragraph,
            _ => null
        };
        target?.AppendChild(Wrapper());

        RtfDocument rtf = document.ToRtf(null, out OfficeConversionReport report);
        string diagnostics = string.Join("\n", report.Diagnostics.Select(diagnostic =>
            diagnostic.Code + " " + diagnostic.Message));
        return (rtf.ToRtf(new RtfSaveOptions { IncludeGenerator = false }), diagnostics);
    }

    // Built as unknown markup with inner XML, the way the package holds it when loaded: the Word elements inside stay untyped.
    private static OpenXmlUnknownElement Wrapper() {
        var wrapper = new OpenXmlUnknownElement("x", "wrapper", VendorNamespace);
        wrapper.InnerXml = $"<x:group xmlns:x=\"{VendorNamespace}\" xmlns:w=\"{WordNamespace}\">{WrappedMarkup}</x:group>";
        return wrapper;
    }
}
