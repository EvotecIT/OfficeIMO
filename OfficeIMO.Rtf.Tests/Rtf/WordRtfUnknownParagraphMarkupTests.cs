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

    [Theory]
    [InlineData("run", "<w:fldChar w:fldCharType=\"begin\"/>")]
    [InlineData("run", "<w:footnoteReference w:id=\"1\"/>")]
    [InlineData("run", "<w:endnoteReference w:id=\"1\"/>")]
    [InlineData("run", "<w:br w:type=\"page\"/>")]
    [InlineData("table cell", "<w:p><w:r><w:t>Hidden</w:t></w:r></w:p>")]
    [InlineData("table cell", "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Hidden</w:t></w:r></w:p></w:tc></w:tr></w:tbl>")]
    [InlineData("section", "<w:pgNumType w:start=\"47\"/>")]
    [InlineData("table properties", "<w:tblW w:w=\"1234\" w:type=\"dxa\"/>")]
    public void UntypedWordMarkupInsideTheModelIsSkipped(string parent, string markup) {
        (string expectedRtf, string expectedReport) = ExportWithUntypedMarkup(parent, markup: null);

        (string actualRtf, string actualReport) = ExportWithUntypedMarkup(parent, markup);

        Assert.Equal(expectedRtf, actualRtf);
        Assert.Equal(expectedReport, actualReport);
    }

    [Fact]
    public void AnXmlCommentKeepsAnImageInsideItsParagraph() {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        using (var stream = new MemoryStream(OnePixelPng(), writable: false)) {
            paragraph.AddImage(stream, "inline.png", OfficeIMO.OfficeLength.Pixels(16), OfficeIMO.OfficeLength.Pixels(16),
                options: new WordImageOptions { AltText = "Inline image" });
        }
        Assert.IsType<RtfImage>(Assert.Single(document.ToRtf().Blocks));

        paragraph._paragraph.AppendChild(new OpenXmlMiscNode(System.Xml.XmlNodeType.Comment, "<!--note-->"));

        RtfParagraph converted = Assert.IsType<RtfParagraph>(Assert.Single(document.ToRtf().Blocks));
        Assert.Equal("Inline image", Assert.IsType<RtfImage>(Assert.Single(converted.Inlines)).Description);
    }

    private static (string Rtf, string Report) ExportWithUntypedMarkup(string parent, string? markup) {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Body text");
        paragraph.AddFootNote("Note text");
        paragraph.AddEndNote("End text");
        paragraph.AddText(" after");
        WordTable table = document.AddTable(1, 1);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Cell text";
        if (markup != null) {
            OpenXmlElement target = parent switch {
                "run" => paragraph._paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().First(),
                "table cell" => table.Rows[0].Cells[0]._tableCell,
                "section" => document._wordprocessingDocument!.MainDocumentPart!.Document!.Body!
                    .Elements<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>().Last(),
                _ => table._table.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.TableProperties>()!
            };
            OpenXmlElement element = UntypedWordElement(markup);
            target.PrependChild(element);
        }

        RtfDocument rtf = document.ToRtf(null, out OfficeConversionReport report);
        string diagnostics = string.Join("\n", report.Diagnostics.Select(diagnostic =>
            diagnostic.Code + " " + diagnostic.Message));
        return (rtf.ToRtf(new RtfSaveOptions { IncludeGenerator = false }), diagnostics);
    }

    // Word markup parsed inside a vendor group keeps no SDK type; detached, it stays untyped.
    private static OpenXmlElement UntypedWordElement(string markup) {
        var wrapper = new OpenXmlUnknownElement("x", "wrapper", VendorNamespace) {
            InnerXml = $"<x:group xmlns:x=\"{VendorNamespace}\" xmlns:w=\"{WordNamespace}\">{markup}</x:group>"
        };
        OpenXmlElement element = wrapper.FirstChild!.FirstChild!;
        element.Remove();
        return element;
    }

    private static byte[] OnePixelPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=");

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
