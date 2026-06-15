using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlFieldTests {
    [Fact]
    public void Html_ToRtfDocument_Parses_Field_Metadata_In_Inline_Order() {
        const string html = "<p>Page <span data-officeimo-rtf-field=\"true\" data-officeimo-rtf-field-instruction=\"PAGE \\* MERGEFORMAT\"><strong>1</strong></span> done</p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();
        RtfParagraph paragraph = Assert.Single(document.Paragraphs);

        Assert.Collection(paragraph.Inlines,
            inline => Assert.Equal("Page ", Assert.IsType<RtfRun>(inline).Text),
            inline => {
                RtfField field = Assert.IsType<RtfField>(inline);
                Assert.Equal(@"PAGE \* MERGEFORMAT", field.Instruction);
                Assert.Equal("1", field.ToPlainText());
                Assert.Contains(field.Result.Runs, run => run.Text == "1" && run.Bold);
            },
            inline => Assert.Equal(" done", Assert.IsType<RtfRun>(inline).Text));

        string rtf = document.ToRtf();
        Assert.Contains(@"{\field{\*\fldinst PAGE \\* MERGEFORMAT}{\fldrslt \b 1\b0 }}", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Field_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Page ");
        RtfField field = paragraph.AddField(@"PAGE \* MERGEFORMAT");
        field.AddText("1").SetBold();
        paragraph.AddText(" done");

        string html = document.ToHtml();

        Assert.Equal("<p>Page <span data-officeimo-rtf-field=\"true\" data-officeimo-rtf-field-instruction=\"PAGE \\* MERGEFORMAT\"><strong>1</strong></span> done</p>", html);

        RtfField roundTripField = Assert.IsType<RtfField>(Assert.Single(html.ToRtfDocumentFromHtml().Paragraphs).Inlines[1]);
        Assert.Equal(@"PAGE \* MERGEFORMAT", roundTripField.Instruction);
        Assert.Equal("1", roundTripField.ToPlainText());
        Assert.Contains(roundTripField.Result.Runs, run => run.Text == "1" && run.Bold);
    }

    [Fact]
    public void RtfDocument_ToHtml_Escapes_Field_Instruction_Attribute() {
        RtfDocument document = RtfDocument.Create();
        RtfField field = document.AddParagraph().AddField("MERGEFIELD Patient<Name>");
        field.AddText("Ada");

        string html = document.ToHtml();

        Assert.Equal("<p><span data-officeimo-rtf-field=\"true\" data-officeimo-rtf-field-instruction=\"MERGEFIELD Patient&lt;Name&gt;\">Ada</span></p>", html);
        RtfField roundTripField = Assert.IsType<RtfField>(Assert.Single(html.ToRtfDocumentFromHtml().Paragraphs).Inlines[0]);
        Assert.Equal("MERGEFIELD Patient<Name>", roundTripField.Instruction);
    }

    [Fact]
    public void Html_ToRtfDocument_DoesNot_Preserve_Field_Span_As_Unknown_Text() {
        const string html = "<p><span data-officeimo-rtf-field=\"true\" data-officeimo-rtf-field-instruction=\"PAGE\">1</span></p>";

        RtfParagraph paragraph = Assert.Single(html.ToRtfDocumentFromHtml(new RtfHtmlReadOptions { PreserveUnknownTagsAsText = true }).Paragraphs);

        RtfField field = Assert.IsType<RtfField>(Assert.Single(paragraph.Inlines));
        Assert.Equal("PAGE", field.Instruction);
        Assert.Equal("1", field.ToPlainText());
    }
}
