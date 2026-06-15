using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlListTests {
    [Fact]
    public void Html_ToRtfDocument_Assigns_Rtf_List_Ids_And_Levels() {
        const string html = "<ul><li>Allergy</li><li>Medication</li></ul><ol><li>Step</li></ol>";

        RtfDocument document = html.ToRtfDocument();

        Assert.Equal(RtfListKind.Bullet, document.Paragraphs[0].ListKind);
        Assert.Equal(1, document.Paragraphs[0].ListId);
        Assert.Equal(0, document.Paragraphs[0].ListLevel);
        Assert.Equal(RtfListKind.Bullet, document.Paragraphs[1].ListKind);
        Assert.Equal(1, document.Paragraphs[1].ListId);
        Assert.Equal(0, document.Paragraphs[1].ListLevel);
        Assert.Equal(RtfListKind.Decimal, document.Paragraphs[2].ListKind);
        Assert.Equal(2, document.Paragraphs[2].ListId);
        Assert.Equal(0, document.Paragraphs[2].ListLevel);

        string rtf = document.ToRtf();
        Assert.Contains(@"{\*\listtable", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\listoverridetable", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ls1\ilvl0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ls2\ilvl0", rtf, StringComparison.Ordinal);

        RtfDocument roundTrip = RtfDocument.Read(rtf).Document;
        Assert.Equal(1, roundTrip.Paragraphs[0].ListId);
        Assert.Equal(1, roundTrip.Paragraphs[1].ListId);
        Assert.Equal(2, roundTrip.Paragraphs[2].ListId);
        Assert.Equal(RtfListKind.Bullet, roundTrip.Paragraphs[0].ListKind);
        Assert.Equal(RtfListKind.Decimal, roundTrip.Paragraphs[2].ListKind);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_List_Metadata_Attributes() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Step")
            .SetList(listId: 42, level: 2, kind: RtfListKind.Decimal)
            .SetListText("7.");
        paragraph.ListDefinitionId = 100;

        string html = document.ToHtml();

        Assert.Equal("<ol><li data-officeimo-rtf-list-id=\"42\" data-officeimo-rtf-list-definition-id=\"100\" data-officeimo-rtf-list-level=\"2\" data-officeimo-rtf-list-text=\"7.\">Step</li></ol>", html);

        RtfParagraph roundTripParagraph = Assert.Single(html.ToRtfDocument().Paragraphs);
        Assert.Equal(RtfListKind.Decimal, roundTripParagraph.ListKind);
        Assert.Equal(42, roundTripParagraph.ListId);
        Assert.Equal(100, roundTripParagraph.ListDefinitionId);
        Assert.Equal(2, roundTripParagraph.ListLevel);
        Assert.Equal("7.", roundTripParagraph.ListText?.ToPlainText());

        string rtf = html.ToRtfDocument().ToRtf();
        Assert.Contains(@"\ls42\ilvl2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\listtext 7.}", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_List_Metadata_Attributes() {
        const string html = "<ul><li data-officeimo-rtf-list-kind=\"decimal\" data-officeimo-rtf-list-id=\"9\" data-officeimo-rtf-list-definition-id=\"90\" data-officeimo-rtf-list-level=\"3\" data-officeimo-rtf-list-text=\"3.\">Nested</li></ul>";

        RtfParagraph paragraph = Assert.Single(html.ToRtfDocument().Paragraphs);

        Assert.Equal(RtfListKind.Decimal, paragraph.ListKind);
        Assert.Equal(9, paragraph.ListId);
        Assert.Equal(90, paragraph.ListDefinitionId);
        Assert.Equal(3, paragraph.ListLevel);
        Assert.Equal("3.", paragraph.ListText?.ToPlainText());
    }
}
