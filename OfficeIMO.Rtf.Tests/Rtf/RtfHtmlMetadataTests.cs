using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlMetadataTests {
    [Fact]
    public void RtfDocument_ToHtml_Renders_Document_Info_Metadata_And_RoundTrips() {
        RtfDocument document = RtfDocument.Create();
        document.Properties.Title = "Clinical note";
        document.Properties.Subject = "Discharge summary";
        document.Properties.Author = "Alice";
        document.Properties.Company = "Contoso Health";
        document.Properties.Keywords = "patient,rtf";
        document.Properties.Comments = "Reviewed";
        document.Properties.Created = new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc);
        document.Properties.NumberOfWords = 42;
        document.AddParagraph("Body");

        string html = document.ToHtml(new RtfHtmlExportOptions {
            IncludeRoundTripMetadata = true,
            EmbedImagesAsDataUri = true,
            FragmentOnly = false,
            NewLine = "\n"
        });

        Assert.Contains("<title>Clinical note</title>", html, StringComparison.Ordinal);
        Assert.Contains("<meta name=\"author\" content=\"Alice\">", html, StringComparison.Ordinal);
        Assert.Contains("<meta name=\"officeimo-rtf-company\" content=\"Contoso Health\">", html, StringComparison.Ordinal);
        Assert.Contains("<meta name=\"officeimo-rtf-created\" content=\"2026-01-02T03:04:05.0000000Z\">", html, StringComparison.Ordinal);
        Assert.Contains("<meta name=\"officeimo-rtf-words\" content=\"42\">", html, StringComparison.Ordinal);

        RtfDocument roundTrip = HtmlConversionDocument.Parse(html).ToRtf();
        Assert.Equal("Clinical note", roundTrip.Properties.Title);
        Assert.Equal("Discharge summary", roundTrip.Properties.Subject);
        Assert.Equal("Alice", roundTrip.Properties.Author);
        Assert.Equal("Contoso Health", roundTrip.Properties.Company);
        Assert.Equal("patient,rtf", roundTrip.Properties.Keywords);
        Assert.Equal("Reviewed", roundTrip.Properties.Comments);
        Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc), roundTrip.Properties.Created);
        Assert.Equal(42, roundTrip.Properties.NumberOfWords);
        Assert.Equal("Body", Assert.Single(roundTrip.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Html_ToRtfDocument_Does_Not_Treat_Head_Text_As_Body() {
        const string html = "<!doctype html><html><head><title>Clinical note</title><meta name=\"author\" content=\"Alice\"><meta name=\"keywords\" content=\"patient,rtf\"><meta name=\"description\" content=\"Summary\"></head><body><p>Body</p></body></html>";

        RtfDocument document = HtmlConversionDocument.Parse(html).ToRtf();

        Assert.Equal("Clinical note", document.Properties.Title);
        Assert.Equal("Alice", document.Properties.Author);
        Assert.Equal("patient,rtf", document.Properties.Keywords);
        Assert.Equal("Summary", document.Properties.Subject);
        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        Assert.Equal("Body", paragraph.ToPlainText());
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Header_Footer_Metadata_And_RoundTrips() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph header = document.AddHeader(RtfHeaderFooterKind.Header).AddParagraph("Header ");
        header.AddText("bold").SetBold();
        document.AddFooter(RtfHeaderFooterKind.Footer).AddParagraph("Footer");
        document.AddParagraph("Body");

        string html = document.ToHtml(new RtfHtmlExportOptions {
            IncludeRoundTripMetadata = true,
            EmbedImagesAsDataUri = true,
            FragmentOnly = false,
            NewLine = "\n"
        });

        Assert.Contains("<meta name=\"officeimo-rtf-header-footer\" content=\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-kind=\"header\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-kind=\"footer\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = HtmlConversionDocument.Parse(html).ToRtf();
        Assert.Equal("Body", Assert.Single(roundTrip.Paragraphs).ToPlainText());
        Assert.Equal(2, roundTrip.HeaderFooters.Count);
        RtfHeaderFooter roundTripHeader = roundTrip.HeaderFooters[0];
        Assert.Equal(RtfHeaderFooterKind.Header, roundTripHeader.Kind);
        Assert.Equal("Header bold", roundTripHeader.ToPlainText());
        Assert.Contains(roundTripHeader.Paragraphs[0].Runs, run => run.Text == "bold" && run.Bold);
        Assert.Equal(RtfHeaderFooterKind.Footer, roundTrip.HeaderFooters[1].Kind);
        Assert.Equal("Footer", roundTrip.HeaderFooters[1].ToPlainText());
    }
}
