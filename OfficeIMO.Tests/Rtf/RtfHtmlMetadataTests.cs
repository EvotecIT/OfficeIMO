using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlMetadataTests {
    [Fact]
    public void RtfDocument_ToHtml_Renders_Document_Info_Metadata_And_RoundTrips() {
        RtfDocument document = RtfDocument.Create();
        document.Info.Title = "Clinical note";
        document.Info.Subject = "Discharge summary";
        document.Info.Author = "Alice";
        document.Info.Company = "Contoso Health";
        document.Info.Keywords = "patient,rtf";
        document.Info.Comments = "Reviewed";
        document.Info.Created = new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc);
        document.Info.NumberOfWords = 42;
        document.AddParagraph("Body");

        string html = document.ToHtml(new RtfHtmlSaveOptions {
            FragmentOnly = false,
            NewLine = "\n"
        });

        Assert.Contains("<title>Clinical note</title>", html, StringComparison.Ordinal);
        Assert.Contains("<meta name=\"author\" content=\"Alice\">", html, StringComparison.Ordinal);
        Assert.Contains("<meta name=\"officeimo-rtf-company\" content=\"Contoso Health\">", html, StringComparison.Ordinal);
        Assert.Contains("<meta name=\"officeimo-rtf-created\" content=\"2026-01-02T03:04:05.0000000Z\">", html, StringComparison.Ordinal);
        Assert.Contains("<meta name=\"officeimo-rtf-words\" content=\"42\">", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadFromHtml();
        Assert.Equal("Clinical note", roundTrip.Info.Title);
        Assert.Equal("Discharge summary", roundTrip.Info.Subject);
        Assert.Equal("Alice", roundTrip.Info.Author);
        Assert.Equal("Contoso Health", roundTrip.Info.Company);
        Assert.Equal("patient,rtf", roundTrip.Info.Keywords);
        Assert.Equal("Reviewed", roundTrip.Info.Comments);
        Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc), roundTrip.Info.Created);
        Assert.Equal(42, roundTrip.Info.NumberOfWords);
        Assert.Equal("Body", Assert.Single(roundTrip.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Html_ToRtfDocument_Does_Not_Treat_Head_Text_As_Body() {
        const string html = "<!doctype html><html><head><title>Clinical note</title><meta name=\"author\" content=\"Alice\"><meta name=\"keywords\" content=\"patient,rtf\"><meta name=\"description\" content=\"Summary\"></head><body><p>Body</p></body></html>";

        RtfDocument document = html.LoadFromHtml();

        Assert.Equal("Clinical note", document.Info.Title);
        Assert.Equal("Alice", document.Info.Author);
        Assert.Equal("patient,rtf", document.Info.Keywords);
        Assert.Equal("Summary", document.Info.Subject);
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

        string html = document.ToHtml(new RtfHtmlSaveOptions {
            FragmentOnly = false,
            NewLine = "\n"
        });

        Assert.Contains("<meta name=\"officeimo-rtf-header-footer\" content=\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-kind=\"header\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-kind=\"footer\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadFromHtml();
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
