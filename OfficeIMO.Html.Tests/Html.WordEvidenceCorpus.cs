using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    [Trait("Category", "OfficeInteroperability")]
    public void HtmlToWord_WordEvidenceCorpus_RoundTripsSupportedStructures() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Word", "EvidenceCorpus", "word-html-reciprocal.html");
        string html = File.ReadAllText(path);

        HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(new HtmlToWordOptions());
        using var document = conversion.RequireValue();

        Assert.Single(document.Tables);
        Assert.Equal("Regional results", document.Tables[0].Title);
        Assert.Equal("A nested table and list corpus", document.Tables[0].Description);
        Assert.Equal(2, document._wordprocessingDocument.MainDocumentPart!.Document.Descendants<Ruby>().Count());
        Assert.NotNull(document.Sections[0].Header.Default);
        Assert.NotNull(document.Sections[0].Footer.Default);
        Assert.Contains(document.Sections[0].Header.Default!.Paragraphs, paragraph => paragraph.Text == "Evidence header");
        Assert.Contains(document.Sections[0].Footer.Default!.Paragraphs, paragraph => paragraph.Text == "Evidence footer");
        Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "HtmlRubyPairingApproximation");
        var validationErrors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2010)
            .Validate(document._wordprocessingDocument)
            .ToList();
        Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine, validationErrors.Select(error =>
            error.Description + " Path=" + error.Path?.XPath + " Node=" + error.Node?.OuterXml)));

        string roundTrip = document.ToHtml(new WordToHtmlOptions {
            ExportHeadersAndFooters = true,
            IncludeSectionMetadata = true
        });
        int rubyEast = roundTrip.IndexOf("<rb>東</rb><rt>とう</rt>", StringComparison.OrdinalIgnoreCase);
        int rubyCapital = roundTrip.IndexOf("<rb>京</rb><rt>きょう</rt>", StringComparison.OrdinalIgnoreCase);
        int nested = roundTrip.IndexOf("Nested", StringComparison.Ordinal);
        int following = roundTrip.IndexOf("Comment and section carrier", StringComparison.Ordinal);
        Assert.True(rubyEast >= 0 && rubyCapital > rubyEast, roundTrip);
        Assert.True(nested >= 0 && following > nested, roundTrip);
        Assert.Contains("aria-label=\"Regional results\"", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("aria-description=\"A nested table and list corpus\"", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<figure>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<figcaption", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("class=\"word-header word-header-default\"", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("class=\"word-footer word-footer-default\"", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("class=\"word-section\"", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("type=\"checkbox\"", roundTrip, StringComparison.OrdinalIgnoreCase);
    }
}
