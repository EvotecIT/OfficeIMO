using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave48RegressionTests {
    [Theory]
    [InlineData("", "")]
    [InlineData(" ", "")]
    [InlineData("", " ")]
    [InlineData(" ", "  ")]
    public void EndNote_preserves_distinct_blank_year_and_publication_date_components(string year, string publicationDate) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles><dates><year>" + year + "</year><pub-dates><date>" + publicationDate + "</date></pub-dates></dates></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items[0].Dates);

        Assert.Equal(publicationDate, reopened.Literal);
        Assert.Equal(year, Assert.Single(reopened.NativeFields, field => field.Name == "year").Value);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void EndNote_preserves_blank_year_only_date_carriers(string year) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles><dates><year>" + year + "</year></dates></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items[0].Dates);

        Assert.Equal(year, reopened.Literal);
        Assert.Equal(year, Assert.Single(reopened.NativeFields, field => field.Name == "year").Value);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void EndNote_preserves_blank_periodical_title_carriers(string title) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><periodical><full-title>" + title + "</full-title></periodical></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Publisher = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        XDocument output = XDocument.Parse(written.Content, LoadOptions.PreserveWhitespace);
        XElement periodical = Assert.Single(output.Descendants(), element => element.Name.LocalName == "periodical");

        Assert.Equal(title, Assert.Single(periodical.Elements(), element => element.Name.LocalName == "full-title").Value);
        Assert.DoesNotContain("secondary-title", written.Content, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void EndNote_preserves_blank_redundant_periodical_title_carriers(string title) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><secondary-title>Journal</secondary-title></titles><periodical><full-title>" + title + "</full-title></periodical></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Publisher = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);
        BibliographyNativeField nativePeriodical = Assert.Single(reopened.NativeFields, field => field.Name == "periodical");
        XElement periodical = XElement.Parse(nativePeriodical.RawValue!, LoadOptions.PreserveWhitespace);

        Assert.Equal("Journal", reopened.ContainerTitle);
        Assert.Equal(title, Assert.Single(periodical.Elements(), element => element.Name.LocalName == "full-title").Value);
    }

    [Fact]
    public void NBIB_publication_type_continuations_scale_linearly_and_keep_the_first_recognized_type() {
        const int count = 20_000;
        var source = new StringBuilder(count * 32);
        source.Append("PT  - Journal\n      Article\n");
        for (int index = 1; index < count; index++) source.Append("PT  - Unknown\n      Type\n");
        source.Append("PMID- 1\n");

        BibliographyItem item = Assert.Single(BibliographyDocument.Parse(source.ToString(), BibliographyFormat.Nbib).Document.Items);

        Assert.Equal(BibliographyItemType.ArticleJournal, item.Type);
        Assert.Equal("Journal Article", item.NativeType);
        Assert.Equal(count, item.NativeFields.Count(field => field.Name == "PT"));
    }

    [Fact]
    public void NBIB_continued_type_can_become_the_first_recognized_publication_type() {
        const string source = "PT  - Conference\n      Paper\nPT  - Journal Article\nPMID- 1\n";

        BibliographyItem item = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document.Items);

        Assert.Equal(BibliographyItemType.PaperConference, item.Type);
        Assert.Equal("Conference Paper", item.NativeType);
    }

    [Theory]
    [InlineData("item", "BIBCONV126")]
    [InlineData("name", "BIBCONV127")]
    [InlineData("date", "BIBCONV128")]
    public void Deep_native_CSL_values_that_exceed_default_reopen_depth_block_strict_output(string owner, string diagnosticCode) {
        string nested = "true";
        for (int index = 0; index < 130; index++) nested = "{\"level\":" + nested + "}";
        BibliographyDocument document = CreateCslDocumentWithNative(owner, nested);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyReadResult reopened = BibliographyDocument.Parse(permissive.Content, BibliographyFormat.CslJson);

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV247" || diagnostic.Code == "BIBCONV126");
        Assert.Contains(permissive.Report.Diagnostics, diagnostic => diagnostic.Code == diagnosticCode && diagnostic.Field!.EndsWith("custom", StringComparison.Ordinal));
        Assert.False(reopened.HasErrors);
    }

    [Theory]
    [InlineData("item")]
    [InlineData("name")]
    [InlineData("date")]
    public void Native_CSL_values_within_default_reopen_depth_remain_exact(string owner) {
        string nested = "true";
        for (int index = 0; index < 100; index++) nested = "{\"level\":" + nested + "}";
        BibliographyDocument document = CreateCslDocumentWithNative(owner, nested);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyReadResult reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson);

        Assert.False(reopened.HasErrors);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV126");
    }

    [Theory]
    [InlineData("custom")]
    [InlineData("")]
    [InlineData(" ")]
    public void Unpreservable_CSL_native_types_block_strict_output(string nativeType) {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, NativeType = nativeType });

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV200" && diagnostic.Field == "type");
    }

    [Fact]
    public void Canonical_CSL_native_type_values_remain_exact_for_created_items() {
        var document = new BibliographyDocument(BibliographyFormat.BibTex);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, NativeType = "book" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.CslJson, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Equal("book", Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items).NativeType);
    }

    private static BibliographyDocument CreateCslDocumentWithNative(string owner, string raw) {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "After" };
        var field = new BibliographyNativeField(BibliographyFormat.CslJson, "custom", raw, raw);
        if (owner == "item") item.NativeFields.Add(field);
        else if (owner == "name") {
            var name = new BibliographyName { Family = "Smith" };
            name.NativeFields.Add(field);
            item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, name));
        } else {
            var date = new BibliographyDate { Role = BibliographyDateRole.Issued, Year = 2026 };
            date.NativeFields.Add(field);
            item.Dates.Add(date);
        }
        document.Items.Add(item);
        return document;
    }
}
