using System.Globalization;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave41RegressionTests {
    [Theory]
    [InlineData(BibliographyItemType.Book)]
    [InlineData(BibliographyItemType.ArticleJournal)]
    public void Canonical_NBIB_output_reopens_through_automatic_detection(BibliographyItemType type) {
        var document = new BibliographyDocument(BibliographyFormat.Nbib);
        var item = new BibliographyItem { Key = "123", Type = type, Title = "Detected" };
        item.Identifiers.Add(new BibliographyIdentifier("PMID", "123"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDocument reopened = BibliographyDocument.Parse(written.Content).Document;

        Assert.Equal(BibliographyFormat.Nbib, reopened.SourceFormat);
        Assert.Equal(type, Assert.Single(reopened.Items).Type);
    }

    [Fact]
    public void Date_component_scanning_rejects_excess_fragments_without_substring_amplification() {
        string value = string.Join("-", Enumerable.Repeat("1", 100_000));
#if NET8_0_OR_GREATER
        long before = GC.GetAllocatedBytesForCurrentThread();
#endif
        BibliographyDate parsed = CodecMappings.ParseDate(BibliographyDateRole.Issued, value);
#if NET8_0_OR_GREATER
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.True(allocated < 1_000_000, $"Date parsing allocated {allocated} bytes.");
#endif
        Assert.Equal(value, parsed.Literal);
    }

    [Theory]
    [InlineData("TI  - Before\nPMID- 1", 0)]
    [InlineData("PT  - Mystery\nTI  - Before\nPMID- 1", 1)]
    public void NBIB_sources_without_recognized_publication_types_do_not_gain_synthetic_tags(string source, int expectedTypeTags) {
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        int typeTags = written.Content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Count(line => line.StartsWith("PT  -", StringComparison.Ordinal));

        Assert.Equal(expectedTypeTags, typeTags);
        Assert.Equal(BibliographyItemType.ArticleJournal, Assert.Single(BibliographyDocument.Parse(written.Content).Document.Items).Type);
    }

    [Fact]
    public void Whitespace_CSL_keys_remain_diagnosed_when_a_native_id_is_retained() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":123,\"type\":\"book\"}]", BibliographyFormat.CslJson).Document;
        document.Items[0].Key = " ";

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV215" && diagnostic.Field == "key");
    }

    [Theory]
    [InlineData(BibliographyFormat.BibLatex, "@book{x, title={Before}, date={Spring }}")]
    [InlineData(BibliographyFormat.Ris, "TY  - JOUR\nID  - x\nTI  - Before\nPY  - Spring \nER  -")]
    [InlineData(BibliographyFormat.Nbib, "PT  - Journal Article\nTI  - Before\nDP  - Spring \nPMID- x")]
    [InlineData(BibliographyFormat.EndNoteXml, "<xml><records><record><rec-number>x</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles><dates><pub-dates><date>Spring </date></pub-dates></dates></record></records></xml>")]
    public void Literal_date_whitespace_is_retained_for_strict_diagnostics(BibliographyFormat format, string source) {
        BibliographyDocument document = BibliographyDocument.Parse(source, format).Document;
        BibliographyDate date = Assert.Single(document.Items[0].Dates);
        document.Items[0].Title = "After";

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Equal("Spring ", date.Literal);
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV246" && diagnostic.Field == "dates.Issued.literal");
    }

    [Theory]
    [InlineData(17, BibliographyItemType.ArticleJournal)]
    [InlineData(6, BibliographyItemType.Book)]
    [InlineData(5, BibliographyItemType.Chapter)]
    [InlineData(47, BibliographyItemType.PaperConference)]
    [InlineData(27, BibliographyItemType.Report)]
    [InlineData(32, BibliographyItemType.Thesis)]
    [InlineData(12, BibliographyItemType.WebPage)]
    [InlineData(21, BibliographyItemType.Patent)]
    [InlineData(13, BibliographyItemType.Document)]
    public void Numeric_only_EndNote_reference_types_map_to_typed_items(int code, BibliographyItemType expected) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type>" + code.ToString(CultureInfo.InvariantCulture) + "</ref-type></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal(expected, reopened.Type);
        Assert.Equal(expected, document.Items[0].Type);
    }

    [Fact]
    public void Unsupported_numeric_EndNote_reference_types_block_strict_output() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type>999</ref-type></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV200" && diagnostic.Field == "type");
    }

    [Theory]
    [InlineData("<custom>v</custom><isbn/>", "custom,isbn")]
    [InlineData("<titles><secondary-title>Journal</secondary-title></titles><custom>v</custom><periodical><full-title>Other</full-title></periodical>", "custom,periodical")]
    [InlineData("<custom>v</custom><urls><related-urls><url>one</url><url>two</url></related-urls></urls>", "custom,url")]
    public void EndNote_native_fields_follow_source_order(string fields, string expectedNames) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type>" + fields + "</record></records></xml>";
        BibliographyItem item = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal(expectedNames.Split(','), item.NativeFields.Select(static field => field.Name));
    }

    [Fact]
    public void EndNote_unknown_fields_stay_before_later_blank_identifiers_after_canonical_edits() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><custom>v</custom><isbn/></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal(new[] { "custom", "isbn" }, reopened.NativeFields.Select(static field => field.Name));
    }
}
