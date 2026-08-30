namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave20RegressionTests {
    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void Blank_CSL_keywords_remain_distinct_from_an_absent_list(string keyword) {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Keywords" };
        item.Keywords.Add(keyword);
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        Assert.Equal(keyword, Assert.Single(reopened.Keywords));
    }

    [Fact]
    public void Recognized_EndNote_native_type_aliases_survive_strict_canonical_edits() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"article\">17</ref-type><titles><title>Before</title></titles></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal(BibliographyItemType.ArticleJournal, reopened.Type);
        Assert.Equal("article", reopened.NativeType);
    }

    [Fact]
    public void Automatic_detection_honors_pre_canceled_tokens_before_format_errors() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            BibliographyDocument.Parse("not bibliography", cancellationToken: cancellation.Token));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Tagged_value_limits_apply_before_large_line_materialization(bool continuation) {
        _ = BibliographyDocument.Parse("TY  - JOUR\nER  -", BibliographyFormat.Ris);
        string prefix = continuation ? "TY  - JOUR\nTI  - x\n      " : "TY  - ";
        string source = prefix + new string('x', 8 * 1024 * 1024);
        var options = new BibliographyReadOptions { MaximumValueLength = 4 };
        long before = GC.GetAllocatedBytesForCurrentThread();

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.Ris, options);
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
        Assert.True(allocated < 1024 * 1024, $"Oversized tagged input allocated {allocated:N0} bytes before rejection.");
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void Blank_EndNote_dates_remain_distinct_from_an_absent_date(string literal) {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Date" };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Literal = literal });
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items).Dates);

        Assert.Equal(literal, reopened.Literal);
    }

    [Fact]
    public void Empty_Bib_field_names_recover_with_diagnostics() {
        const string source = "@book{x, = {value}, title={Retained}}";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex);

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBBIB005");
        Assert.Equal("Retained", Assert.Single(read.Document.Items).Title);
    }
}
