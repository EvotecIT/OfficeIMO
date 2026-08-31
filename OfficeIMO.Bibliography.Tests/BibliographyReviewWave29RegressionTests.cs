namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave29RegressionTests {
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Blank_NBIB_serial_tags_survive_unrelated_strict_edits(string value) {
        string source = "PMID- 1\nPT  - Journal Article\nTI  - Before\nIS  - " + value + "\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Nbib).Document.Items);

        BibliographyNativeField field = Assert.Single(reopened.NativeFields, candidate => candidate.Format == BibliographyFormat.Nbib && candidate.Name == "IS");
        Assert.Equal(string.Empty, field.Value);
        Assert.Contains("IS  - ", written.Content, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Empty_Bib_keyword_positions_survive_unrelated_strict_edits(BibliographyFormat format) {
        BibliographyDocument document = BibliographyDocument.Parse("@book{x,title={Before},keywords={alpha,,beta,}}", format).Document;
        Assert.Equal(new[] { "alpha", "", "beta", "" }, document.Items[0].Keywords);
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items);

        Assert.Equal(new[] { "alpha", "", "beta", "" }, reopened.Keywords);
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Entirely_empty_Bib_keyword_fields_survive_unrelated_strict_edits(BibliographyFormat format) {
        BibliographyDocument document = BibliographyDocument.Parse("@book{x,title={Before},keywords={}}", format).Document;
        Assert.Equal(new[] { "" }, document.Items[0].Keywords);
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items);

        Assert.Equal(new[] { "" }, reopened.Keywords);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void EndNote_numeric_year_and_empty_publication_date_remain_distinct(string value) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles><dates><year>2026</year><pub-dates><date>" + value + "</date></pub-dates></dates></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items[0].Dates);

        Assert.Equal(2026, reopened.Year);
        BibliographyNativeField date = Assert.Single(reopened.NativeFields, field => field.Format == BibliographyFormat.EndNoteXml && field.Name == "date");
        Assert.Equal(value, date.Value);
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Odd_terminal_Bib_backslashes_never_consume_following_fields_or_records(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        document.Items.Add(new BibliographyItem { Key = "first", Type = BibliographyItemType.Book, Title = "Danger\\", Publisher = "Publisher" });
        document.Items.Add(new BibliographyItem { Key = "second", Type = BibliographyItemType.Book, Title = "Second" });

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV133" && diagnostic.Field == "title");

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyDocument reopened = BibliographyDocument.Parse(written.Content, format).Document;

        Assert.Contains(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV133" && diagnostic.Field == "title");
        Assert.Equal(2, reopened.Items.Count);
        Assert.Equal("Publisher", reopened.Items[0].Publisher);
        Assert.Equal("Second", reopened.Items[1].Title);
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Even_terminal_Bib_backslashes_reopen_exactly(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Exact\\\\", Publisher = "Publisher" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items);

        Assert.Equal("Exact\\\\", reopened.Title);
        Assert.Equal("Publisher", reopened.Publisher);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV133");
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Odd_terminal_native_Bib_backslashes_are_diagnosed_and_safely_bounded(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "first", Type = BibliographyItemType.Book, Title = "First" };
        item.NativeFields.Add(new BibliographyNativeField(format, "custom", "Native\\"));
        document.Items.Add(item);
        document.Items.Add(new BibliographyItem { Key = "second", Type = BibliographyItemType.Book, Title = "Second" });

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV133" && diagnostic.Field == "custom");

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyDocument reopened = BibliographyDocument.Parse(written.Content, format).Document;
        Assert.Equal(2, reopened.Items.Count);
        Assert.Contains(reopened.Items[0].NativeFields, field => field.Name == "custom");
        Assert.Equal("Second", reopened.Items[1].Title);
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Nested_Bib_values_with_terminal_backslashes_are_omitted_instead_of_malformed(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "first", Type = BibliographyItemType.Book, Title = "First", Publisher = "Publisher" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Literal = "Organization\\" }));
        document.Items.Add(item);
        document.Items.Add(new BibliographyItem { Key = "second", Type = BibliographyItemType.Book, Title = "Second" });

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV134" && diagnostic.Field == "author");

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyDocument reopened = BibliographyDocument.Parse(written.Content, format).Document;
        Assert.Equal(2, reopened.Items.Count);
        Assert.Empty(reopened.Items[0].Contributors);
        Assert.Equal("Publisher", reopened.Items[0].Publisher);
        Assert.Equal("Second", reopened.Items[1].Title);
    }
}
