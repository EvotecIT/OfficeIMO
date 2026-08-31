namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave22RegressionTests {
    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Retained_Bib_scalars_cannot_be_promoted_after_the_typed_owner_is_cleared(BibliographyFormat format) {
        BibliographyDocument document = BibliographyDocument.Parse("@Book{x,title={First},title={Second}}", format).Document;
        document.Items[0].Title = null;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV119" && diagnostic.Field == "title");
    }

    [Fact]
    public void Retained_CSL_strings_cannot_be_promoted_after_the_typed_owner_is_cleared() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"x\",\"type\":\"book\",\"title\":\"First\",\"title\":\"Second\"}]", BibliographyFormat.CslJson).Document;
        document.Items[0].Title = null;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV120" && diagnostic.Field == "title");
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris, "TY  - JOUR\nID  - x\nTI  - First\nTI  - Second\nER  -")]
    [InlineData(BibliographyFormat.Nbib, "PMID- x\nTI  - First\nTI  - Second")]
    public void Retained_tagged_scalars_cannot_be_promoted_after_the_typed_owner_is_cleared(BibliographyFormat format, string source) {
        BibliographyDocument document = BibliographyDocument.Parse(source, format).Document;
        document.Items[0].Title = null;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV122" && diagnostic.Field == "TI");
    }

    [Fact]
    public void Parseable_retained_CSL_contributors_and_dates_cannot_change_owners() {
        const string source = "[{\"id\":\"x\",\"type\":\"book\",\"author\":[{\"family\":\"First\"}],\"author\":[{\"family\":\"Second\"}],\"issued\":{\"date-parts\":[[2020]]},\"issued\":{\"date-parts\":[[2021]]}}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        document.Items[0].Contributors.Clear();
        document.Items[0].Dates.Clear();

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV120" && diagnostic.Field == "author");
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV120" && diagnostic.Field == "issued");
    }

    [Fact]
    public void Retained_CSL_date_parts_cannot_become_typed_after_the_typed_value_is_cleared() {
        const string source = "[{\"id\":\"x\",\"type\":\"book\",\"issued\":{\"date-parts\":[[2020]],\"date-parts\":[[2021]]}}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        BibliographyDate issued = Assert.Single(document.Items[0].Dates);
        issued.Year = null;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV125" && diagnostic.Field == "issued.date-parts");
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Canonical_Bib_edits_preserve_native_type_spelling(BibliographyFormat format) {
        BibliographyDocument document = BibliographyDocument.Parse("@Book{x,title={Before}}", format).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items);

        Assert.StartsWith("@Book{", written.Content, StringComparison.Ordinal);
        Assert.Equal("Book", reopened.NativeType);
    }

    [Theory]
    [InlineData("SP", "SP  - 20", "EP  - 20")]
    [InlineData("EP", "EP  - 20", "SP  - 20")]
    public void Canonical_RIS_edits_preserve_single_page_endpoint_roles(string sourceTag, string expected, string unexpected) {
        string source = "TY  - JOUR\nID  - x\nTI  - Before\n" + sourceTag + "  - 20\nER  -";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains(expected, written.Content, StringComparison.Ordinal);
        Assert.DoesNotContain(unexpected, written.Content, StringComparison.Ordinal);
        Assert.Equal("20", Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items).Pages);
    }
}
