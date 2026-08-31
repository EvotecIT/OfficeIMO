namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave35RegressionTests {
    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Non_CSL_all_null_contributor_names_are_diagnosed_before_strict_output(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName()));
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV244" && diagnostic.Field == "contributors");
    }

    [Fact]
    public void CSL_all_null_contributor_names_reopen_exactly() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName()));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyName reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items[0].Contributors).Name;

        Assert.Null(reopened.Given);
        Assert.Null(reopened.Family);
        Assert.Null(reopened.Literal);
        Assert.Null(reopened.Suffix);
        Assert.Null(reopened.DroppingParticle);
        Assert.Null(reopened.NonDroppingParticle);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CSL_days_without_their_owning_month_are_diagnosed_before_strict_output(bool rangeEnd) {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Dates.Add(rangeEnd
            ? new BibliographyDate { Role = BibliographyDateRole.Issued, Year = 2020, Month = 1, Day = 2, EndYear = 2021, EndDay = 3 }
            : new BibliographyDate { Role = BibliographyDateRole.Issued, Year = 2020, Day = 3 });
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV218" && diagnostic.Field == "dates.Issued");
    }

    [Theory]
    [InlineData(nameof(BibliographyName.DroppingParticle))]
    [InlineData(nameof(BibliographyName.NonDroppingParticle))]
    public void Tagged_whitespace_only_particles_are_diagnosed_before_strict_output(string property) {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        var name = new BibliographyName { Family = "Doe" };
        if (property == nameof(BibliographyName.DroppingParticle)) name.DroppingParticle = " ";
        else name.NonDroppingParticle = " ";
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, name));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV243" && diagnostic.Field == "contributors");
    }
}
