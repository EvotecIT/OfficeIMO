namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave38RegressionTests {
    [Theory]
    [InlineData(" alpha ")]
    [InlineData("alpha ")]
    [InlineData("\talpha\t")]
    public void Bib_keywords_preserve_surrounding_whitespace(string keyword) {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Keywords" };
        item.Keywords.Add(keyword);
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibLatex).Document.Items);

        Assert.Equal(keyword, Assert.Single(reopened.Keywords));
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    public void Tagged_native_extension_names_preserve_their_casing(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "123", Type = BibliographyItemType.Book, Title = "Native tags" };
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "123"));
        item.NativeFields.Add(new BibliographyNativeField(format, "zZ", "value"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyNativeField reopened = Assert.Single(Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items).NativeFields, field => field.Value == "value");

        Assert.Equal("zZ", reopened.Name);
    }

    [Theory]
    [InlineData("item")]
    [InlineData("name")]
    [InlineData("date")]
    public void Inconsistent_programmatic_CSL_raw_values_are_diagnosed_for_every_native_owner(string owner) {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Raw values" };
        var field = new BibliographyNativeField(BibliographyFormat.CslJson, "x-native", "expected", "123");
        if (owner == "item") item.NativeFields.Add(field);
        else if (owner == "name") {
            var name = new BibliographyName { Family = "Smith" };
            name.NativeFields.Add(field);
            item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, name));
        } else {
            var date = new BibliographyDate { Role = BibliographyDateRole.Issued, Year = 2024 };
            date.NativeFields.Add(field);
            item.Dates.Add(date);
        }
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV247");
    }

    [Theory]
    [InlineData("<custom>123</custom>")]
    [InlineData("<other>expected</other>")]
    public void Inconsistent_programmatic_EndNote_raw_values_or_names_are_diagnosed(string rawValue) {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Raw values" };
        item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, "custom", "expected", rawValue));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV247" && diagnostic.Field == "native.custom");
    }

    [Fact]
    public void Qualified_RIS_identifiers_preserve_leading_value_whitespace_without_a_false_loss() {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Identifiers" };
        item.Identifiers.Add(new BibliographyIdentifier("arXiv", " 123"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyIdentifier reopened = Assert.Single(Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items).Identifiers);

        Assert.Equal("arXiv", reopened.Scheme);
        Assert.Equal(" 123", reopened.Value);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV239");
    }

    [Fact]
    public void Unqualified_RIS_identifier_tags_still_diagnose_leading_whitespace() {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Identifiers" };
        item.Identifiers.Add(new BibliographyIdentifier("DOI", " 123"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV239" && diagnostic.Field == "identifiers.DOI");
    }
}
