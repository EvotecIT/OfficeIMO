namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyFormatBoundaryTests {
    [Fact]
    public void Generic_CSL_article_remains_distinct_after_an_edit() {
        const string source = "[{\"id\":\"x\",\"type\":\"article\",\"title\":\"Original\"}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;

        Assert.Equal(BibliographyItemType.Article, Assert.Single(document.Items).Type);
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        Assert.Contains("\"type\": \"article\"", written.Content, StringComparison.Ordinal);
        Assert.Equal(BibliographyItemType.Article, reopened.Type);
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Generic_CSL_article_blocks_strict_narrower_type_conversion(BibliographyFormat format) {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Article, Title = "Generic article" });

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Format = format, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV200" && diagnostic.Field == "type");
    }

    [Fact]
    public void RIS_colon_bearing_identifier_scheme_blocks_strict_output() {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Identifiers" };
        item.Identifiers.Add(new BibliographyIdentifier("info:doi", "10.1/x"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV228" && diagnostic.Field == "identifiers.info:doi");
    }

    [Fact]
    public void RIS_end_page_before_start_page_reopens_as_one_range() {
        const string source = "TY  - BOOK\nID  - x\nEP  - 20\nSP  - 10\nTI  - Original\nER  -\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;

        Assert.Equal("10-20", Assert.Single(document.Items).Pages);
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.Equal("10-20", reopened.Pages);
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Structured_name_particles_block_strict_flattening_formats(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = format == BibliographyFormat.Nbib ? "1" : "x", Type = format == BibliographyFormat.Nbib ? BibliographyItemType.ArticleJournal : BibliographyItemType.Book, Title = "Names" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Given = "Ludwig", Family = "Beethoven", NonDroppingParticle = "van" }));
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV229" && diagnostic.Field == "contributors");
    }

    [Fact]
    public void Classic_BibTeX_omits_and_diagnoses_accessed_dates() {
        var document = new BibliographyDocument(BibliographyFormat.BibTex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Dates" };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Accessed, Year = 2026, Month = 8, Day = 29 });
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.DoesNotContain("urldate", written.Content, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV202" && diagnostic.Field == "dates.Accessed");
    }

    [Fact]
    public async Task Auto_detecting_entry_points_enforce_character_limits_before_detection() {
        const string source = "TY  - BOOK\nID  - x\nTI  - Bounded\nER  -\n";
        var options = new BibliographyReadOptions { MaximumInputCharacters = 8, MaximumInputBytes = 1024 };

        Assert.Throws<InvalidDataException>(() => BibliographyDocument.Parse(source, options));

        string path = Path.Combine(Path.GetTempPath(), "officeimo-bibliography-" + Guid.NewGuid().ToString("N") + ".unknown");
        File.WriteAllText(path, source);
        try {
            Assert.Throws<InvalidDataException>(() => BibliographyDocument.Load(path, options: options));
            await Assert.ThrowsAsync<InvalidDataException>(() => BibliographyDocument.LoadAsync(path, options: options));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("pages")]
    [InlineData("url")]
    public void Foreign_EndNote_elements_that_match_known_local_names_remain_native(string fieldName) {
        string source = "<xml xmlns=\"urn:endnote\" xmlns:ext=\"urn:extension\"><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><ext:" + fieldName + ">extension-value</ext:" + fieldName + "></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        BibliographyItem item = Assert.Single(document.Items);

        Assert.True(string.IsNullOrEmpty(fieldName == "pages" ? item.Pages : item.Url));
        BibliographyNativeField native = Assert.Single(item.NativeFields, field => field.Name == fieldName);
        Assert.Contains("urn:extension", native.RawValue, StringComparison.Ordinal);
        item.Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Contains("urn:extension", written.Content, StringComparison.Ordinal);
        Assert.True(string.IsNullOrEmpty(fieldName == "pages" ? reopened.Pages : reopened.Url));
        Assert.Contains(reopened.NativeFields, field => field.Name == fieldName && field.RawValue != null && field.RawValue.Contains("urn:extension", StringComparison.Ordinal));
    }
}
