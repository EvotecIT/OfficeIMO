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

    [Fact]
    public void EndNote_retains_distinct_secondary_and_periodical_titles() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Work</title><secondary-title>Secondary</secondary-title></titles><periodical><full-title>Periodical</full-title></periodical></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        BibliographyItem item = Assert.Single(document.Items);

        Assert.Equal("Secondary", item.ContainerTitle);
        Assert.Single(item.NativeFields, field => field.Name == "periodical" && field.Value.Trim() == "Periodical");

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal("Secondary", reopened.ContainerTitle);
        Assert.Single(reopened.NativeFields, field => field.Name == "periodical" && field.Value.Trim() == "Periodical");
    }

    [Fact]
    public void EndNote_preserves_a_periodical_only_container_title() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><periodical><full-title>Periodical</full-title></periodical></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal("Periodical", reopened.ContainerTitle);
        Assert.Contains("<periodical>", written.Content, StringComparison.Ordinal);
        Assert.DoesNotContain("<secondary-title>", written.Content, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("{\"literal\":\"Example Group\"}")]
    [InlineData("[42]")]
    [InlineData("[{\"literal\":{\"value\":\"Example Group\"}}]")]
    public void Wrong_shaped_CSL_contributor_properties_remain_native(string author) {
        string source = "{\"id\":\"x\",\"type\":\"book\",\"author\":" + author + "}";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        BibliographyItem item = Assert.Single(document.Items);

        Assert.Empty(item.Contributors);
        Assert.Single(item.NativeFields, field => field.Name == "author");

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        Assert.Empty(reopened.Contributors);
        Assert.Single(reopened.NativeFields, field => field.Name == "author");
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.CslJson)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Role_grouping_formats_diagnose_cross_role_contributor_reordering(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Names" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "First" }));
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Editor, new BibliographyName { Family = "Editor" }));
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Last" }));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV230" && diagnostic.Field == "contributors");
    }

    [Fact]
    public void NBIB_diagnoses_personal_and_collective_author_reordering() {
        var document = new BibliographyDocument(BibliographyFormat.Nbib);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.ArticleJournal, Title = "Names" };
        item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Literal = "Example Group" }));
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Person", Given = "A" }));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV230" && diagnostic.Field == "contributors");
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    public void Tagged_native_line_normalization_blocks_strict_output(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = format == BibliographyFormat.Nbib ? BibliographyItemType.ArticleJournal : BibliographyItemType.Book, Title = "Native" };
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        item.NativeFields.Add(new BibliographyNativeField(format, "ZZ", "first\nsecond"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV209" && diagnostic.Field == "native.ZZ");
    }

    [Fact]
    public void Bib_identifier_allocations_respect_value_length_limits() {
        string[] sources = {
            "@verylongentrytype{x,title={A}}",
            "@string{averylongmacroname={A}}",
            "@book{x,averylongfieldname={A}}"
        };
        var options = new BibliographyReadOptions { MaximumValueLength = 10 };

        foreach (string source in sources) {
            BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex, options);
            Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
        }
    }

    [Fact]
    public void Serialized_EndNote_extensions_respect_value_length_limits() {
        string markup = string.Concat(Enumerable.Repeat("<x/>", 30));
        string[] sources = {
            "<xml><extension>" + markup + "</extension><records/></xml>",
            "<xml xmlns:ext=\"urn:extension\"><records><record><rec-number>1</rec-number><ext:data>" + markup + "</ext:data></record></records></xml>"
        };
        var options = new BibliographyReadOptions { MaximumValueLength = 80 };

        foreach (string source in sources) {
            BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml, options);
            Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
        }
    }

    [Fact]
    public void Bib_booktitle_and_series_map_to_container_and_collection_titles() {
        const string source = "@incollection{x,title={Part},booktitle={Containing Work},series={Collection}}";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex).Document;
        BibliographyItem item = Assert.Single(document.Items);

        Assert.Equal("Containing Work", item.ContainerTitle);
        Assert.Equal("Collection", item.CollectionTitle);

        BibliographyWriteResult csl = document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.CslJson, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        Assert.Contains("\"container-title\": \"Containing Work\"", csl.Content, StringComparison.Ordinal);
        Assert.Contains("\"collection-title\": \"Collection\"", csl.Content, StringComparison.Ordinal);

        BibliographyWriteResult bib = document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.BibLatex, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        Assert.Contains("booktitle = {Containing Work}", bib.Content, StringComparison.Ordinal);
        Assert.Contains("series = {Collection}", bib.Content, StringComparison.Ordinal);
    }
}
