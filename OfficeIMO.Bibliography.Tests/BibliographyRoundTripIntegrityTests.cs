namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyRoundTripIntegrityTests {
    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Literal_contributors_reopen_as_literal_names(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem {
            Key = format == BibliographyFormat.Nbib ? "12345" : "x",
            Type = format == BibliographyFormat.Nbib ? BibliographyItemType.ArticleJournal : BibliographyItemType.Book,
            Title = "Corporate author"
        };
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "12345"));
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Literal = "World Health Organization" }));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyContributor reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items[0].Contributors);

        Assert.Equal("World Health Organization", reopened.Name.Literal);
        Assert.Null(reopened.Name.Given);
        Assert.Null(reopened.Name.Family);
    }

    [Theory]
    [InlineData("TI", "T1")]
    [InlineData("T2", "JF")]
    [InlineData("AB", "N2")]
    [InlineData("UR", "L1")]
    [InlineData("PB", "PB")]
    public void Repeated_RIS_scalar_values_survive_canonical_reopen(string firstTag, string secondTag) {
        string source = $"TY  - BOOK\nID  - x\n{firstTag}  - First\n{secondTag}  - Second\nER  -\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;

        BibliographyWriteResult first = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDocument reopened = BibliographyDocument.Parse(first.Content, BibliographyFormat.Ris).Document;
        BibliographyWriteResult second = reopened.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains(reopened.Items[0].NativeFields, field => field.Name == secondTag && field.Value == "Second");
        Assert.Equal(first.Content, second.Content);
    }

    [Fact]
    public void Repeated_RIS_scalar_continuation_stays_with_the_repeated_value() {
        const string source = "TY  - BOOK\nID  - x\nUR  - https://first.example\nL1  - https://second.example/path\n      continued\nER  -\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.Equal("https://first.example", reopened.Url);
        Assert.Equal("https://second.example/path continued", Assert.Single(reopened.NativeFields, field => field.Name == "L1").Value);
    }

    [Fact]
    public void Repeated_NBIB_scalar_values_survive_canonical_reopen() {
        const string source = "PMID- 12345\nTI  - First\nTI  - Second\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Nbib).Document.Items);

        Assert.Equal("First", reopened.Title);
        Assert.Equal("Second", Assert.Single(reopened.NativeFields, field => field.Name == "TI").Value);
    }

    [Fact]
    public void Unsupported_Bib_identifier_scheme_is_omitted_and_blocks_strict_output() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Identifiers" };
        item.Identifiers.Add(new BibliographyIdentifier("arXiv", "2401.00001"));
        document.Items.Add(item);

        BibliographyWriteResult canonical = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.DoesNotContain("arxiv =", canonical.Content, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV129" && diagnostic.Field == "identifiers.arXiv");
    }

    [Fact]
    public void Blank_line_ends_a_PMID_less_NBIB_record() {
        const string source = "DP  - 2024\nVI  - 1\n\nTI  - Second record\n";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.Nbib);

        Assert.Equal(2, read.Document.Items.Count);
        Assert.Equal("1", read.Document.Items[0].Volume);
        Assert.Equal("Second record", read.Document.Items[1].Title);
    }

    [Theory]
    [InlineData("ISBN")]
    [InlineData("ISSN")]
    public void Ambiguous_EndNote_serial_identifier_reopens_with_its_declared_scheme(string scheme) {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Serial" };
        item.Identifiers.Add(new BibliographyIdentifier(scheme, "catalogue-value (print)"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyIdentifier reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items[0].Identifiers);

        Assert.Contains($"type=\"{scheme}\"", written.Content, StringComparison.Ordinal);
        Assert.Equal(scheme, reopened.Scheme);
        Assert.Equal("catalogue-value (print)", reopened.Value);

        BibliographyDocument reopenedDocument = BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document;
        BibliographyWriteResult second = reopenedDocument.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        Assert.Equal(written.Content, second.Content);
    }

    [Fact]
    public void Qualified_EndNote_serial_identifier_is_inferred_from_the_identifier_not_the_qualifier() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><isbn>978-1-4028-9462-6 (paperback)</isbn></record></records></xml>";

        BibliographyIdentifier identifier = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document.Items[0].Identifiers);

        Assert.Equal("ISBN", identifier.Scheme);
    }

    [Fact]
    public void RIS_type_tag_observes_the_value_length_limit() {
        BibliographyReadResult read = BibliographyDocument.Parse("TY  - BOOK\nER  -\n", BibliographyFormat.Ris, new BibliographyReadOptions { MaximumValueLength = 3 });

        Assert.True(read.HasErrors);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Theory]
    [InlineData("[{\"id\":\"x\",\"type\":\"book\",\"extension\":{\"a\":\"1234567890\",\"b\":\"1234567890\"}}]")]
    [InlineData("[{\"id\":\"x\",\"type\":\"book\",\"author\":[{\"family\":\"Doe\",\"extension\":{\"a\":\"1234567890\",\"b\":\"1234567890\"}}]}]")]
    [InlineData("[{\"id\":\"x\",\"type\":\"book\",\"issued\":{\"literal\":\"2024\",\"extension\":{\"a\":\"1234567890\",\"b\":\"1234567890\"}}}]")]
    public void Aggregate_native_CSL_values_observe_the_value_length_limit(string source) {
        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, new BibliographyReadOptions { MaximumValueLength = 30 });

        Assert.True(read.HasErrors);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Fact]
    public void Malformed_JSON_location_handles_non_ASCII_without_per_character_substrings() {
        BibliographyReadResult read = BibliographyDocument.Parse("[{\"id\":\"😀\" \"type\":\"book\"}]", BibliographyFormat.CslJson);

        BibliographyDiagnostic diagnostic = Assert.Single(read.Diagnostics, value => value.Code == "BIBCSL002");
        Assert.Equal(1, diagnostic.Line);
        Assert.True(diagnostic.Column > 10);
        Assert.True(diagnostic.Offset > 10);
    }

    [Theory]
    [InlineData("% Bib comment\n@book{x,title={x}}", BibliographyFormat.BibLatex)]
    [InlineData("// JSON comment\n[{\"id\":\"x\",\"type\":\"book\"}]", BibliographyFormat.CslJson)]
    [InlineData("/* JSON comment */ [{\"id\":\"x\",\"type\":\"book\"}]", BibliographyFormat.CslJson)]
    public void Format_detection_skips_supported_leading_comments(string source, BibliographyFormat expected) {
        Assert.Equal(expected, BibliographyDocument.Parse(source).Document.SourceFormat);
    }

    [Theory]
    [InlineData("% Bib-only comment\n[{\"id\":\"x\",\"type\":\"book\"}]")]
    [InlineData("// JSON-only comment\n@book{x,title={x}}")]
    public void Format_detection_does_not_apply_one_formats_comments_to_another(string source) {
        Assert.Throws<FormatException>(() => BibliographyDocument.Parse(source));
    }

    [Theory]
    [InlineData("string", "unsafe}name", "value")]
    [InlineData("comment", null, "unsafe}")]
    [InlineData("line-comment", null, "first\n@book{injected,title={x}}")]
    public void Unsafe_Bib_document_directives_are_omitted_and_block_strict_output(string kind, string? name, string value) {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.BibLatex, kind, value, name));

        BibliographyWriteResult canonical = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.DoesNotContain("injected", canonical.Content, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV118");
    }

    [Fact]
    public void EndNote_record_discovery_ignores_records_nested_in_root_extensions() {
        const string source = "<xml><extension><record><rec-number>hidden</rec-number><titles><title>Hidden</title></titles></record></extension><records><record><rec-number>visible</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Visible</title></titles></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;

        BibliographyItem item = Assert.Single(document.Items);
        Assert.Equal("visible", item.Key);
        Assert.Single(document.NativeEntries);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDocument reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document;
        Assert.Equal("visible", Assert.Single(reopened.Items).Key);
    }

    [Fact]
    public void EndNote_records_root_parses_its_direct_records() {
        const string source = "<records><record><rec-number>visible</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Visible</title></titles></record></records>";

        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;

        Assert.Equal("visible", Assert.Single(document.Items).Key);
        Assert.Empty(document.NativeEntries);
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris, "TY  - BOOK\nID  - x\nAU  - World Health,\n      Organization\nER  -\n")]
    [InlineData(BibliographyFormat.Nbib, "PMID- 12345\nCN  - World Health\n      Organization\n")]
    public void Tagged_literal_contributor_continuations_remain_literal(BibliographyFormat format, string source) {
        BibliographyName name = Assert.Single(BibliographyDocument.Parse(source, format).Document.Items[0].Contributors).Name;

        Assert.Equal("World Health Organization", name.Literal);
        Assert.Null(name.Family);
    }

    [Fact]
    public void Repeated_known_Bib_fields_survive_strict_canonical_reopen() {
        const string source = "@book{x,title={First},title={Second},year={2024},year={2025}}";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex).Document;

        BibliographyWriteResult first = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDocument reopened = BibliographyDocument.Parse(first.Content, BibliographyFormat.BibLatex).Document;
        BibliographyWriteResult second = reopened.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Equal("First", reopened.Items[0].Title);
        Assert.Equal(2, reopened.Items[0].NativeFields.Count);
        Assert.Equal(first.Content, second.Content);
    }

    [Theory]
    [InlineData("[{\"id\":\"x\",\"type\":\"book\",\"title\":\"First\",\"title\":\"Second\"}]")]
    [InlineData("[{\"id\":\"x\",\"type\":\"book\",\"author\":[{\"family\":\"First\",\"family\":\"Second\"}]}]")]
    [InlineData("[{\"id\":\"x\",\"type\":\"book\",\"issued\":{\"literal\":\"First\",\"literal\":\"Second\"}}]")]
    public void Duplicate_known_CSL_properties_are_retained_and_block_lossless_canonical_output(string source) {
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV120" || diagnostic.Code == "BIBCONV124" || diagnostic.Code == "BIBCONV125");
    }

    [Theory]
    [InlineData("<pages>1</pages><pages>2</pages>")]
    [InlineData("<titles><title>First</title><title>Second</title></titles>")]
    public void Duplicate_known_EndNote_values_are_retained_and_block_lossless_canonical_output(string repeatedContent) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type>" + repeatedContent + "</record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV123");
    }
}
