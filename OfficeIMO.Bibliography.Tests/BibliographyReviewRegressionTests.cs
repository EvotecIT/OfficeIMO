using System.Globalization;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewRegressionTests {
    [Fact]
    public void Preserve_fingerprint_distinguishes_collection_boundaries() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"x\",\"type\":\"book\",\"keyword\":\"alpha\"}]", BibliographyFormat.CslJson).Document;
        document.Items[0].Keywords.Add("beta");
        string moved = document.Items[0].Keywords[1];
        document.Items[0].Keywords.RemoveAt(1);
        document.Items[0].Notes.Add(moved);

        BibliographyWriteResult written = document.Write();

        Assert.True(document.IsModified);
        Assert.False(written.UsedOriginalSource);
        Assert.Contains("note", written.Content, StringComparison.Ordinal);
    }

    [Fact]
    public void Bib_diagnostics_are_bounded_for_hostile_input() {
        string source = string.Concat(Enumerable.Repeat("outside@?", 100));
        var options = new BibliographyReadOptions { MaximumDiagnosticCount = 3 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibTex, options);

        Assert.True(read.HasErrors);
        Assert.True(read.Diagnostics.Count <= 4);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM002");
    }

    [Fact]
    public void Unrepresentable_writer_encoding_is_reported_and_strictly_rejected() {
        BibliographyDocument document = BibliographyDocument.Parse("@book{x,title={Łódź}}", BibliographyFormat.BibLatex).Document;
        var permissive = new BibliographyWriteOptions { Encoding = Encoding.ASCII };

        BibliographyWriteResult written = document.Write(permissive);
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Encoding = Encoding.ASCII, RequireNoLoss = true }));

        Assert.Contains(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV220");
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV220");
    }

    [Fact]
    public void Tagged_repeated_authors_and_title_continuations_are_retained() {
        BibliographyItem nbib = Assert.Single(BibliographyDocument.Parse("PMID- 1\nAU  - Doe J\nAU  - Roe R\nTI  - Title\n", BibliographyFormat.Nbib).Document.Items);
        BibliographyItem ris = Assert.Single(BibliographyDocument.Parse("TY  - BOOK\nID  - x\nTI  - First\n      AI-based second\nER  -\n", BibliographyFormat.Ris).Document.Items);

        Assert.Equal(2, nbib.Contributors.Count);
        Assert.Equal("First AI-based second", ris.Title);
        BibliographyDocument risDocument = BibliographyDocument.Parse("TY  - BOOK\nID  - x\nTI  - First\n      AI-based second\nER  -\n", BibliographyFormat.Ris).Document;
        risDocument.Items[0].Issue = "2";
        Assert.Equal("First AI-based second", BibliographyDocument.Parse(risDocument.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }).Content, BibliographyFormat.Ris).Document.Items[0].Title);
    }

    [Fact]
    public void Tagged_continuations_observe_value_length_limit() {
        var options = new BibliographyReadOptions { MaximumValueLength = 5 };

        BibliographyReadResult read = BibliographyDocument.Parse("TY  - BOOK\nTI  - A\n      123456\nER  -\n", BibliographyFormat.Ris, options);

        Assert.True(read.HasErrors);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Fact]
    public void Ris_serial_identifier_schemes_survive_hyphenated_values() {
        const string source = "TY  - BOOK\nID  - x\nSN  - 978-1-4028-9462-6\nSN  - 1234-567X\nER  -\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.Contains(reopened.Identifiers, identifier => identifier.Scheme == "ISBN" && identifier.Value == "978-1-4028-9462-6");
        Assert.Contains(reopened.Identifiers, identifier => identifier.Scheme == "ISSN" && identifier.Value == "1234-567X");
    }

    [Fact]
    public void EndNote_identifier_schemes_survive_strict_round_trip() {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Identifiers" };
        item.Identifiers.Add(new BibliographyIdentifier("ISBN", "978-1-4028-9462-6"));
        item.Identifiers.Add(new BibliographyIdentifier("ISSN", "1234-567X"));
        item.Identifiers.Add(new BibliographyIdentifier("accession", "PMID:archive-7"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal("978-1-4028-9462-6", reopened.GetIdentifier("ISBN"));
        Assert.Equal("1234-567X", reopened.GetIdentifier("ISSN"));
        Assert.Equal("PMID:archive-7", reopened.GetIdentifier("accession"));
    }

    [Fact]
    public void EndNote_strict_write_rejects_untyped_pmid_scheme() {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "PMID" };
        item.Identifiers.Add(new BibliographyIdentifier("PMID", "12345678"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV204" && diagnostic.Field == "identifiers.PMID");
    }

    [Fact]
    public void Partial_text_date_does_not_invent_day_precision() {
        BibliographyDocument document = BibliographyDocument.Parse("TY  - JOUR\nID  - x\nPY  - 2024 May\nER  -\n", BibliographyFormat.Ris).Document;
        BibliographyDate date = Assert.Single(document.Items[0].Dates);
        document.Items[0].Title = "Edited";

        BibliographyDate reopened = Assert.Single(BibliographyDocument.Parse(document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }).Content, BibliographyFormat.Ris).Document.Items[0].Dates);

        Assert.Equal(2024, date.Year); Assert.Equal(5, date.Month); Assert.Null(date.Day);
        Assert.Equal(2024, reopened.Year); Assert.Equal(5, reopened.Month); Assert.Null(reopened.Day);
    }

    [Fact]
    public void Named_bib_month_survives_strict_canonical_write() {
        BibliographyDocument document = BibliographyDocument.Parse("@book{x,title={Month},year=2024,month=jan}", BibliographyFormat.BibTex).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate date = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibTex).Document.Items[0].Dates);

        Assert.Equal(1, date.Month);
    }

    [Fact]
    public void Supplemental_literal_date_blocks_strict_destinations_that_drop_it() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Qualified date" };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Year = 2024, Month = 5, Literal = "circa May 2024" });
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.Ris, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV221");
    }

    [Fact]
    public void Unsupported_nested_EndNote_content_remains_raw_and_blocks_strict_edit() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Main</title><short-title>Short</short-title></titles></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "Edited";

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(document.Items[0].NativeFields, field => field.Name == "titles" && field.RawValue!.Contains("short-title", StringComparison.Ordinal));
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV123" && diagnostic.Field == "titles");
    }

    [Fact]
    public void EndNote_preserves_valid_supplementary_unicode() {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Emoji 😀" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Equal("Emoji 😀", Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items).Title);
    }

    [Fact]
    public void Deep_native_csl_json_remains_structured() {
        string nested = "true";
        for (int index = 0; index < 80; index++) nested = "{\"level\":" + nested + "}";
        string source = "[{\"id\":\"deep\",\"type\":\"book\",\"x-deep\":" + nested + "}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, new BibliographyReadOptions { MaximumNestingDepth = 128 }).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyNativeField field = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson, new BibliographyReadOptions { MaximumNestingDepth = 128 }).Document.Items[0].NativeFields);

        Assert.StartsWith("{", field.RawValue, StringComparison.Ordinal);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV126");
    }

    [Fact]
    public void Detection_accepts_a_single_CSL_object_without_a_type_property() {
        BibliographyReadResult read = BibliographyDocument.Parse("{\"id\":\"item\",\"title\":\"Example\"}");

        BibliographyItem item = Assert.Single(read.Document.Items);
        Assert.Equal(BibliographyFormat.CslJson, read.Document.SourceFormat);
        Assert.Equal("item", item.Key);
        Assert.Equal(BibliographyItemType.Unknown, item.Type);
    }

    [Fact]
    public void Secondary_NBIB_publication_types_survive_strict_canonical_writes() {
        const string source = "PMID- 1\nPT  - Journal Article\nPT  - Book\nTI  - Multiple types\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Nbib).Document.Items);

        Assert.Equal(new[] { "Journal Article", "Book" }, reopened.NativeFields.Where(field => field.Name == "PT").Select(field => field.Value));
    }

    [Fact]
    public void RIS_dates_preserve_their_cross_role_source_order() {
        const string source = "TY  - BOOK\nID  - x\nY2  - 2025-02-03\nPY  - 2024-01-02\nER  -\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.True(written.Content.IndexOf("Y2  -", StringComparison.Ordinal) < written.Content.IndexOf("PY  -", StringComparison.Ordinal));
        Assert.Equal(new[] { BibliographyDateRole.Accessed, BibliographyDateRole.Issued }, reopened.Dates.Select(static date => date.Role));
    }

    [Fact]
    public void Edited_structured_CSL_native_JSON_preserves_exact_formatting_when_valid() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"x\",\"type\":\"book\",\"custom\":{\"old\":true}}]", BibliographyFormat.CslJson).Document;
        Assert.Single(document.Items[0].NativeFields).Value = "{\"edited\":true}";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyNativeField reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items[0].NativeFields);

        Assert.StartsWith("{", reopened.RawValue, StringComparison.Ordinal);
        Assert.Equal("{\"edited\":true}", reopened.RawValue);
        Assert.Equal("{\"edited\":true}", reopened.Value);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV126");
    }

    [Fact]
    public void Edited_structured_CSL_native_JSON_reports_shape_flattening() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"x\",\"type\":\"book\",\"custom\":{\"old\":true}}]", BibliographyFormat.CslJson).Document;
        Assert.Single(document.Items[0].NativeFields).Value = "flattened";

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV126" && diagnostic.Field == "custom");
    }

    [Fact]
    public void Brace_wrapped_Bib_keywords_survive_strict_canonical_writes() {
        BibliographyDocument document = BibliographyDocument.Parse("@book{x,title={Keyword},keywords={{{tag}}}}", BibliographyFormat.BibLatex).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibLatex).Document.Items);

        Assert.Equal("{tag}", Assert.Single(reopened.Keywords));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Preserve_fingerprint_distinguishes_null_and_empty_scalars(bool removeEmptyTitle) {
        string source = removeEmptyTitle ? "TY  - BOOK\nID  - x\nTI  - \nER  -\n" : "TY  - BOOK\nID  - x\nER  -\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;
        document.Items[0].Title = removeEmptyTitle ? null : string.Empty;

        BibliographyWriteResult written = document.Write();

        Assert.True(document.IsModified);
        Assert.False(written.UsedOriginalSource);
        Assert.Equal(!removeEmptyTitle, written.Content.Contains("TI  - ", StringComparison.Ordinal));
    }

    [Fact]
    public void Generic_EndNote_documents_round_trip_in_strict_mode() {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Document, Title = "Generic" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Equal(BibliographyItemType.Document, Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items).Type);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void Blank_CSL_scalars_round_trip_in_strict_mode(string title) {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = title });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Equal(title, Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items).Title);
    }

    [Fact]
    public void Null_CSL_scalars_remain_native_null_values() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"x\",\"type\":\"book\",\"title\":null,\"author\":[{\"family\":null}],\"issued\":{\"literal\":null}}]", BibliographyFormat.CslJson).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        Assert.Null(reopened.Title);
        Assert.Contains(reopened.NativeFields, field => field.Name == "title" && field.RawValue == "null");
        Assert.Contains(reopened.NativeFields, field => field.Name == "author" && field.RawValue!.StartsWith("[", StringComparison.Ordinal));
        Assert.Contains(Assert.Single(reopened.Dates).NativeFields, field => field.Name == "literal" && field.RawValue == "null");
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Tagged_name_components_with_commas_block_strict_output(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Names" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Doe, Smith", Given = "Jane" }));
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV236" && diagnostic.Field == "contributors");
    }

    [Fact]
    public void CSL_writes_observe_cancellation_within_one_large_item() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book };
        for (int index = 0; index < 100_000; index++) item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.CslJson, "field-" + index.ToString(CultureInfo.InvariantCulture), "true", "true"));
        document.Items.Add(item);
        using var cancellation = new CancellationTokenSource();
        cancellation.CancelAfter(TimeSpan.FromMilliseconds(10));

        Assert.Throws<OperationCanceledException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, cancellation.Token));
    }

    [Fact]
    public void EndNote_item_limits_are_enforced_before_the_XML_DOM_is_materialized() {
        const string source = "<xml><records><record/><record/>";
        var options = new BibliographyReadOptions { MaximumItemCount = 1 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml, options);

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Code == "BIBEND002");
    }

    [Fact]
    public void EndNote_item_limits_ignore_records_nested_in_root_extensions() {
        const string source = "<xml><extension><records><record/><record/></records></extension><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type></record></records></xml>";
        var options = new BibliographyReadOptions { MaximumItemCount = 1 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml, options);

        Assert.False(read.HasErrors);
        Assert.Single(read.Document.Items);
    }

    [Fact]
    public void Clearing_the_primary_EndNote_URL_preserves_additional_URL_roles() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><urls><related-urls><url>https://primary.example</url><url>https://secondary.example</url></related-urls></urls></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Url = null;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Null(reopened.Url);
        Assert.Contains(reopened.NativeFields, field => field.Name == "url" && field.Value == "https://secondary.example");
    }

    [Fact]
    public void Empty_EndNote_URLs_reopen_as_empty_values() {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Url = string.Empty });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal(string.Empty, reopened.Url);
    }

    [Fact]
    public void Separate_EndNote_records_container_attributes_block_strict_coalescing() {
        const string source = "<xml><records source=\"first\"><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type></record></records><records source=\"second\"><record><rec-number>2</rec-number><ref-type name=\"Book\">6</ref-type></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV238" && diagnostic.Field == "records");
    }

    [Fact]
    public void CSL_item_limits_are_enforced_before_the_JSON_DOM_is_materialized() {
        const string source = "[{},{}";
        var options = new BibliographyReadOptions { MaximumItemCount = 1 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, options);

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Code == "BIBCSL002");
    }

    [Fact]
    public void CSL_value_limits_are_enforced_before_the_JSON_DOM_is_materialized() {
        const string source = "[{\"a\":1,\"b\":2}";
        var options = new BibliographyReadOptions { MaximumValueCount = 1 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, options);

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Code == "BIBCSL002");
    }

    [Fact]
    public void CSL_decoded_value_lengths_are_enforced_before_materialization() {
        const string source = "[{\"title\":\"abcdef\"}";
        var options = new BibliographyReadOptions { MaximumValueLength = 5 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, options);

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Code == "BIBCSL002");
    }

    [Fact]
    public void CSL_decoded_value_lengths_do_not_count_escape_syntax() {
        const string source = "[{\"title\":\"\\u0061\"}]";
        var options = new BibliographyReadOptions { MaximumValueLength = 5 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, options);

        Assert.False(read.HasErrors);
        Assert.Equal("a", Assert.Single(read.Document.Items).Title);
    }

    [Fact]
    public void CSL_root_array_items_are_not_double_counted_as_values() {
        const string source = "[{},{}]";
        var options = new BibliographyReadOptions { MaximumItemCount = 2, MaximumValueCount = 1 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, options);

        Assert.False(read.HasErrors);
        Assert.Equal(2, read.Document.Items.Count);
    }

    [Fact]
    public void EndNote_value_limits_are_enforced_before_the_XML_DOM_is_materialized() {
        const string source = "<xml><records><record><keywords><keyword>a</keyword><keyword>b</keyword>";
        var options = new BibliographyReadOptions { MaximumValueCount = 1 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml, options);

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Code == "BIBEND002");
    }

    [Fact]
    public void EndNote_attribute_limits_are_enforced_before_the_XML_DOM_is_materialized() {
        const string source = "<xml first=\"1\" second=\"2\"><records>";
        var options = new BibliographyReadOptions { MaximumValueCount = 1 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml, options);

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Code == "BIBEND002");
    }

    [Fact]
    public void EndNote_value_lengths_are_enforced_before_the_XML_DOM_is_materialized() {
        const string source = "<xml><records><record><title>abcdef</title>";
        var options = new BibliographyReadOptions { MaximumValueLength = 5 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml, options);

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Code == "BIBEND002");
    }

    [Fact]
    public void Absent_EndNote_titles_remain_null_after_strict_reopen() {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = null });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Null(reopened.Title);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void Blank_EndNote_titles_remain_distinct_from_absent_titles(string title) {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = title });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal(title, reopened.Title);
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    public void Leading_whitespace_in_tagged_values_blocks_strict_output(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = " Leading" };
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV239" && diagnostic.Field == "title");
    }

    [Fact]
    public void Leading_whitespace_in_native_tagged_values_blocks_strict_output() {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Native" };
        item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Ris, "C1", " retained"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV239" && diagnostic.Field == "native.C1");
    }

    [Fact]
    public void Detection_accepts_namespace_prefixed_EndNote_roots() {
        const string source = "<e:xml xmlns:e=\"urn:endnote\"><e:records><e:record><e:rec-number>1</e:rec-number><e:ref-type name=\"Book\">6</e:ref-type></e:record></e:records></e:xml>";

        BibliographyReadResult read = BibliographyDocument.Parse(source);

        Assert.Equal(BibliographyFormat.EndNoteXml, read.Document.SourceFormat);
        Assert.Equal("1", Assert.Single(read.Document.Items).Key);
    }

    [Fact]
    public void Detection_accepts_namespace_prefixed_EndNote_records_roots() {
        const string source = "<e:records xmlns:e=\"urn:endnote\"><e:record><e:rec-number>1</e:rec-number><e:ref-type name=\"Book\">6</e:ref-type></e:record></e:records>";

        BibliographyReadResult read = BibliographyDocument.Parse(source);

        Assert.Equal(BibliographyFormat.EndNoteXml, read.Document.SourceFormat);
        Assert.Equal("1", Assert.Single(read.Document.Items).Key);
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Literal_names_with_commas_round_trip_in_tagged_destinations(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Organization" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Literal = "Acme, Inc." }));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyName reopened = Assert.Single(Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items).Contributors).Name;

        Assert.Equal("Acme, Inc.", reopened.Literal);
        Assert.Null(reopened.Family);
    }

    [Theory]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Parseable_literal_dates_block_strict_output_that_promotes_them(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Literal date" };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Literal = "2025-01-02" });
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV240" && diagnostic.Field == "dates.Issued.literal");
    }

    [Theory]
    [InlineData(99, null)]
    [InlineData((int)BibliographyItemType.Unknown, "JOUR")]
    public void RIS_types_that_reopen_with_different_semantics_block_strict_output(int type, string? nativeType) {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        document.Items.Add(new BibliographyItem { Key = "x", Type = (BibliographyItemType)type, NativeType = nativeType, Title = "Type" });

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV200" && diagnostic.Field == "type");
    }

    [Fact]
    public void Unknown_native_RIS_types_that_remain_unknown_reopen_exactly() {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Unknown, NativeType = "CUSTOM", Title = "Type" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.Equal(BibliographyItemType.Unknown, reopened.Type);
        Assert.Equal("CUSTOM", reopened.NativeType);
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex, "")]
    [InlineData(BibliographyFormat.BibTex, " ")]
    [InlineData(BibliographyFormat.BibLatex, "")]
    [InlineData(BibliographyFormat.BibLatex, " ")]
    public void Blank_Bib_scalars_remain_distinct_from_absent_values(BibliographyFormat format, string title) {
        var document = new BibliographyDocument(format);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = title });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items);

        Assert.Equal(title, reopened.Title);
    }

    [Fact]
    public void CSL_aggregate_value_lengths_are_enforced_before_the_JSON_DOM_is_materialized() {
        const string source = "[{\"id\":\"x\",\"custom\":{\"a\":1,\"b\":2}";
        var options = new BibliographyReadOptions { MaximumValueLength = 10 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, options);

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Code == "BIBCSL002");
    }

    [Fact]
    public void CSL_aggregate_value_lengths_use_UTF16_coordinates() {
        const string source = "[{\"id\":\"x\",\"custom\":{\"text\":\"😀\"}}]";

        BibliographyReadResult accepted = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, new BibliographyReadOptions { MaximumValueLength = 13 });
        BibliographyReadResult rejected = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, new BibliographyReadOptions { MaximumValueLength = 12 });

        Assert.False(accepted.HasErrors);
        Assert.Contains(rejected.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Fact]
    public void Structured_Bib_names_with_contributor_separators_block_strict_output() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Names" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Smith and Jones", Given = "Jane" }));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV226" && diagnostic.Field == "contributors");
    }

    [Fact]
    public void Bib_raw_directives_observe_cancellation_during_scan() {
        string source = "@comment{" + new string('x', 8 * 1024 * 1024);
        var options = new BibliographyReadOptions { MaximumValueLength = 16 * 1024 * 1024 };
        using var cancellation = new CancellationTokenSource();
        cancellation.CancelAfter(TimeSpan.FromMilliseconds(10));

        Assert.Throws<OperationCanceledException>(() =>
            BibliographyDocument.Parse(source, BibliographyFormat.BibLatex, options, cancellation.Token));
    }

    [Fact]
    public void Empty_RIS_pages_remain_distinct_from_absent_pages() {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Pages", Pages = string.Empty });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.Equal(string.Empty, reopened.Pages);
    }

    [Theory]
    [InlineData("null")]
    [InlineData("{\"legacy\":1}")]
    public void Native_CSL_identifiers_do_not_trigger_generated_key_loss(string nativeId) {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":" + nativeId + ",\"type\":\"book\",\"title\":\"Native id\"}]", BibliographyFormat.CslJson).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        Assert.True(string.IsNullOrWhiteSpace(reopened.Key));
        BibliographyNativeField field = Assert.Single(reopened.NativeFields, field => field.Format == BibliographyFormat.CslJson && field.Name == "id");
        if (nativeId == "null") Assert.Equal("null", field.RawValue);
        else Assert.StartsWith("{", field.RawValue, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("ty  - jour", BibliographyFormat.Ris)]
    [InlineData("PmId- 123", BibliographyFormat.Nbib)]
    [InlineData("own - medline", BibliographyFormat.Nbib)]
    public void Detection_accepts_tag_casing_supported_by_tagged_parsers(string source, BibliographyFormat expected) {
        Assert.Equal(expected, BibliographyDocument.Parse(source).Document.SourceFormat);
    }

    [Theory]
    [InlineData("titles")]
    [InlineData("contributors")]
    [InlineData("dates")]
    [InlineData("keywords")]
    [InlineData("urls")]
    public void EndNote_containers_retain_unsupported_direct_text(string container) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><" + container + ">loose text</" + container + "></record></records></xml>";

        BibliographyItem item = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document.Items);

        BibliographyNativeField field = Assert.Single(item.NativeFields, field => field.Name == container);
        Assert.Contains("loose text", field.RawValue, StringComparison.Ordinal);
    }

    [Fact]
    public void EndNote_containers_retain_unhandled_child_nodes() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><!--native note--><title>Typed</title></titles></record></records></xml>";

        BibliographyItem item = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document.Items);

        BibliographyNativeField field = Assert.Single(item.NativeFields, field => field.Name == "titles");
        Assert.Contains("<!--native note-->", field.RawValue, StringComparison.Ordinal);
    }
}
