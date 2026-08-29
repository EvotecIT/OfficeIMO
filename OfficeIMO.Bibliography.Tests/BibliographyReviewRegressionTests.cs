namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewRegressionTests {
    [Fact]
    public void Preserve_fingerprint_distinguishes_collection_boundaries() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"x\",\"type\":\"book\",\"keyword\":\"alpha, beta\"}]", BibliographyFormat.CslJson).Document;
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
}
