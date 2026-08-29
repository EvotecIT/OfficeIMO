namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyIoAndSafetyTests {
    [Fact]
    public async Task Async_stream_round_trip_preserves_utf8_bom_bytes() {
        byte[] preamble = new UTF8Encoding(true).GetPreamble();
        byte[] body = Encoding.UTF8.GetBytes("@book{x,title={BOM source}}");
        byte[] source = preamble.Concat(body).ToArray();
        using var input = new MemoryStream(source);

        BibliographyReadResult read = await BibliographyDocument.LoadAsync(input, BibliographyFormat.BibTex);
        using var output = new MemoryStream();
        BibliographyWriteResult written = await read.Document.SaveAsync(output);

        Assert.Equal(source, written.Bytes);
        Assert.Equal(source, output.ToArray());
    }

    [Fact]
    public void EndNote_xml_declaration_matches_selected_encoding_and_reopens_from_bytes() {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        document.Items.Add(new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Unicode Łódź" });
        var options = new BibliographyWriteOptions { Format = BibliographyFormat.EndNoteXml, Mode = BibliographyWriterMode.Canonical, Encoding = new UnicodeEncoding(false, true, true) };

        BibliographyWriteResult result = document.Write(options);
        BibliographyReadResult reopened = BibliographyDocument.Load(new MemoryStream(result.Bytes), BibliographyFormat.EndNoteXml);

        Assert.Contains("encoding=\"utf-16\"", result.Content, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("Unicode Łódź", Assert.Single(reopened.Document.Items).Title);
    }

    [Fact]
    public void Format_detection_covers_each_family() {
        Assert.Equal(BibliographyFormat.BibLatex, BibliographyDocument.Parse("@book{x,title={x}}").Document.SourceFormat);
        Assert.Equal(BibliographyFormat.CslJson, BibliographyDocument.Parse("[{\"id\":\"x\",\"type\":\"book\"}]").Document.SourceFormat);
        Assert.Equal(BibliographyFormat.Ris, BibliographyDocument.Parse("TY  - BOOK\nER  -").Document.SourceFormat);
        Assert.Equal(BibliographyFormat.Nbib, BibliographyDocument.Parse("PMID- 1\nTI  - x").Document.SourceFormat);
        Assert.Equal(BibliographyFormat.EndNoteXml, BibliographyDocument.Parse("<xml><records /></xml>").Document.SourceFormat);
    }

    [Fact]
    public void Input_character_limit_is_enforced_before_parsing() {
        var options = new BibliographyReadOptions { MaximumInputCharacters = 8 };
        Assert.Throws<InvalidDataException>(() => BibliographyDocument.Parse("@book{x,title={too long}}", BibliographyFormat.BibTex, options));
    }

    [Fact]
    public void Stream_byte_limit_is_enforced_before_decoding() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("@book{x,title={too long}}"));
        var options = new BibliographyReadOptions { MaximumInputBytes = 8 };

        Assert.Throws<InvalidDataException>(() => BibliographyDocument.Load(stream, BibliographyFormat.BibTex, options));
    }

    [Fact]
    public void Unknown_path_extension_uses_bounded_content_detection() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-bibliography-" + Guid.NewGuid().ToString("N") + ".data");
        try {
            File.WriteAllText(path, "TY  - BOOK\nID  - detected\nTI  - Detected path\nER  -\n");
            BibliographyReadResult result = BibliographyDocument.Load(path);
            Assert.Equal(BibliographyFormat.Ris, result.Document.SourceFormat);
            Assert.Equal("Detected path", Assert.Single(result.Document.Items).Title);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Item_limit_returns_diagnostic_and_bounded_partial_model() {
        var options = new BibliographyReadOptions { MaximumItemCount = 1 };
        BibliographyReadResult result = BibliographyDocument.Parse("@book{a,title={a}}\n@book{b,title={b}}", BibliographyFormat.BibTex, options);

        Assert.True(result.HasErrors);
        Assert.Single(result.Document.Items);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Fact]
    public void Nested_value_limit_returns_diagnostic() {
        var options = new BibliographyReadOptions { MaximumValueCount = 3 };
        BibliographyReadResult result = BibliographyDocument.Parse("[{\"id\":\"x\",\"author\":[{\"family\":\"Doe\",\"given\":\"Jane\"}]}]", BibliographyFormat.CslJson, options);

        Assert.True(result.HasErrors);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Fact]
    public void EndNote_xml_nesting_limit_returns_diagnostic() {
        var options = new BibliographyReadOptions { MaximumNestingDepth = 3 };

        BibliographyReadResult result = BibliographyDocument.Parse("<xml><records><record><titles><title>x</title></titles></record></records></xml>", BibliographyFormat.EndNoteXml, options);

        Assert.True(result.HasErrors);
        Assert.Empty(result.Document.Items);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Fact]
    public void Strict_writers_reject_required_safety_approximations() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "invalid\u0001xml" });

        BibliographyConversionLossException xml = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.EndNoteXml, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        Assert.Contains(xml.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV210");

        document.Items[0].Title = "unbalanced { brace";
        BibliographyConversionLossException bib = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.BibLatex, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        Assert.Contains(bib.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV211");
    }

    [Fact]
    public void EndNote_xml_prohibits_dtd_processing() {
        const string xml = "<!DOCTYPE xml [<!ENTITY xxe SYSTEM 'file:///etc/passwd'>]><xml><records><record><rec-number>1</rec-number><titles><title>&xxe;</title></titles></record></records></xml>";

        BibliographyReadResult result = BibliographyDocument.Parse(xml, BibliographyFormat.EndNoteXml);

        Assert.True(result.HasErrors);
        Assert.Empty(result.Document.Items);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "BIBEND002");
    }

    [Fact]
    public void Cancellation_is_observed() {
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => BibliographyDocument.Parse("[]", BibliographyFormat.CslJson, cancellationToken: cancellation.Token));
    }

    [Fact]
    public void Oversized_csl_date_number_is_retained_with_diagnostic_instead_of_throwing() {
        const string source = "[{\"id\":\"large-date\",\"type\":\"book\",\"issued\":{\"date-parts\":[[999999999999999999999]]}}]";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson);
        read.Document.Items[0].Title = "Edited";
        BibliographyWriteResult written = read.Document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.CslJson, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.False(read.HasErrors);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBCSL005");
        Assert.Contains("999999999999999999999", written.Content, StringComparison.Ordinal);
        Assert.False(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).HasErrors);
    }

    [Fact]
    public void Structurally_incomplete_bib_input_is_an_error() {
        BibliographyReadResult read = BibliographyDocument.Parse("@book{x,title={unterminated", BibliographyFormat.BibTex);

        Assert.True(read.HasErrors);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBBIB007" && diagnostic.Severity == BibliographyDiagnosticSeverity.Error);
    }
}
