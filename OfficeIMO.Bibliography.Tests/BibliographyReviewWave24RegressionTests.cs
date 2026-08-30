namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave24RegressionTests {
    [Fact]
    public void CSL_canonical_edits_preserve_recognized_native_type_spelling() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"x\",\"type\":\"JOUR\",\"title\":\"Before\"}]", BibliographyFormat.CslJson).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        Assert.Equal("JOUR", reopened.NativeType);
        Assert.Contains("\"type\": \"JOUR\"", written.Content, StringComparison.Ordinal);
    }

    [Fact]
    public void RIS_canonical_edits_preserve_recognized_native_type_spelling() {
        BibliographyDocument document = BibliographyDocument.Parse("TY  - jour\nID  - x\nTI  - Before\nER  -\n", BibliographyFormat.Ris).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.Equal("jour", reopened.NativeType);
        Assert.StartsWith("TY  - jour\n", written.Content, StringComparison.Ordinal);
    }

    [Fact]
    public void Blank_CSL_identifiers_remain_native_after_other_edits() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"x\",\"type\":\"book\",\"DOI\":\"\",\"title\":\"Before\"}]", BibliographyFormat.CslJson).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        Assert.Empty(reopened.Identifiers);
        BibliographyNativeField native = Assert.Single(reopened.NativeFields, field => field.Name == "DOI");
        Assert.Equal(string.Empty, native.Value);
    }

    [Theory]
    [InlineData("<!--long-->")]
    [InlineData("<?review long?>")]
    public void EndNote_XML_materialization_bounds_comments_and_processing_instructions(string trivia) {
        string source = "<xml>" + trivia + "<records/></xml>";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml, new BibliographyReadOptions { MaximumValueLength = 3 });

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Fact]
    public void Invalid_UTF16_in_CSL_input_is_rejected_before_replacement() {
        string source = "[{\"id\":\"x\",\"title\":\"" + '\ud800' + "\"}]";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson);

        Assert.Empty(read.Document.Items);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBCSL002" && diagnostic.Severity == BibliographyDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData("isbn")]
    [InlineData("electronic-resource-num")]
    [InlineData("accession-num")]
    public void Empty_EndNote_identifier_elements_reopen_after_strict_edits(string elementName) {
        string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles><" + elementName + "/></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Contains(reopened.NativeFields, field => string.Equals(field.Name, elementName, StringComparison.OrdinalIgnoreCase) && field.Value.Length == 0);
    }

    [Fact]
    public void Bib_macro_expansion_is_bounded_cumulatively() {
        string large = new string('x', 100);
        string source = "@string{a={" + large + "}}\n" + string.Join("\n", Enumerable.Range(0, 8).Select(index => "@string{s" + index + "=a}"));

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex, new BibliographyReadOptions { MaximumInputCharacters = 512 });

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001" && diagnostic.Message.Contains("cumulative expanded", StringComparison.Ordinal));
    }

    [Fact]
    public void Bib_contributor_serialization_observes_cancellation() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book };
        for (int index = 0; index < 200_000; index++) item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Family" + index }));
        document.Items.Add(item);
        using var cancellation = new CancellationTokenSource();
        cancellation.CancelAfter(1);

        Assert.Throws<OperationCanceledException>(() => BibCodec.Write(document, BibliographyFormat.BibLatex, new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, new BibliographyConversionReport(), cancellation.Token));
    }
}
