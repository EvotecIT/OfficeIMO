using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyParserFidelityTests {
    [Fact]
    public void Partial_NBIB_full_author_list_is_matched_to_compact_authors() {
        const string source = "PMID- 1\nFAU - Smith, John\nAU  - Smith J\nAU  - Brown B\nFAU - Jones, Jane\nAU  - Jones J\nTI  - Authors\n";

        BibliographyItem item = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document.Items);

        Assert.Equal(new[] { "Smith", "Brown", "Jones" }, item.Contributors.Select(static contributor => contributor.Name.Family));
        Assert.DoesNotContain(item.NativeFields, static field => field.Name == "AU");
    }

    [Fact]
    public void NBIB_full_compact_matching_preserves_interleaved_collective_author_order() {
        const string source = "PMID- 1\nFAU - Smith, John\nAU  - Smith J\nCN  - Research Group\nAU  - Brown B\nFAU - Jones, Jane\nAU  - Jones J\nTI  - Authors\n";

        BibliographyItem item = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document.Items);

        Assert.Equal(new[] { "Smith", "Research Group", "Brown", "Jones" }, item.Contributors.Select(static contributor => contributor.Name.Literal ?? contributor.Name.Family));
    }

    [Theory]
    [InlineData("title")]
    [InlineData("publisher")]
    [InlineData("DOI")]
    [InlineData("keyword")]
    public void Object_valued_CSL_scalars_remain_native_JSON(string property) {
        string source = "[{\"id\":\"x\",\"type\":\"book\",\"" + property + "\":{\"value\":\"Example\"}}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        using JsonDocument reopened = JsonDocument.Parse(written.Content);
        JsonElement value = reopened.RootElement[0].GetProperty(property);

        Assert.Equal(JsonValueKind.Object, value.ValueKind);
        Assert.Equal("Example", value.GetProperty("value").GetString());
    }

    [Theory]
    [InlineData("ISBN", "ABC-123")]
    [InlineData("ISSN", "ABC-123")]
    [InlineData("SN", "9781402894626")]
    public void Ambiguous_RIS_serial_scheme_blocks_strict_output(string scheme, string value) {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Serial" };
        item.Identifiers.Add(new BibliographyIdentifier(scheme, value));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV228" && diagnostic.Field == "identifiers." + scheme);
    }

    [Theory]
    [InlineData("<!-- generated -->")]
    [InlineData("<?generated value?>")]
    [InlineData("<!-- generated --><?generated value?>")]
    public void EndNote_auto_detection_skips_leading_XML_trivia(string trivia) {
        BibliographyReadResult read = BibliographyDocument.Parse(trivia + "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type></record></records></xml>");

        Assert.Equal(BibliographyFormat.EndNoteXml, read.Document.SourceFormat);
        Assert.Equal("1", Assert.Single(read.Document.Items).Key);
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Single_token_family_names_reopen_as_structured_names(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = format == BibliographyFormat.Nbib ? "1" : "x", Type = format == BibliographyFormat.Nbib ? BibliographyItemType.ArticleJournal : BibliographyItemType.Book, Title = "Names" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Smith" }));
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyContributor reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items[0].Contributors);

        Assert.Equal("Smith", reopened.Name.Family);
        Assert.Null(reopened.Name.Literal);
    }

    [Fact]
    public void Recognized_RIS_scalar_continuation_updates_the_typed_field() {
        const string source = "TY  - BOOK\nID  - x\nPB  - Example\n      Press\nTI  - Work\nER  -\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.Equal("Example Press", reopened.Publisher);
        Assert.DoesNotContain(reopened.NativeFields, static field => field.Name == "PB");
    }

    [Fact]
    public void Recognized_RIS_page_date_and_identifier_continuations_update_typed_values() {
        const string source = "TY  - BOOK\nID  - x\nPY  - Spring\n      2025\nSP  - S\n      10\nEP  - S\n      20\nDO  - 10.1/\n      example\nER  -\n";
        BibliographyItem item = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document.Items);

        Assert.Equal("Spring 2025", Assert.Single(item.Dates).Literal);
        Assert.Equal("S 10-S 20", item.Pages);
        Assert.Equal("10.1/ example", Assert.Single(item.Identifiers).Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BOM_based_loading_detects_UTF32_before_UTF16(bool bigEndian) {
        const string source = "@book{x,title={UTF32}}";
        var encoding = new UTF32Encoding(bigEndian, true, true);
        byte[] bytes = encoding.GetPreamble().Concat(encoding.GetBytes(source)).ToArray();

        BibliographyReadResult read = BibliographyDocument.Load(new MemoryStream(bytes), BibliographyFormat.BibLatex);

        Assert.Equal("UTF32", Assert.Single(read.Document.Items).Title);
    }

    [Theory]
    [InlineData("doi")]
    [InlineData("isbn")]
    [InlineData("pmid")]
    public void Empty_Bib_identifiers_are_retained_instead_of_throwing(string field) {
        string source = "@book{x,title={Empty}," + field + "={}}";

        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex).Document;
        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains(document.Items[0].NativeFields, native => native.Name == field && native.Value == string.Empty);
        Assert.Contains(field + " = {}", written.Content, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("10.1/example []")]
    [InlineData("10.1/example [   ]")]
    [InlineData(" [doi]")]
    public void Invalid_NBIB_identifier_qualifiers_are_retained_as_native_tags(string value) {
        string source = "PMID- 1\nAID - " + value + "\nTI  - Identifier\n";

        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;
        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains(document.Items[0].NativeFields, field => field.Name == "AID" && field.Value == value.TrimStart());
        Assert.Contains("AID - " + value.TrimStart(), written.Content, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Book", BibliographyItemType.Book)]
    [InlineData("Book Chapter", BibliographyItemType.Chapter)]
    public void NBIB_publication_type_sets_the_typed_item_kind(string publicationType, BibliographyItemType expected) {
        string source = "PMID- 1\nPT  - " + publicationType + "\nTI  - Typed\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;

        Assert.Equal(expected, Assert.Single(document.Items).Type);
        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        Assert.Equal(expected, Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Nbib).Document.Items).Type);
    }
}
