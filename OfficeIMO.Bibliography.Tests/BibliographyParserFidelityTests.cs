using System.Text.Json;
using System.Xml.Linq;

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

    [Theory]
    [InlineData("Ludwig van Beethoven", "Ludwig", "van", "Beethoven")]
    [InlineData("Charles de la Vallée Poussin", "Charles", "de la", "Vallée Poussin")]
    public void No_comma_Bib_names_apply_first_von_last_rules(string sourceName, string expectedGiven, string expectedParticle, string expectedFamily) {
        string source = "@book{x,title={Names},author={" + sourceName + "}}";

        BibliographyName name = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.BibTex).Document.Items[0].Contributors).Name;

        Assert.Equal(expectedGiven, name.Given);
        Assert.Equal(expectedParticle, name.NonDroppingParticle);
        Assert.Equal(expectedFamily, name.Family);
    }

    [Fact]
    public void NBIB_location_identifier_tag_survives_strict_edit_and_reopen() {
        const string source = "PMID- 1\nLID - 10.1000/example [doi]\nTI  - Original\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains("LID - 10.1000/example [doi]", written.Content, StringComparison.Ordinal);
        Assert.Equal("10.1000/example", Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Nbib).Document.Items[0].Identifiers, identifier => identifier.Scheme == "doi").Value);
    }

    [Fact]
    public void Identifier_properties_preserve_constructor_invariants() {
        var identifier = new BibliographyIdentifier("DOI", "10.1/example");

        Assert.Throws<ArgumentException>(() => identifier.Scheme = " ");
        Assert.Throws<ArgumentException>(() => identifier.Scheme = "bad\nscheme");
        Assert.Throws<ArgumentException>(() => identifier.Value = string.Empty);
        Assert.Equal("DOI", identifier.Scheme);
        Assert.Equal("10.1/example", identifier.Value);
    }

    [Fact]
    public void Classic_BibTeX_translator_blocks_strict_output() {
        var document = new BibliographyDocument(BibliographyFormat.BibTex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Translation" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Translator, new BibliographyName { Given = "Taylor", Family = "Smith" }));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV201" && diagnostic.Field == "contributors.Translator");
    }

    [Theory]
    [InlineData("custom")]
    [InlineData("title")]
    [InlineData("url")]
    public void Edited_EndNote_extension_retains_its_foreign_namespace(string fieldName) {
        string source = "<xml xmlns:ext=\"urn:extension\"><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><ext:" + fieldName + " mode=\"safe\">old</ext:" + fieldName + "></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        BibliographyNativeField field = Assert.Single(document.Items[0].NativeFields, native => native.Name == fieldName);
        field.Value = "new";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        XDocument reopened = XDocument.Parse(written.Content);
        XElement extension = Assert.Single(reopened.Descendants(XName.Get(fieldName, "urn:extension")));

        Assert.Equal("new", extension.Value);
        Assert.Equal("safe", extension.Attribute("mode")?.Value);
    }

    [Fact]
    public void CSL_typed_property_matching_is_case_sensitive_at_every_level() {
        const string source = "[{\"id\":\"x\",\"type\":\"book\",\"title\":\"typed\",\"Title\":\"extension\",\"URL\":\"https://typed.example\",\"Url\":\"extension-url\",\"author\":[{\"family\":\"Smith\",\"Family\":\"extension-family\"}],\"issued\":{\"literal\":\"2026\",\"Literal\":\"extension-date\"}}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        using JsonDocument json = JsonDocument.Parse(written.Content);
        JsonElement item = json.RootElement[0];

        Assert.Equal("typed", item.GetProperty("title").GetString());
        Assert.Equal("extension", item.GetProperty("Title").GetString());
        Assert.Equal("https://typed.example", item.GetProperty("URL").GetString());
        Assert.Equal("extension-url", item.GetProperty("Url").GetString());
        Assert.Equal("extension-family", item.GetProperty("author")[0].GetProperty("Family").GetString());
        Assert.Equal("extension-date", item.GetProperty("issued").GetProperty("Literal").GetString());
    }

    [Fact]
    public void Classic_BibTeX_named_month_is_emitted_as_a_month_name() {
        BibliographyDocument document = BibliographyDocument.Parse("@book{x,title={Month},year=2024,month=jan}", BibliographyFormat.BibTex).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains("month = {January}", written.Content, StringComparison.Ordinal);
        Assert.Equal(1, Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibTex).Document.Items[0].Dates).Month);
    }

    [Fact]
    public void Unknown_classic_BibTeX_month_is_retained_as_a_native_field() {
        BibliographyDocument document = BibliographyDocument.Parse("@book{x,title={Month},year=2024,month=notamonth}", BibliographyFormat.BibTex).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains("month = {notamonth}", written.Content, StringComparison.Ordinal);
    }

    [Fact]
    public void NBIB_type_edit_reports_stale_publication_type_before_replacement() {
        BibliographyDocument document = BibliographyDocument.Parse("PMID- 1\nPT  - Book\nTI  - Type\n", BibliographyFormat.Nbib).Document;
        document.Items[0].Type = BibliographyItemType.ArticleJournal;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Nbib).Document.Items);

        Assert.Contains(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV122" && diagnostic.Field == "PT");
        Assert.DoesNotContain("PT  - Book", written.Content, StringComparison.Ordinal);
        Assert.Contains("PT  - Journal Article", written.Content, StringComparison.Ordinal);
        Assert.Equal(BibliographyItemType.ArticleJournal, reopened.Type);
        Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
    }

    [Fact]
    public void Quoted_BibTeX_values_track_braces_around_inner_quotes() {
        const string source = "@book{x,title=\"An {\"inner\"} title\",publisher={Example Press}}";

        BibliographyItem item = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.BibTex).Document.Items);

        Assert.Equal("An {\"inner\"} title", item.Title);
        Assert.Equal("Example Press", item.Publisher);
    }

    [Fact]
    public void Invalid_classic_BibTeX_month_returns_conversion_diagnostics() {
        var document = new BibliographyDocument(BibliographyFormat.BibTex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Month" };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Year = 2026, Month = 14 });
        document.Items.Add(item);

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains("month = {14}", permissive.Content, StringComparison.Ordinal);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV218");
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Mixed_literal_and_personal_names_block_strict_output(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = format == BibliographyFormat.Nbib ? BibliographyItemType.ArticleJournal : BibliographyItemType.Book, Title = "Names" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Literal = "Research Group", Given = "Ignored", Family = "Person" }));
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV231" && diagnostic.Field == "contributors");
    }

    [Fact]
    public void Ambiguous_NBIB_identifier_scheme_blocks_strict_output() {
        var document = new BibliographyDocument(BibliographyFormat.Nbib);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.ArticleJournal, Title = "Identifier" };
        item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        item.Identifiers.Add(new BibliographyIdentifier("archive [local", "123"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV232" && diagnostic.Field == "identifiers.archive [local");
    }

    [Fact]
    public void Parsed_EndNote_preserve_mode_honors_the_declared_byte_encoding() {
        const string source = "<?xml version=\"1.0\" encoding=\"utf-16\"?><xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;

        BibliographyWriteResult written = document.Write();

        Assert.Equal(new byte[] { 0xFF, 0xFE }, written.Bytes.Take(2));
        Assert.Equal(source, Encoding.Unicode.GetString(written.Bytes, 2, written.Bytes.Length - 2));
        Assert.Equal("1", Assert.Single(BibliographyDocument.Load(new MemoryStream(written.Bytes), BibliographyFormat.EndNoteXml).Document.Items).Key);
    }

    [Fact]
    public void Generic_document_is_exact_in_CSL_JSON() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Document, Title = "Document" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        Assert.Equal(BibliographyItemType.Document, reopened.Type);
    }

    [Fact]
    public void CSL_JSON_citation_keys_are_compared_case_sensitively() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        document.Items.Add(new BibliographyItem { Key = "item", Type = BibliographyItemType.Book, Title = "Lower" });
        document.Items.Add(new BibliographyItem { Key = "ITEM", Type = BibliographyItemType.Book, Title = "Upper" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Equal(new[] { "item", "ITEM" }, BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items.Select(static item => item.Key));
    }

    [Fact]
    public void EndNote_year_and_publication_text_merge_without_duplicating_the_year() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><dates><year>2024</year><pub-dates><date>May 2024</date></pub-dates></dates></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        BibliographyDate parsed = Assert.Single(document.Items[0].Dates);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items[0].Dates);

        Assert.Equal(2024, parsed.Year);
        Assert.Equal(5, parsed.Month);
        Assert.Equal(2024, reopened.Year);
        Assert.Equal(5, reopened.Month);
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Native_Bib_field_brace_escaping_blocks_strict_output(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Native" };
        item.NativeFields.Add(new BibliographyNativeField(format, "custom", "unbalanced { value"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV233" && diagnostic.Field == "native.custom");
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Parsed_Bib_item_type_edits_use_the_new_exact_type(BibliographyFormat format) {
        BibliographyDocument document = BibliographyDocument.Parse("@book{x,title={Type}}", format).Document;
        document.Items[0].Type = BibliographyItemType.ArticleJournal;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items);

        Assert.StartsWith("@article{", written.Content, StringComparison.Ordinal);
        Assert.Equal(BibliographyItemType.ArticleJournal, reopened.Type);
    }

    [Fact]
    public void CSL_date_role_order_survives_strict_canonical_output() {
        const string source = "[{\"id\":\"x\",\"type\":\"book\",\"accessed\":{\"date-parts\":[[2026,2,3]]},\"issued\":{\"date-parts\":[[2025,1,2]]}}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        Assert.Equal(new[] { BibliographyDateRole.Accessed, BibliographyDateRole.Issued }, reopened.Dates.Select(static date => date.Role));
    }

    [Fact]
    public void Generic_thesis_is_exact_in_BibLaTeX() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Thesis, Title = "Thesis" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibLatex).Document.Items);

        Assert.StartsWith("@thesis{", written.Content, StringComparison.Ordinal);
        Assert.Equal(BibliographyItemType.Thesis, reopened.Type);
    }

    [Fact]
    public void Generic_thesis_reports_narrowing_in_classic_BibTeX() {
        var document = new BibliographyDocument(BibliographyFormat.BibTex);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Thesis, Title = "Thesis" });

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.StartsWith("@phdthesis{", permissive.Content, StringComparison.Ordinal);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV200" && diagnostic.Field == "type");
    }

    [Theory]
    [InlineData("title", "title")]
    [InlineData("date", "dates.Issued.literal")]
    [InlineData("native", "native.custom")]
    public void EndNote_carriage_return_normalization_blocks_strict_output(string valueOwner, string expectedField) {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Text" };
        if (valueOwner == "title") item.Title = "A\r\nB";
        else if (valueOwner == "date") item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Literal = "A\rB" });
        else item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.EndNoteXml, "custom", "A\rB"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV235" && diagnostic.Field == expectedField);
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris, "given")]
    [InlineData(BibliographyFormat.Ris, "suffix")]
    [InlineData(BibliographyFormat.Nbib, "given")]
    [InlineData(BibliographyFormat.Nbib, "suffix")]
    [InlineData(BibliographyFormat.EndNoteXml, "given")]
    [InlineData(BibliographyFormat.EndNoteXml, "suffix")]
    public void Tagged_name_output_diagnoses_missing_family_positions(BibliographyFormat format, string component) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = format == BibliographyFormat.Nbib ? "1" : "x", Type = format == BibliographyFormat.Nbib ? BibliographyItemType.ArticleJournal : BibliographyItemType.Book, Title = "Names" };
        var name = new BibliographyName();
        if (component == "given") name.Given = "Cher";
        else name.Suffix = "Jr.";
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, name));
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV244" && diagnostic.Field == "contributors");
    }

    [Fact]
    public void Generic_document_is_exact_in_RIS() {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Document, Title = "Document" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.Contains("TY  - GEN", written.Content, StringComparison.Ordinal);
        Assert.Equal(BibliographyItemType.Document, reopened.Type);
    }

    [Theory]
    [InlineData("\uFEFF@book{x,title={BOM}}", BibliographyFormat.BibLatex)]
    [InlineData("\uFEFF[{\"id\":\"x\",\"type\":\"book\"}]", BibliographyFormat.CslJson)]
    [InlineData("\uFEFFTY  - BOOK\nID  - x\nER  -\n", BibliographyFormat.Ris)]
    [InlineData("\uFEFFPMID- 1\nTI  - BOM\n", BibliographyFormat.Nbib)]
    [InlineData("\uFEFF<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type></record></records></xml>", BibliographyFormat.EndNoteXml)]
    public void Auto_detection_skips_a_leading_BOM(string source, BibliographyFormat expected) {
        BibliographyReadResult read = BibliographyDocument.Parse(source);

        Assert.Equal(expected, read.Document.SourceFormat);
        Assert.False(read.HasErrors);
        Assert.Single(read.Document.Items);
    }

    [Fact]
    public void NBIB_identifier_order_survives_strict_canonical_output() {
        var document = new BibliographyDocument(BibliographyFormat.Nbib);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.ArticleJournal, Title = "Identifiers" };
        item.Identifiers.Add(new BibliographyIdentifier("DOI", "10.1/example"));
        item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        item.Identifiers.Add(new BibliographyIdentifier("ISSN", "1234-5678"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Nbib).Document.Items);

        Assert.Equal(new[] { "DOI", "PMID", "ISSN" }, reopened.Identifiers.Select(static identifier => identifier.Scheme.ToUpperInvariant()));
    }

    [Fact]
    public void EndNote_identifier_order_survives_strict_canonical_output() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><electronic-resource-num>10.1/example</electronic-resource-num><isbn>978-1-4028-9462-6</isbn><accession-num>2608.00001</accession-num></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;

        Assert.Equal(new[] { "DOI", "ISBN", "accession" }, document.Items[0].Identifiers.Select(static identifier => identifier.Scheme));

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal(new[] { "DOI", "ISBN", "accession" }, reopened.Identifiers.Select(static identifier => identifier.Scheme));
    }

    [Fact]
    public void Undefined_CSL_contributor_roles_block_strict_output() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Role" };
        item.Contributors.Add(new BibliographyContributor((BibliographyContributorRole)99, new BibliographyName { Family = "Doe" }));
        document.Items.Add(item);

        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.DoesNotContain("Doe", permissive.Content, StringComparison.Ordinal);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV201" && diagnostic.Field == "contributors.99");
    }

    [Fact]
    public void Undefined_CSL_item_and_date_roles_block_strict_output() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "x", Type = (BibliographyItemType)99, Title = "Enums" };
        item.Dates.Add(new BibliographyDate { Role = (BibliographyDateRole)99, Year = 2026 });
        document.Items.Add(item);

        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV200" && diagnostic.Field == "type");
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV202" && diagnostic.Field == "dates.99");
    }

    [Fact]
    public void Escaped_BibTeX_braces_do_not_group_name_separators() {
        const string source = "@book{x,author={Doe, John \\{ and Smith, Jane}}";

        BibliographyItem item = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.BibTex).Document.Items);

        Assert.Equal(2, item.Contributors.Count);
        Assert.Equal(new[] { "Doe", "Smith" }, item.Contributors.Select(static contributor => contributor.Name.Family));
    }

    [Theory]
    [InlineData(BibliographyItemType.PaperConference, "CPAPER")]
    [InlineData(BibliographyItemType.Proceedings, "CONF")]
    public void RIS_conference_types_use_distinct_standard_tokens(BibliographyItemType type, string token) {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        document.Items.Add(new BibliographyItem { Key = "x", Type = type, Title = "Conference" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.StartsWith("TY  - " + token, written.Content, StringComparison.Ordinal);
        Assert.Equal(type, reopened.Type);
    }
}
