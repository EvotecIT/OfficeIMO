namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyCodecContractTests {
    public static TheoryData<string, BibliographyFormat> Fixtures => new TheoryData<string, BibliographyFormat> {
        { "sample.bib", BibliographyFormat.BibLatex },
        { "sample.json", BibliographyFormat.CslJson },
        { "sample.ris", BibliographyFormat.Ris },
        { "sample.nbib", BibliographyFormat.Nbib },
        { "sample.xml", BibliographyFormat.EndNoteXml }
    };

    public static TheoryData<string, BibliographyFormat, BibliographyFormat> ConversionMatrix {
        get {
            var data = new TheoryData<string, BibliographyFormat, BibliographyFormat>();
            foreach (object[] source in Fixtures) {
                foreach (BibliographyFormat destination in Enum.GetValues(typeof(BibliographyFormat))) data.Add((string)source[0], (BibliographyFormat)source[1], destination);
            }
            return data;
        }
    }

    [Theory]
    [MemberData(nameof(Fixtures))]
    public void Unchanged_fixture_is_preserved_exactly(string fileName, BibliographyFormat format) {
        byte[] source = File.ReadAllBytes(Fixture(fileName));
        BibliographyReadResult read = BibliographyDocument.Load(new MemoryStream(source), format);

        BibliographyWriteResult written = read.Document.Write();

        Assert.False(read.HasErrors);
        Assert.True(written.UsedOriginalSource);
        Assert.Equal(source, written.Bytes);
        Assert.False(written.Report.HasLoss);
    }

    [Theory]
    [MemberData(nameof(Fixtures))]
    public void Edited_canonical_output_is_deterministic_and_reopens(string fileName, BibliographyFormat format) {
        BibliographyDocument document = BibliographyDocument.Load(Fixture(fileName), format).Document;
        Assert.Single(document.Items);
        document.Items[0].Title = "Edited citation title";
        var options = new BibliographyWriteOptions { Format = format, Mode = BibliographyWriterMode.Canonical, LineEnding = "\n" };

        BibliographyWriteResult first = document.Write(options);
        BibliographyWriteResult second = document.Write(options);
        BibliographyReadResult reopened = BibliographyDocument.Parse(first.Content, format);

        Assert.False(first.UsedOriginalSource);
        Assert.Equal(first.Bytes, second.Bytes);
        Assert.False(reopened.HasErrors);
        Assert.Equal("Edited citation title", Assert.Single(reopened.Document.Items).Title);
        Assert.Contains("retained", first.Content, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [MemberData(nameof(Fixtures))]
    public void Same_format_canonical_write_can_require_no_loss(string fileName, BibliographyFormat format) {
        BibliographyDocument document = BibliographyDocument.Load(Fixture(fileName), format).Document;
        document.Items[0].Title = "Strict same-format edit";

        BibliographyWriteResult result = document.Write(new BibliographyWriteOptions { Format = format, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.False(result.Report.HasLoss);
        Assert.Equal("Strict same-format edit", BibliographyDocument.Parse(result.Content, format).Document.Items[0].Title);
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.CslJson)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void New_document_writes_and_reopens_in_every_format(BibliographyFormat format) {
        BibliographyDocument document = CreateDocument();

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Format = format, Mode = BibliographyWriterMode.Canonical });
        BibliographyReadResult reopened = BibliographyDocument.Parse(written.Content, format);

        Assert.False(reopened.HasErrors);
        BibliographyItem item = Assert.Single(reopened.Document.Items);
        Assert.Equal("smith2025", item.Key);
        Assert.Equal("Interoperable Bibliographies", item.Title);
        Assert.NotNull(item.GetDate(BibliographyDateRole.Issued));
    }

    [Theory]
    [MemberData(nameof(ConversionMatrix))]
    public void Every_source_format_converts_deterministically_and_reopens(string fileName, BibliographyFormat sourceFormat, BibliographyFormat destinationFormat) {
        BibliographyDocument source = BibliographyDocument.Load(Fixture(fileName), sourceFormat).Document;
        var options = new BibliographyWriteOptions { Format = destinationFormat, Mode = BibliographyWriterMode.Canonical };

        BibliographyWriteResult first = source.Write(options);
        BibliographyWriteResult second = source.Write(options);
        BibliographyReadResult reopened = BibliographyDocument.Parse(first.Content, destinationFormat);

        Assert.Equal(first.Bytes, second.Bytes);
        Assert.False(reopened.HasErrors);
        Assert.Equal(source.Items[0].Title, Assert.Single(reopened.Document.Items).Title);
        bool sameBibFamily = (sourceFormat == BibliographyFormat.BibTex || sourceFormat == BibliographyFormat.BibLatex) && (destinationFormat == BibliographyFormat.BibTex || destinationFormat == BibliographyFormat.BibLatex);
        if (sourceFormat != destinationFormat && !sameBibFamily) Assert.True(first.Report.HasLoss);
    }

    [Fact]
    public void Strict_conversion_rejects_unrepresentable_data() {
        BibliographyDocument document = CreateDocument();
        document.Items[0].Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Submitted, Year = 2023 });
        document.Items[0].NativeFields.Add(new BibliographyNativeField(BibliographyFormat.CslJson, "x-private", "value", "{\"nested\":true}"));

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions {
            Format = BibliographyFormat.Ris,
            Mode = BibliographyWriterMode.Canonical,
            RequireNoLoss = true
        }));

        Assert.True(exception.Report.HasLoss);
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Action == BibliographyConversionAction.Omitted && diagnostic.Field == "dates.Submitted");
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Field == "x-private");
    }

    [Fact]
    public void Bib_parser_handles_nested_braces_concatenation_and_string_macros() {
        const string source = "@string{j = {Journal}}\n@article{k,title={A {Nested} Title},journal=j # { Letters},author={Doe, Jane and {Example Group}},year=2024}";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex);
        BibliographyItem item = Assert.Single(read.Document.Items);

        Assert.Equal("A {Nested} Title", item.Title);
        Assert.Equal("Journal Letters", item.ContainerTitle);
        Assert.Equal(2, item.Contributors.Count);
        Assert.Equal(2024, item.GetDate(BibliographyDateRole.Issued)?.Year);
    }

    [Fact]
    public void Csl_unknown_json_value_survives_canonical_same_format_write() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"x\",\"type\":\"book\",\"author\":[{\"literal\":\"Team\",\"ORCID\":\"id\"}],\"issued\":{\"date-parts\":[[2025]],\"circa\":true},\"x-data\":{\"a\":[1,true]}}]", BibliographyFormat.CslJson).Document;
        document.Items[0].Title = "Changed";

        string output = document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.CslJson, Mode = BibliographyWriterMode.Canonical }).Content;

        Assert.Contains("\"x-data\"", output);
        Assert.Contains("\"a\"", output);
        Assert.Contains("true", output);
        Assert.Contains("\"ORCID\"", output);
        Assert.Contains("\"circa\"", output);
    }

    [Theory]
    [InlineData("@custom{x,title={Custom}}", BibliographyFormat.BibLatex, "custom")]
    [InlineData("[{\"id\":\"x\",\"type\":\"custom-type\",\"title\":\"Custom\"}]", BibliographyFormat.CslJson, "custom-type")]
    [InlineData("TY  - CUST\nID  - x\nTI  - Custom\nER  -\n", BibliographyFormat.Ris, "CUST")]
    public void Safe_custom_types_survive_strict_same_format_writes(string source, BibliographyFormat format, string expectedType) {
        BibliographyDocument document = BibliographyDocument.Parse(source, format).Document;
        document.Items[0].Title = "Changed";

        BibliographyWriteResult result = document.Write(new BibliographyWriteOptions { Format = format, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = BibliographyDocument.Parse(result.Content, format).Document.Items[0];

        Assert.Equal(BibliographyItemType.Unknown, reopened.Type);
        Assert.Equal(expectedType, reopened.NativeType, ignoreCase: true);
    }

    [Fact]
    public void Csl_date_range_survives_strict_canonical_edit() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"range\",\"type\":\"book\",\"issued\":{\"date-parts\":[[2024,1,2],[2025,3,4]]}}]", BibliographyFormat.CslJson).Document;
        document.Items[0].Title = "Edited range";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.CslJson, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate date = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items[0].Dates);

        Assert.Equal(2024, date.Year); Assert.Equal(1, date.Month); Assert.Equal(2, date.Day);
        Assert.Equal(2025, date.EndYear); Assert.Equal(3, date.EndMonth); Assert.Equal(4, date.EndDay);
    }

    [Fact]
    public void Date_range_reports_loss_for_destination_without_exact_range_contract() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"range\",\"type\":\"book\",\"issued\":{\"date-parts\":[[2024],[2025]]}}]", BibliographyFormat.CslJson).Document;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.Ris, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV219");
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.CslJson)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Missing_keys_receive_unique_deterministic_fallbacks(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        document.Items.Add(new BibliographyItem { Type = BibliographyItemType.Book, Title = "First" });
        document.Items.Add(new BibliographyItem { Type = BibliographyItemType.Book, Title = "Second" });

        BibliographyWriteResult first = document.Write(new BibliographyWriteOptions { Format = format, Mode = BibliographyWriterMode.Canonical });
        BibliographyWriteResult second = document.Write(new BibliographyWriteOptions { Format = format, Mode = BibliographyWriterMode.Canonical });
        string[] keys = BibliographyDocument.Parse(first.Content, format).Document.Items.Select(static item => item.Key).ToArray();

        Assert.Equal(first.Bytes, second.Bytes);
        Assert.Equal(new[] { "item-1", "item-2" }, keys);
        Assert.Contains(first.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV215");
    }

    [Fact]
    public void Strict_write_rejects_duplicate_keys() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        document.Items.Add(new BibliographyItem { Key = "duplicate", Type = BibliographyItemType.Book, Title = "First" });
        document.Items.Add(new BibliographyItem { Key = "DUPLICATE", Type = BibliographyItemType.Book, Title = "Second" });

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Equal(2, exception.Report.Diagnostics.Count(diagnostic => diagnostic.Code == "BIBCONV216"));
    }

    [Fact]
    public void Nbib_unknown_safe_tag_survives_strict_same_format_edit() {
        BibliographyDocument document = BibliographyDocument.Parse("PMID- 1\nPT  - Randomized Controlled Trial\nTI  - Original\n", BibliographyFormat.Nbib).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.Nbib, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Nbib).Document.Items);

        Assert.Contains(reopened.NativeFields, field => field.Name == "PT" && field.Value == "Randomized Controlled Trial");
    }

    private static BibliographyDocument CreateDocument() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem {
            Key = "smith2025", Type = BibliographyItemType.ArticleJournal, Title = "Interoperable Bibliographies", ContainerTitle = "Journal of Formats",
            Publisher = "Example Press", PublisherPlace = "Warsaw", Volume = "12", Issue = "3", Pages = "10-20", Abstract = "A format-neutral test.", Language = "en", Url = "https://example.test/item"
        };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Given = "Alex", Family = "Smith" }));
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Year = 2025, Month = 6, Day = 2 });
        item.Identifiers.Add(new BibliographyIdentifier("DOI", "10.1000/test")); item.Keywords.Add("bibliography"); item.Notes.Add("Verified fixture"); document.Items.Add(item);
        return document;
    }

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", name);
}
