using System.Text.Json;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographySourceFidelityTests {
    [Fact]
    public void Single_item_CSL_object_root_survives_a_strict_canonical_edit() {
        BibliographyDocument document = BibliographyDocument.Parse("{\"id\":\"x\",\"type\":\"book\",\"title\":\"Before\"}", BibliographyFormat.CslJson).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        using JsonDocument json = JsonDocument.Parse(written.Content);

        Assert.Equal(JsonValueKind.Object, json.RootElement.ValueKind);
        Assert.Equal("After", json.RootElement.GetProperty("title").GetString());
        Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);
    }

    [Fact]
    public void Single_item_CSL_object_root_blocks_strict_output_after_item_count_changes() {
        BibliographyDocument document = BibliographyDocument.Parse("{\"id\":\"x\",\"type\":\"book\"}", BibliographyFormat.CslJson).Document;
        document.Items.Add(new BibliographyItem { Key = "y", Type = BibliographyItemType.Book });

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV130");
    }

    [Fact]
    public void Empty_CSL_extension_property_names_survive_a_strict_canonical_edit() {
        const string source = "[{\"id\":\"x\",\"type\":\"book\",\"\":1,\"   \":{\"nested\":true}}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        using JsonDocument json = JsonDocument.Parse(written.Content);
        JsonElement item = json.RootElement[0];

        Assert.Equal(1, item.GetProperty(string.Empty).GetInt32());
        Assert.True(item.GetProperty("   ").GetProperty("nested").GetBoolean());
    }

    [Fact]
    public void EndNote_root_container_and_record_attributes_survive_a_strict_edit() {
        const string source = "<xml xmlns=\"urn:endnote\" data-root=\"root\" xmlns:p=\"urn:test\" p:mode=\"exact\"><records data-container=\"records\"><record data-record=\"item\"><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        XDocument xml = XDocument.Parse(written.Content);
        XElement root = xml.Root!;
        XElement records = root.Elements().Single(element => element.Name.LocalName == "records");
        XElement record = records.Elements().Single(element => element.Name.LocalName == "record");

        Assert.Equal("urn:endnote", root.Name.NamespaceName);
        Assert.Equal("root", root.Attribute("data-root")?.Value);
        Assert.Equal("exact", root.Attribute(XName.Get("mode", "urn:test"))?.Value);
        Assert.Equal("records", records.Attribute("data-container")?.Value);
        Assert.Equal("item", record.Attribute("data-record")?.Value);
    }

    [Fact]
    public void Direct_EndNote_records_root_and_unknown_container_children_survive_a_strict_edit() {
        const string source = "<records data-container=\"records\"><metadata><label>retained</label></metadata><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles></record></records>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        XDocument xml = XDocument.Parse(written.Content);

        Assert.Equal("records", xml.Root?.Name.LocalName);
        Assert.Equal("records", xml.Root?.Attribute("data-container")?.Value);
        Assert.Equal("retained", xml.Root?.Elements().Single(element => element.Name.LocalName == "metadata").Value);
        Assert.Equal("After", Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items).Title);
    }

    [Theory]
    [InlineData("unsafe{key")]
    [InlineData("unsafe%key")]
    [InlineData("unsafe#key")]
    [InlineData("unsafe(key")]
    [InlineData("unsafe\\key")]
    public void Unsafe_Bib_citation_key_characters_are_normalized_and_reported(string key) {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        document.Items.Add(new BibliographyItem { Key = key, Type = BibliographyItemType.Book, Title = "Safe output" });

        BibliographyWriteResult canonical = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains("@book{unsafe_key,", canonical.Content, StringComparison.Ordinal);
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV217");
        Assert.False(BibliographyDocument.Parse(canonical.Content, BibliographyFormat.BibLatex).HasErrors);
    }

    [Fact]
    public void NBIB_journal_abbreviation_tag_survives_a_strict_edit() {
        const string source = "PMID- 1\nTA  - J Abbrev\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains("TA  - J Abbrev", written.Content, StringComparison.Ordinal);
        Assert.DoesNotContain("JT  - J Abbrev", written.Content, StringComparison.Ordinal);
        Assert.Equal(written.Content, BibliographyDocument.Parse(written.Content, BibliographyFormat.Nbib).Document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }).Content);
    }

    [Fact]
    public void Qualified_NBIB_ISSN_survives_a_strict_edit() {
        const string source = "PMID- 1\nIS  - 1234-5678 (Electronic)\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;

        Assert.Equal("1234-5678 (Electronic)", Assert.Single(document.Items[0].Identifiers, identifier => identifier.Scheme == "ISSN").Value);
        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        Assert.Contains("IS  - 1234-5678 (Electronic)", written.Content, StringComparison.Ordinal);
    }

    [Fact]
    public void NBIB_MeSH_headings_remain_distinct_from_other_terms() {
        const string source = "PMID- 1\nMH  - Controlled heading\nOT  - Author keyword\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;

        Assert.Equal("Author keyword", Assert.Single(document.Items[0].Keywords));
        Assert.Equal("Controlled heading", Assert.Single(document.Items[0].NativeFields, field => field.Name == "MH").Value);
        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        Assert.Contains("MH  - Controlled heading", written.Content, StringComparison.Ordinal);
        Assert.Contains("OT  - Author keyword", written.Content, StringComparison.Ordinal);
    }

    [Fact]
    public void Compact_NBIB_authors_parse_in_family_first_order() {
        const string source = "PMID- 1\nAU  - Smith J\n";

        BibliographyName name = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document.Items[0].Contributors).Name;

        Assert.Equal("Smith", name.Family);
        Assert.Equal("J", name.Given);
    }

    [Fact]
    public void RIS_contributor_roles_preserve_their_cross_role_source_order() {
        const string source = "TY  - BOOK\nID  - x\nAU  - Alpha, Alice\nED  - Editor, Erin\nAU  - Beta, Bob\nER  -\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        int firstAuthor = written.Content.IndexOf("AU  - Alpha, Alice", StringComparison.Ordinal);
        int editor = written.Content.IndexOf("ED  - Editor, Erin", StringComparison.Ordinal);
        int secondAuthor = written.Content.IndexOf("AU  - Beta, Bob", StringComparison.Ordinal);

        Assert.True(firstAuthor < editor && editor < secondAuthor);
        Assert.Equal(new[] { BibliographyContributorRole.Author, BibliographyContributorRole.Editor, BibliographyContributorRole.Author },
            BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items[0].Contributors.Select(contributor => contributor.Role));
    }

    [Theory]
    [InlineData("A1", "Y1")]
    [InlineData("AU", "DA")]
    public void RIS_contributor_and_date_aliases_survive_a_strict_edit(string contributorTag, string dateTag) {
        string source = $"TY  - BOOK\nID  - x\n{contributorTag}  - Alpha, Alice\n{dateTag}  - 2024\nER  -\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains($"{contributorTag}  - Alpha, Alice", written.Content, StringComparison.Ordinal);
        Assert.Contains($"{dateTag}  - 2024", written.Content, StringComparison.Ordinal);
    }
}
