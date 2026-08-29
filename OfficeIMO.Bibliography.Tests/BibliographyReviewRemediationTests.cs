using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewRemediationTests {
    [Fact]
    public void Tagged_parser_stops_after_the_configured_diagnostic_limit() {
        string source = string.Join("\n", Enumerable.Repeat("malformed", 100));

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.Ris, new BibliographyReadOptions { MaximumDiagnosticCount = 2 });

        Assert.True(read.HasErrors);
        Assert.Equal(3, read.Diagnostics.Count);
        Assert.Equal("BIBLIM002", read.Diagnostics[2].Code);
    }

    [Fact]
    public void Edited_raw_backed_CSL_fields_are_written_from_their_public_values() {
        const string source = "[{\"id\":\"x\",\"type\":\"book\",\"x-item\":{\"enabled\":true},\"author\":[{\"literal\":\"Team\",\"x-name\":{\"rank\":1}}],\"issued\":{\"literal\":\"soon\",\"x-date\":{\"certainty\":\"low\"}}}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        BibliographyItem item = document.Items[0];
        item.NativeFields[0].Value = "item-edited";
        item.Contributors[0].Name.NativeFields[0].Value = "name-edited";
        item.Dates[0].NativeFields[0].Value = "date-edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        using JsonDocument json = JsonDocument.Parse(written.Content);
        JsonElement root = json.RootElement[0];
        Assert.Equal("item-edited", root.GetProperty("x-item").GetString());
        Assert.Equal("name-edited", root.GetProperty("author")[0].GetProperty("x-name").GetString());
        Assert.Equal("date-edited", root.GetProperty("issued").GetProperty("x-date").GetString());
    }

    [Fact]
    public void Bib_writer_omits_unsafe_and_typed_field_identifier_schemes() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Safe output" };
        item.Identifiers.Add(new BibliographyIdentifier("custom id", "unsafe"));
        item.Identifiers.Add(new BibliographyIdentifier("title", "collision"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.False(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibLatex).HasErrors);
        Assert.DoesNotContain("custom id", written.Content, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(2, strict.Report.Diagnostics.Count(diagnostic => diagnostic.Code == "BIBCONV129"));
    }

    [Fact]
    public void EndNote_aggregate_element_values_observe_the_value_length_limit() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>abc<empty />def</title></titles></record></records></xml>";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml, new BibliographyReadOptions { MaximumValueLength = 5 });

        Assert.True(read.HasErrors);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Fact]
    public void EndNote_attribute_values_observe_the_value_length_limit() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Oversized\">6</ref-type></record></records></xml>";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml, new BibliographyReadOptions { MaximumValueLength = 5 });

        Assert.True(read.HasErrors);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Fact]
    public void Bib_native_entries_observe_the_value_count_limit() {
        const string source = "@comment{a}\n@comment{b}\n@comment{c}";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex, new BibliographyReadOptions { MaximumValueCount = 2 });

        Assert.True(read.HasErrors);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Fact]
    public void Strict_canonical_write_rejects_partially_recovered_source() {
        BibliographyDocument document = BibliographyDocument.Parse("@book{a,title={A}}\n@book", BibliographyFormat.BibLatex).Document;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV222");
    }

    [Theory]
    [InlineData("ignored\n@book{x,title={A}}", BibliographyFormat.BibLatex)]
    [InlineData("malformed\nTY  - BOOK\nID  - x\nER  -\n", BibliographyFormat.Ris)]
    [InlineData("[1,{\"id\":\"x\",\"type\":\"book\"}]", BibliographyFormat.CslJson)]
    public void Strict_canonical_write_rejects_ignored_source_fragments(string source, BibliographyFormat format) {
        BibliographyDocument document = BibliographyDocument.Parse(source, format).Document;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV222");
    }

    [Fact]
    public void Bib_literal_contributor_with_and_reopens_as_one_name() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Corporate author" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Literal = "Research and Development Team" }));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyContributor contributor = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibLatex).Document.Items[0].Contributors);

        Assert.Contains("{{Research and Development Team}}", written.Content, StringComparison.Ordinal);
        Assert.Equal("Research and Development Team", contributor.Name.Literal);
    }

    [Fact]
    public void Csl_keyword_delimiters_survive_strict_canonical_round_trip() {
        BibliographyDocument document = BibliographyDocument.Parse("[{\"id\":\"x\",\"type\":\"book\",\"keyword\":\"alpha, beta; gamma\"}]", BibliographyFormat.CslJson).Document;

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items);

        Assert.Equal("alpha, beta; gamma", Assert.Single(reopened.Keywords));
    }

    [Fact]
    public void Empty_CSL_contributor_objects_count_toward_the_value_limit() {
        const string source = "[{\"id\":\"x\",\"type\":\"book\",\"author\":[{},{},{}]}]";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, new BibliographyReadOptions { MaximumValueCount = 4 });

        Assert.True(read.HasErrors);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Fact]
    public void Csl_syntax_diagnostics_use_absolute_UTF16_locations() {
        const string source = "[\n{\"title\":\"Ł\",?}]";
        int expectedOffset = source.IndexOf('?');
        int lineStart = source.IndexOf('\n') + 1;

        BibliographyDiagnostic diagnostic = Assert.Single(BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Diagnostics);

        Assert.Equal("BIBCSL002", diagnostic.Code);
        Assert.Equal(expectedOffset, diagnostic.Offset);
        Assert.Equal(2, diagnostic.Line);
        Assert.Equal(expectedOffset - lineStart + 1, diagnostic.Column);
    }

    [Fact]
    public void Edited_raw_backed_EndNote_field_is_written_as_a_safe_element() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><custom><nested>original</nested></custom></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].NativeFields.Single(field => field.Name == "custom").Value = "edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyNativeField reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items[0].NativeFields.Single(field => field.Name == "custom");

        Assert.Equal("edited", reopened.Value);
        Assert.Contains("<custom>edited</custom>", written.Content, StringComparison.Ordinal);
    }
}
