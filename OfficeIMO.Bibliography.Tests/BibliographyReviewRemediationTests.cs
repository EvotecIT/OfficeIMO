using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewRemediationTests {
    [Theory]
    [InlineData(BibliographyFormat.Ris, "malformed", "\n", 2)]
    [InlineData(BibliographyFormat.BibTex, "outside@?", "", 3)]
    public void Parsers_stop_after_the_configured_diagnostic_limit(BibliographyFormat format, string fragment, string separator, int limit) {
        string source = string.Join(separator, Enumerable.Repeat(fragment, 100));

        BibliographyReadResult read = BibliographyDocument.Parse(source, format, new BibliographyReadOptions { MaximumDiagnosticCount = limit });

        Assert.True(read.HasErrors);
        Assert.Equal(limit + 1, read.Diagnostics.Count);
        Assert.Equal("BIBLIM002", read.Diagnostics[limit].Code);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Edited_raw_backed_CSL_fields_report_shape_flattening_and_use_public_values(bool includeNameAndDateOwners) {
        string source = includeNameAndDateOwners
            ? "[{\"id\":\"x\",\"type\":\"book\",\"x-item\":{\"enabled\":true},\"author\":[{\"literal\":\"Team\",\"x-name\":{\"rank\":1}}],\"issued\":{\"literal\":\"soon\",\"x-date\":{\"certainty\":\"low\"}}}]"
            : "[{\"id\":\"x\",\"type\":\"book\",\"custom\":{\"old\":true}}]";
        string itemField = includeNameAndDateOwners ? "x-item" : "custom";
        string itemValue = includeNameAndDateOwners ? "item-edited" : "flattened";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        BibliographyItem item = document.Items[0];
        item.NativeFields[0].Value = itemValue;
        if (includeNameAndDateOwners) {
            item.Contributors[0].Name.NativeFields[0].Value = "name-edited";
            item.Dates[0].NativeFields[0].Value = "date-edited";
        }

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        using JsonDocument json = JsonDocument.Parse(written.Content);
        JsonElement root = json.RootElement[0];
        Assert.Equal(itemValue, root.GetProperty(itemField).GetString());
        Assert.Contains(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV126" && diagnostic.Field == itemField);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV126" && diagnostic.Field == itemField);
        if (includeNameAndDateOwners) {
            Assert.Equal("name-edited", root.GetProperty("author")[0].GetProperty("x-name").GetString());
            Assert.Equal("date-edited", root.GetProperty("issued").GetProperty("x-date").GetString());
            Assert.Contains(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV127" && diagnostic.Field == "author.x-name");
            Assert.Contains(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV128" && diagnostic.Field == "issued.x-date");
            Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV127" && diagnostic.Field == "author.x-name");
            Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV128" && diagnostic.Field == "issued.x-date");
        }
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

    [Theory]
    [InlineData("<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>abc<empty />def</title></titles></record></records></xml>", 5, false)]
    [InlineData("<xml><records><record><rec-number>1</rec-number><ref-type name=\"Oversized\">6</ref-type></record></records></xml>", 5, false)]
    [InlineData("<xml><!--long--><records/></xml>", 3, false)]
    [InlineData("<xml><?review long?><records/></xml>", 3, false)]
    [InlineData("<records>oversized<record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles></record></records>", 4, true)]
    [InlineData("<records><record>oversized<rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles></record></records>", 4, true)]
    [InlineData("<xml><extension><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/></extension><records/></xml>", 80, false)]
    [InlineData("<xml xmlns:ext=\"urn:extension\"><records><record><rec-number>1</rec-number><ext:data><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/><x/></ext:data></record></records></xml>", 80, false)]
    public void EndNote_XML_values_observe_the_value_length_limit(string source, int limit, bool requireLengthMessage) {
        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml, new BibliographyReadOptions { MaximumValueLength = limit });

        Assert.True(read.HasErrors);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001"
            && (!requireLengthMessage || diagnostic.Message.Contains("value length", StringComparison.Ordinal)));
    }


    [Fact]
    public void EndNote_limit_diagnostics_report_character_offsets() {
        const string source = "<xml>\r\n  <records>\r\n    <record>\r\n      <rec-number>1</rec-number>\r\n      <ref-type name=\"Book\">6</ref-type>\r\n      <titles><title>abcdef</title></titles>\r\n    </record>\r\n  </records>\r\n</xml>";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml, new BibliographyReadOptions { MaximumValueLength = 5 });

        BibliographyDiagnostic diagnostic = Assert.Single(read.Diagnostics, static diagnostic => diagnostic.Code == "BIBLIM001");
        Assert.Equal(source.IndexOf("<title>", StringComparison.Ordinal), diagnostic.Offset);
    }

    [Fact]
    public void Bib_native_entries_observe_the_value_count_limit() {
        const string source = "@comment{a}\n@comment{b}\n@comment{c}";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex, new BibliographyReadOptions { MaximumValueCount = 2 });

        Assert.True(read.HasErrors);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Fact]
    public void Bib_keyword_items_observe_the_value_count_limit() {
        const string source = "@book{x,keywords={alpha,beta,gamma}}";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex, new BibliographyReadOptions { MaximumValueCount = 4 });

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
    public void Edited_structured_EndNote_field_reports_flattened_child_markup() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><custom><nested>original</nested></custom></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].NativeFields.Single(field => field.Name == "custom").Value = "edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        BibliographyNativeField reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items[0].NativeFields.Single(field => field.Name == "custom");

        Assert.Equal("edited", reopened.Value);
        Assert.Contains("<custom>edited</custom>", written.Content, StringComparison.Ordinal);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV234" && diagnostic.Field == "native.custom");
    }

    [Theory]
    [InlineData("TY  - BOOK\r\nTI  - abcde\r\nER  -\r\n", BibliographyFormat.Ris, "TI")]
    [InlineData("PMID- 1\r\nTI  - abcde\r\n", BibliographyFormat.Nbib, "TI")]
    public void Tagged_limit_diagnostics_report_character_offsets(string source, BibliographyFormat format, string failingTag) {
        BibliographyReadResult read = BibliographyDocument.Parse(source, format, new BibliographyReadOptions { MaximumValueLength = 4 });

        BibliographyDiagnostic diagnostic = Assert.Single(read.Diagnostics, static diagnostic => diagnostic.Code == "BIBLIM001");
        Assert.Equal(source.IndexOf(failingTag, StringComparison.Ordinal), diagnostic.Offset);
    }

    [Fact]
    public void CSL_writer_rejects_native_values_that_require_nondefault_read_depth() {
        string nested = new string('[', 1010) + "0" + new string(']', 1010);
        string source = "[{\"id\":\"x\",\"type\":\"book\",\"custom\":" + nested + "}]";
        var readOptions = new BibliographyReadOptions { MaximumNestingDepth = 1024 };
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, readOptions).Document;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });
        BibliographyReadResult reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson);

        Assert.False(reopened.HasErrors);
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV247" || diagnostic.Code == "BIBCONV126");
        Assert.Contains(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV126" && diagnostic.Field == "custom");
    }

    [Fact]
    public void CSL_parser_observes_cancellation_while_materializing_an_empty_root() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            CslJsonCodec.Parse("[]", new BibliographyReadOptions(), new System.Collections.Generic.List<BibliographyDiagnostic>(), out _, cancellation.Token));
    }
}
