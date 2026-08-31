namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewClosureTests {
    [Fact]
    public void Bib_macro_expansion_observes_the_value_length_limit() {
        const string source = "@string{x={123456}}\n@book{a,title=x # x}";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex, new BibliographyReadOptions { MaximumValueLength = 10 });

        Assert.True(read.HasErrors);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Fact]
    public void Nbib_strict_write_rejects_a_citation_key_distinct_from_PMID() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "smith2025", Type = BibliographyItemType.ArticleJournal, Title = "Key contract" };
        item.Identifiers.Add(new BibliographyIdentifier("PMID", "12345"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.Nbib, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV224" && diagnostic.Field == "key");
    }

    [Fact]
    public void Nbib_strict_write_rejects_using_a_key_as_a_missing_PMID() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        document.Items.Add(new BibliographyItem { Key = "smith2025", Type = BibliographyItemType.ArticleJournal, Title = "Identifier contract" });

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.Nbib, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV223" && diagnostic.Field == "identifiers.PMID");
    }

    [Fact]
    public void Csl_strict_write_rejects_unsupported_identifier_schemes() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Identifiers" };
        item.Identifiers.Add(new BibliographyIdentifier("arXiv", "2401.00001"));
        document.Items.Add(item);

        BibliographyWriteResult canonical = document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.CslJson, Mode = BibliographyWriterMode.Canonical });
        BibliographyConversionLossException strict = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.CslJson, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.DoesNotContain("ARXIV", canonical.Content, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(strict.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV225" && diagnostic.Field == "identifiers.arXiv");
    }

    [Fact]
    public void Bib_keyword_delimiters_survive_strict_canonical_round_trip() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Keyword contract" };
        item.Keywords.Add("alpha, beta; gamma");
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibLatex).Document.Items);

        Assert.Contains("keywords = {{alpha, beta; gamma}}", written.Content, StringComparison.Ordinal);
        Assert.Equal("alpha, beta; gamma", Assert.Single(reopened.Keywords));
    }

    [Fact]
    public void Multiple_Bib_keywords_survive_strict_canonical_round_trip() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Keywords" };
        item.Keywords.Add("alpha, beta");
        item.Keywords.Add("gamma");
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibLatex).Document.Items);

        Assert.Equal(item.Keywords, reopened.Keywords);
    }

    [Fact]
    public void Recognized_native_Bib_type_survives_an_edit() {
        BibliographyDocument document = BibliographyDocument.Parse("@mastersthesis{x,title={Original}}", BibliographyFormat.BibLatex).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibLatex).Document.Items);

        Assert.StartsWith("@mastersthesis", written.Content, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(BibliographyItemType.Thesis, reopened.Type);
        Assert.Equal("mastersthesis", reopened.NativeType);
    }

    [Theory]
    [InlineData(BibliographyItemType.ArticleJournal)]
    [InlineData(BibliographyItemType.ArticleMagazine)]
    [InlineData(BibliographyItemType.ArticleNewspaper)]
    [InlineData(BibliographyItemType.Book)]
    [InlineData(BibliographyItemType.Chapter)]
    [InlineData(BibliographyItemType.PaperConference)]
    [InlineData(BibliographyItemType.Report)]
    [InlineData(BibliographyItemType.Thesis)]
    [InlineData(BibliographyItemType.WebPage)]
    [InlineData(BibliographyItemType.Dataset)]
    [InlineData(BibliographyItemType.Software)]
    [InlineData(BibliographyItemType.Patent)]
    [InlineData(BibliographyItemType.PersonalCommunication)]
    public void Exactly_supported_RIS_types_reopen_as_the_same_type(BibliographyItemType type) {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        document.Items.Add(new BibliographyItem { Key = "x", Type = type, Title = "Type contract" });

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.Equal(type, reopened.Type);
    }

    [Fact]
    public void Three_part_Bib_name_uses_family_suffix_given_order() {
        BibliographyDocument document = BibliographyDocument.Parse("@book{x,title={Names},author={Smith, Jr., John}}", BibliographyFormat.BibLatex).Document;
        BibliographyName name = Assert.Single(document.Items[0].Contributors).Name;

        Assert.Equal("Smith", name.Family);
        Assert.Equal("Jr.", name.Suffix);
        Assert.Equal("John", name.Given);

        document.Items[0].Title = "Edited";
        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        Assert.Contains("Smith, Jr., John", written.Content, StringComparison.Ordinal);
    }

    [Fact]
    public void Additional_EndNote_related_URLs_survive_an_edit() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Original</title></titles><urls><related-urls><url>https://one.example</url><url>https://two.example</url></related-urls></urls></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal("https://one.example", reopened.Url);
        Assert.Equal("https://two.example", Assert.Single(reopened.NativeFields, field => field.Name == "url").Value);
    }

    [Theory]
    [InlineData("institution")]
    [InlineData("organization")]
    public void Bib_publisher_semantics_survive_an_edit(string fieldName) {
        string source = $"@techreport{{x,title={{Original}},{fieldName}={{Research Lab}}}}";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibLatex).Document.Items);

        Assert.Contains(fieldName + " = {Research Lab}", written.Content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("publisher = {Research Lab}", written.Content, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("Research Lab", reopened.Publisher);
    }

    [Theory]
    [InlineData("journaltitle", "Journal of Tests")]
    [InlineData("location", "Warsaw")]
    [InlineData("issue", "special")]
    [InlineData("eid", "e123")]
    [InlineData("langid", "english")]
    public void Bib_typed_aliases_keep_their_source_field_names(string fieldName, string value) {
        string source = $"@article{{x,title={{Original}},{fieldName}={{{value}}}}}";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains(fieldName + " = {" + value + "}", written.Content, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RIS_accession_without_ID_survives_as_key_and_identifier() {
        const string source = "TY  - BOOK\nAN  - 12345\nTI  - Original\nER  -\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;
        BibliographyItem parsed = Assert.Single(document.Items);

        Assert.Equal("12345", parsed.Key);
        Assert.Equal("12345", parsed.GetIdentifier("accession"));

        parsed.Title = "Edited";
        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);
        Assert.Equal("12345", reopened.Key);
        Assert.Equal("12345", reopened.GetIdentifier("accession"));
    }

    [Fact]
    public void RIS_unknown_serial_number_stays_an_SN_identifier() {
        const string source = "TY  - BOOK\nID  - x\nSN  - ABC-123\nER  -\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Ris).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.Contains("SN  - ABC-123", written.Content, StringComparison.Ordinal);
        Assert.Equal("ABC-123", reopened.GetIdentifier("SN"));
    }

    [Fact]
    public void BibLaTex_typed_aliases_are_mapped_to_classic_BibTeX_fields() {
        const string source = "@article{x,title={Aliases},journaltitle={Journal},location={Warsaw},issue={special},eid={e123},langid={english}}";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex).Document;
        document.Items[0].Title = "Edited";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.BibTex, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains("journal = {Journal}", written.Content, StringComparison.Ordinal);
        Assert.Contains("address = {Warsaw}", written.Content, StringComparison.Ordinal);
        Assert.Contains("number = {special}", written.Content, StringComparison.Ordinal);
        Assert.Contains("pages = {e123}", written.Content, StringComparison.Ordinal);
        Assert.Contains("language = {english}", written.Content, StringComparison.Ordinal);
        Assert.DoesNotContain("journaltitle", written.Content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("location =", written.Content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("langid", written.Content, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void BibLaTex_only_native_type_blocks_strict_classic_BibTeX_conversion() {
        BibliographyDocument document = BibliographyDocument.Parse("@thesis{x,title={Thesis}}", BibliographyFormat.BibLatex).Document;

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Format = BibliographyFormat.BibTex, Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV200" && diagnostic.Field == "type");
    }

    [Fact]
    public void Standard_Bib_name_particles_and_family_only_names_reopen_structurally() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Names" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Given = "Ludwig", Family = "Beethoven", NonDroppingParticle = "van", DroppingParticle = "de" }));
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = "Plato" }));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyContributor[] reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.BibLatex).Document.Items[0].Contributors.ToArray();

        Assert.Equal("van", reopened[0].Name.NonDroppingParticle);
        Assert.Equal("Beethoven", reopened[0].Name.Family);
        Assert.Equal("Ludwig", reopened[0].Name.Given);
        Assert.Equal("de", reopened[0].Name.DroppingParticle);
        Assert.Equal("Plato", reopened[1].Name.Family);
        Assert.Null(reopened[1].Name.Literal);
    }

    [Fact]
    public void Nonstandard_Bib_particle_case_blocks_strict_round_trip() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Names" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Given = "Ludwig", Family = "Beethoven", NonDroppingParticle = "Van" }));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV226" && diagnostic.Field == "contributors");
    }

    [Fact]
    public void Ambiguous_lowercase_Bib_family_blocks_strict_round_trip() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Names" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Given = "Jane", Family = "van Example" }));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV226");
    }

    [Fact]
    public void RIS_accession_with_a_colon_reopens_as_an_accession() {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Accession" };
        item.Identifiers.Add(new BibliographyIdentifier("accession", "archive:123"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Ris).Document.Items);

        Assert.Contains("AN  - accession:archive:123", written.Content, StringComparison.Ordinal);
        Assert.Equal("archive:123", reopened.GetIdentifier("accession"));
    }

    [Fact]
    public void Multiple_PMIDs_block_strict_NBIB_output() {
        var document = new BibliographyDocument(BibliographyFormat.Nbib);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.ArticleJournal, Title = "PMIDs" };
        item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        item.Identifiers.Add(new BibliographyIdentifier("PMID", "2"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() => document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV227" && diagnostic.Field == "identifiers.PMID");
    }

    [Theory]
    [InlineData("journal", "journaltitle")]
    [InlineData("publisher", "institution")]
    [InlineData("address", "location")]
    [InlineData("number", "issue")]
    [InlineData("pages", "eid")]
    [InlineData("language", "langid")]
    public void Distinct_Bib_alias_values_remain_stable_across_reopen(string firstField, string secondField) {
        string source = $"@article{{x,title={{Original}},{firstField}={{First}},{secondField}={{Second}}}}";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex).Document;
        document.Items[0].Title = "Edited";
        var options = new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true };

        BibliographyWriteResult first = document.Write(options);
        BibliographyDocument reopened = BibliographyDocument.Parse(first.Content, BibliographyFormat.BibLatex).Document;
        BibliographyWriteResult second = reopened.Write(options);

        Assert.Equal(first.Content, second.Content);
        Assert.Contains(firstField + " = {First}", first.Content, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(secondField + " = {Second}", first.Content, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("@comment{12345678901}")]
    [InlineData("@book{x,title=12345678901}")]
    [InlineData("%12345678901\n@book{x,title={ok}}")]
    public void Bib_sibling_value_paths_enforce_the_limit_before_materializing_output(string source) {
        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex, new BibliographyReadOptions { MaximumValueLength = 10 });

        Assert.True(read.HasErrors);
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }
}
