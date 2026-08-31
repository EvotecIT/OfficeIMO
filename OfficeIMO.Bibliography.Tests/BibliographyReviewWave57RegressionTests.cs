using System.Xml;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave57RegressionTests {
    [Theory]
    [InlineData("\t", "\n")]
    [InlineData("\n", "\t")]
    [InlineData("  ", "\r\n")]
    public void Bib_name_lists_accept_any_whitespace_around_and(string before, string after) {
        string source = "@book{one, author={Doe, John" + before + "and" + after + "Smith, Jane}}";

        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.BibTex).Document;

        Assert.Collection(document.Items[0].Contributors,
            contributor => { Assert.Equal("Doe", contributor.Name.Family); Assert.Equal("John", contributor.Name.Given); },
            contributor => { Assert.Equal("Smith", contributor.Name.Family); Assert.Equal("Jane", contributor.Name.Given); });
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Structured_names_with_a_suffix_and_missing_given_name_are_lossy(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "one", Type = BibliographyItemType.Book, Title = "Title" };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author,
            new BibliographyName { Family = "Doe", Suffix = "Jr" }));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV244" && diagnostic.Field == "contributors");
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris, "TY  - JOUR\nID  - one\nDO  -\n      10.1/example\nER  -\n", "DOI", "10.1/example")]
    [InlineData(BibliographyFormat.Ris, "TY  - JOUR\nID  - one\nSN  -\n      1234-5678\nER  -\n", "ISSN", "1234-5678")]
    [InlineData(BibliographyFormat.Ris, "TY  - JOUR\nAN  -\n      local:42\nER  -\n", "local", "42")]
    [InlineData(BibliographyFormat.Nbib, "PMID- 1\nIS  -\n      1234-5678\n", "ISSN", "1234-5678")]
    [InlineData(BibliographyFormat.Nbib, "PMID- 1\nLID -\n      10.1/example [doi]\n", "doi", "10.1/example")]
    [InlineData(BibliographyFormat.Nbib, "PMID- 1\nAID -\n      S1234 [pii]\n", "pii", "S1234")]
    public void Tagged_identifier_continuations_promote_blank_native_tags(BibliographyFormat format, string source, string scheme, string value) {
        BibliographyDocument document = BibliographyDocument.Parse(source, format).Document;

        BibliographyIdentifier identifier = Assert.Single(document.Items[0].Identifiers,
            candidate => string.Equals(candidate.Scheme, scheme, StringComparison.OrdinalIgnoreCase) && candidate.Value == value);

        Assert.Equal(value, identifier.Value);
        Assert.DoesNotContain(document.Items[0].NativeFields,
            field => field.Format == format && (field.Name == "DO" || field.Name == "SN" || field.Name == "AN" || field.Name == "IS" || field.Name == "LID" || field.Name == "AID"));

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyIdentifier reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items[0].Identifiers,
            candidate => string.Equals(candidate.Scheme, scheme, StringComparison.OrdinalIgnoreCase) && candidate.Value == value);
        Assert.Equal(value, reopened.Value);
    }

    [Fact]
    public void EndNote_primary_parsing_observes_cancellation_inside_large_attribute_tokens() {
        string value = new string('x', 64 * 1024 * 1024);
        string source = "\uFEFF<xml><records metadata=\"" + value + "\"/></xml>";
        var options = new BibliographyReadOptions { MaximumInputCharacters = source.Length + 1, MaximumValueLength = source.Length + 1 };

        BibliographyCancellationTest.AssertObserved(token =>
            BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml, options, token));
    }

    [Fact]
    public void EndNote_offset_mapping_keeps_BOM_in_multiline_UTF16_offsets() {
        const string source = "\uFEFF<xml>\n  <record/>\n</xml>";
        var offsets = new EndNoteSourceOffsetMap(source, 1, CancellationToken.None);
        using var text = new EndNoteCancellableTextReader(source, CancellationToken.None, 1);
        using XmlReader reader = XmlReader.Create(text);
        while (reader.Read() && !(reader.NodeType == XmlNodeType.Element && reader.LocalName == "record")) { }

        Assert.Equal(source.IndexOf("<record", StringComparison.Ordinal), offsets.GetOffset((IXmlLineInfo)reader));
    }

    [Theory]
    [InlineData("valid")]
    [InlineData("invalid")]
    public void Bib_key_validation_and_normalization_observe_cancellation(string shape) {
        string key = new string(shape == "valid" ? 'a' : ' ', 64 * 1024 * 1024);
        if (shape == "invalid") key = "a" + key;
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        document.Items.Add(new BibliographyItem { Key = key, Type = BibliographyItemType.Book, Title = "Title" });

        BibliographyCancellationTest.AssertObserved(token =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, token));
    }
}
