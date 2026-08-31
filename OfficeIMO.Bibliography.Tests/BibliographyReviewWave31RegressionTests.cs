using System.Xml;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave31RegressionTests {
    [Theory]
    [InlineData("utf16-le")]
    [InlineData("utf16-be")]
    [InlineData("utf32-le")]
    [InlineData("utf32-be")]
    public void Declaration_free_BOMless_UTF_XML_is_detected(string encodingName) {
        const string source = " \n<xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Detected</title></titles></record></records></xml>";
        Encoding encoding = CreateBomlessEncoding(encodingName);
        using var stream = new MemoryStream(encoding.GetBytes(source));

        BibliographyReadResult read = BibliographyDocument.Load(stream, BibliographyFormat.EndNoteXml);

        Assert.Equal("Detected", Assert.Single(read.Document.Items).Title);
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Severity == BibliographyDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Leading_whitespace_in_literal_names_is_diagnosed(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Literal = " Acme" }));
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV243" && diagnostic.Field == "contributors");
    }

    [Fact]
    public void EndNote_source_offsets_align_with_XmlReader_element_positions() {
        const string source = "<xml>\n  <record/>\n</xml>";
        var offsets = new EndNoteSourceOffsetMap(source, 0, CancellationToken.None);
        using var text = new StringReader(source);
        using XmlReader reader = XmlReader.Create(text);
        while (reader.Read() && !(reader.NodeType == XmlNodeType.Element && reader.LocalName == "record")) { }

        Assert.Equal(source.IndexOf("<record", StringComparison.Ordinal), offsets.GetOffset((IXmlLineInfo)reader));
    }

    [Fact]
    public void EndNote_PMID_scheme_loss_is_rejected_before_strict_output() {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Identifiers.Add(new BibliographyIdentifier("PMID", "123"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV204" && diagnostic.Field == "identifiers.PMID");
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris, BibliographyDateRole.Issued)]
    [InlineData(BibliographyFormat.Ris, BibliographyDateRole.Accessed)]
    [InlineData(BibliographyFormat.Nbib, BibliographyDateRole.Issued)]
    public void Tagged_date_ranges_survive_strict_canonical_output(BibliographyFormat format, BibliographyDateRole role) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.ArticleJournal };
        item.Dates.Add(new BibliographyDate { Role = role, Year = 2020, Month = 2, Day = 3, EndYear = 2021, EndMonth = 4, EndDay = 5 });
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items[0].Dates);

        Assert.Equal(role, reopened.Role);
        Assert.Equal(2020, reopened.Year);
        Assert.Equal(2, reopened.Month);
        Assert.Equal(3, reopened.Day);
        Assert.Equal(2021, reopened.EndYear);
        Assert.Equal(4, reopened.EndMonth);
        Assert.Equal(5, reopened.EndDay);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV219");
    }

    [Fact]
    public void CSL_identifier_value_whitespace_survives_strict_canonical_edits() {
        const string source = "[{\"id\":\"x\",\"type\":\"book\",\"title\":\"Before\",\"DOI\":\" 10.1000/example \"}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        Assert.Equal(" 10.1000/example ", Assert.Single(document.Items[0].Identifiers).Value);
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyIdentifier reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items[0].Identifiers);

        Assert.Equal(" 10.1000/example ", reopened.Value);
    }

    [Fact]
    public void Identifier_value_setters_preserve_nonempty_source_whitespace() {
        var identifier = new BibliographyIdentifier(" DOI ", " value ");
        var item = new BibliographyItem();
        item.Identifiers.Add(identifier);

        Assert.Equal("DOI", identifier.Scheme);
        Assert.Equal(" value ", identifier.Value);
        item.SetIdentifier(" doi ", " replacement ");
        Assert.Equal(" replacement ", identifier.Value);
    }

    [Fact]
    public void NBIB_qualified_identifier_whitespace_is_rejected_before_strict_output() {
        var document = new BibliographyDocument(BibliographyFormat.Nbib);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.ArticleJournal };
        item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        item.Identifiers.Add(new BibliographyIdentifier("DOI", "10.1000/example "));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV232" && diagnostic.Field == "identifiers.DOI");
    }

    private static Encoding CreateBomlessEncoding(string name) {
        switch (name) {
            case "utf16-le": return new UnicodeEncoding(false, false, true);
            case "utf16-be": return new UnicodeEncoding(true, false, true);
            case "utf32-le": return new UTF32Encoding(false, false, true);
            case "utf32-be": return new UTF32Encoding(true, false, true);
            default: throw new ArgumentOutOfRangeException(nameof(name));
        }
    }

}
