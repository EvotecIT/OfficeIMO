namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave43RegressionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EndNote_root_prefixes_survive_strict_canonical_edits(bool recordsRoot) {
        string source = recordsRoot
            ? "<e:records xmlns:e=\"urn:endnote\"><e:record><e:rec-number>1</e:rec-number><e:ref-type name=\"Book\">6</e:ref-type><e:titles><e:title>Before</e:title></e:titles></e:record></e:records>"
            : "<e:xml xmlns:e=\"urn:endnote\"><e:records><e:record><e:rec-number>1</e:rec-number><e:ref-type name=\"Book\">6</e:ref-type><e:titles><e:title>Before</e:title></e:titles></e:record></e:records></e:xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        string rootName = recordsRoot ? "records" : "xml";
        string originalCarrier = Assert.Single(document.NativeEntries, entry => entry.Kind == "attributes" && entry.Name == rootName).Value;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDocument reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document;

        Assert.Contains("<e:" + rootName, written.Content, StringComparison.Ordinal);
        Assert.Equal(originalCarrier, Assert.Single(reopened.NativeEntries, entry => entry.Kind == "attributes" && entry.Name == rootName).Value);
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    [InlineData(BibliographyFormat.Ris)]
    [InlineData(BibliographyFormat.Nbib)]
    [InlineData(BibliographyFormat.EndNoteXml)]
    public void Positive_source_years_above_9999_round_trip(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Long year" };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Year = 10000 });
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items).Dates);

        Assert.Equal(10000, reopened.Year);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV218");
    }

    [Fact]
    public void Large_undelimited_Bib_names_do_not_allocate_per_character_substrings() {
        string name = new string('A', 512 * 1024);
        string source = "@book{x,author={" + name + "}}";
        var options = new BibliographyReadOptions { MaximumValueLength = name.Length + 1 };
        BibliographyDocument.Parse("@book{x,author={Doe}}", BibliographyFormat.BibLatex);
#if NET472
        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex, options);
#else
        long before = GC.GetAllocatedBytesForCurrentThread();
        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex, options);
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
#endif

        Assert.Equal(name, Assert.Single(Assert.Single(read.Document.Items).Contributors).Name.Family);
#if !NET472
        Assert.True(allocated < 12 * 1024 * 1024, $"Undelimited Bib name parsing allocated {allocated:N0} bytes.");
#endif
    }

    [Fact]
    public void Case_variant_EndNote_type_attributes_are_retained_as_unsupported_native_XML() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type NAME=\"Custom\">17</ref-type><isbn TYPE=\"ISBN\">9781402894626</isbn></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        BibliographyItem item = Assert.Single(document.Items);

        Assert.Contains(item.NativeFields, field => field.Name == "ref-type" && field.RawValue!.Contains("NAME=", StringComparison.Ordinal));
        Assert.Contains(item.NativeFields, field => field.Name == "isbn" && field.RawValue!.Contains("TYPE=", StringComparison.Ordinal));
        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV123" && diagnostic.Field == "ref-type");
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV123" && diagnostic.Field == "isbn");
    }

    [Fact]
    public void NBIB_publication_types_keep_their_order_among_native_extensions() {
        const string source = "ZZ  - first\nPT  - Book\nZY  - second\nPT  - Custom Type\nZX  - third\nPMID- 1\nTI  - Ordered\n";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;
        BibliographyItem item = Assert.Single(document.Items);
        string[] originalOrder = item.NativeFields.Select(field => field.Name + "=" + field.Value).ToArray();

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Nbib).Document.Items);

        Assert.Equal(originalOrder, reopened.NativeFields.Select(field => field.Name + "=" + field.Value));
    }

    [Fact]
    public void Declared_EndNote_encoding_loss_returns_replacement_bytes_in_permissive_mode() {
        const string source = "<?xml version=\"1.0\" encoding=\"us-ascii\"?><xml><records><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Łódź</title></titles></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;

        BibliographyWriteResult written = document.Write();

        Assert.True(written.UsedOriginalSource);
        Assert.Contains(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV220" && diagnostic.Field == "encoding");
        Assert.Contains("?", Encoding.ASCII.GetString(written.Bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void Strict_encoding_instances_return_replacement_bytes_after_permissive_diagnostics() {
        BibliographyDocument document = BibliographyDocument.Parse("@book{x,title={Łódź}}", BibliographyFormat.BibLatex).Document;
        Encoding strictAscii = Encoding.GetEncoding(Encoding.ASCII.CodePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, Encoding = strictAscii });

        Assert.Contains(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV220" && diagnostic.Field == "encoding");
        Assert.Contains("?", Encoding.ASCII.GetString(written.Bytes), StringComparison.Ordinal);
    }
}
