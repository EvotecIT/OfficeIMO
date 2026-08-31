namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave51RegressionTests {
    [Fact]
    public void Bib_surplus_name_segments_retain_the_complete_native_field_and_block_strict_output() {
        const string source = "@book{x, title={Before}, author={Last, Jr, First, Extra}}";
        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.BibLatex);
        BibliographyItem item = Assert.Single(read.Document.Items);
        item.Title = "After";

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            read.Document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBBIB011" && diagnostic.Field == "author");
        Assert.Equal("Last, Jr, First, Extra", Assert.Single(item.NativeFields, field => field.Name == "author").Value);
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV222" || diagnostic.Code == "BIBCONV119");
    }

    [Theory]
    [InlineData("null")]
    [InlineData("[]")]
    [InlineData("\"value\"")]
    public void CSL_non_item_root_array_elements_count_toward_the_value_limit(string value) {
        string source = "[" + value + "," + value + "]";

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, new BibliographyReadOptions {
            MaximumValueCount = 1
        });

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris, false)]
    [InlineData(BibliographyFormat.Ris, true)]
    [InlineData(BibliographyFormat.Nbib, false)]
    [InlineData(BibliographyFormat.Nbib, true)]
    public void Tagged_writers_observe_cancellation_inside_large_typed_and_native_values(BibliographyFormat format, bool native) {
        string value = new string('x', 64 * 1024 * 1024);
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = native ? "Title" : value };
        if (native) item.NativeFields.Add(new BibliographyNativeField(format, "ZZ", value));
        document.Items.Add(item);
        using var cancellation = new CancellationTokenSource();
        var cancellationThread = new Thread(() => { Thread.Sleep(1); cancellation.Cancel(); });
        cancellationThread.Start();

        try {
            Assert.Throws<OperationCanceledException>(() => {
                var options = new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical };
                var report = new BibliographyConversionReport();
                if (format == BibliographyFormat.Ris) TaggedCodec.WriteRis(document, options, report, cancellation.Token);
                else TaggedCodec.WriteNbib(document, options, report, cancellation.Token);
            });
        } finally {
            cancellationThread.Join();
        }
    }

    [Theory]
    [InlineData("AU")]
    [InlineData("A1")]
    [InlineData("ED")]
    [InlineData("A2")]
    [InlineData("PY")]
    [InlineData("Y1")]
    [InlineData("DA")]
    [InlineData("Y2")]
    public void Blank_RIS_native_name_and_date_tags_block_strict_output(string tag) {
        var document = new BibliographyDocument(BibliographyFormat.Ris);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Title" };
        item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.Ris, tag, string.Empty));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV122" && diagnostic.Field == tag);
    }
}
