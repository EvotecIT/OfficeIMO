using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave53RegressionTests {
    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Bib_whitespace_only_literal_names_round_trip(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Literal = "  " }));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyName reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items[0].Contributors).Name;

        Assert.Equal("  ", reopened.Literal);
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Bib_literal_names_with_structured_components_report_their_omission(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Literal = "Organization", Family = "Hidden" }));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV226" && diagnostic.Field == "contributors");
    }

    [Theory]
    [InlineData("typed")]
    [InlineData("native-field")]
    [InlineData("native-entry")]
    public void Bib_writers_observe_cancellation_inside_large_value_safety_scans(string owner) {
        string value = owner == "typed" ? new string('x', 64 * 1024 * 1024)
            : owner == "native-field" ? new string('\\', 64 * 1024 * 1024)
            : "}" + new string('x', 64 * 1024 * 1024);
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        if (owner == "native-entry") {
            document.NativeEntries.Add(new BibliographyNativeEntry(BibliographyFormat.BibLatex, "comment", value));
        } else {
            var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = owner == "typed" ? value : "Title" };
            if (owner == "native-field") item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.BibLatex, "custom", value));
            document.Items.Add(item);
        }
        using var cancellation = new CancellationTokenSource();
        var cancellationThread = new Thread(() => { Thread.Sleep(1); cancellation.Cancel(); });
        cancellationThread.Start();

        try {
            Assert.Throws<OperationCanceledException>(() =>
                BibCodec.Write(document, BibliographyFormat.BibLatex, new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, new BibliographyConversionReport(), cancellation.Token));
        } finally {
            cancellationThread.Join();
        }
    }

    [Theory]
    [InlineData("{\"a\":1,}", JsonValueKind.Object)]
    [InlineData("{\"a\":1/* retained comment */}", JsonValueKind.Object)]
    [InlineData("[1,2,]", JsonValueKind.Array)]
    public void CSL_permissive_native_aggregates_normalize_without_changing_shape(string rawValue, JsonValueKind expectedKind) {
        string source = "[{\"id\":\"1\",\"type\":\"book\",\"custom\":" + rawValue + ",\"title\":\"Before\"}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        using JsonDocument output = JsonDocument.Parse(written.Content);
        JsonElement custom = output.RootElement[0].GetProperty("custom");
        BibliographyNativeField reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items[0].NativeFields, field => field.Name == "custom");
        using JsonDocument reopenedRaw = JsonDocument.Parse(reopened.RawValue!);

        Assert.Equal(expectedKind, custom.ValueKind);
        Assert.Equal(expectedKind, reopenedRaw.RootElement.ValueKind);
        Assert.DoesNotContain("/*", written.Content, StringComparison.Ordinal);
    }
}
