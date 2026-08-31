using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave45RegressionTests {
    [Fact]
    public void Malformed_CSL_location_scans_observe_cancellation() {
        string source = new string('a', 64 * 1024 * 1024);
        var exception = new JsonException("Malformed JSON.", "$", 0, source.Length - 1);
        using var cancellation = new CancellationTokenSource();
        var cancellationThread = new Thread(() => { Thread.Sleep(1); cancellation.Cancel(); });
        cancellationThread.Start();

        try {
            Assert.Throws<OperationCanceledException>(() =>
                CslJsonCodec.GetJsonLocation(source, exception, cancellation.Token, out _, out _, out _));
        } finally {
            cancellationThread.Join();
        }
    }

    [Fact]
    public void EndNote_additional_URL_inspection_observes_cancellation() {
        var item = new BibliographyItem { Url = string.Empty };
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            BibliographyConversionInspector.InspectProperties(item, BibliographyFormat.EndNoteXml, new BibliographyConversionReport(), cancellation.Token));
    }

    [Theory]
    [InlineData("title")]
    [InlineData("contributor")]
    [InlineData("date")]
    [InlineData("native-value")]
    [InlineData("native-name")]
    [InlineData("type")]
    public void Invalid_UTF16_is_diagnosed_and_safely_replaced_in_CSL_output(string owner) {
        const string invalid = "before\uD800after";
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "item", Type = BibliographyItemType.Book, Title = "Valid" };
        document.Items.Add(item);
        switch (owner) {
            case "title": item.Title = invalid; break;
            case "contributor": item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, new BibliographyName { Family = invalid })); break;
            case "date": item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Literal = invalid }); break;
            case "native-value": item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.CslJson, "custom", invalid)); break;
            case "native-name": item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.CslJson, invalid, "value")); break;
            case "type": item.Type = BibliographyItemType.Unknown; item.NativeType = invalid; break;
        }

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));
        BibliographyWriteResult permissive = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical });

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV250");
        Assert.Contains(permissive.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV250");
        Assert.Contains("\\uFFFD", permissive.Content, StringComparison.OrdinalIgnoreCase);
        using JsonDocument reopened = JsonDocument.Parse(permissive.Content);
        Assert.Equal(JsonValueKind.Array, reopened.RootElement.ValueKind);
    }
}
