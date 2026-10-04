namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocReferenceContractTests {
    [Fact]
    public void CalloutsCorrelateAutomaticAndRepeatedMarkersAndKeepNativeNumbering() {
        const string source = "----\nFirst <.>\nRepeated <1>\nEscaped \\<2>\nSecond <.>\n----\n\n<.> First explanation\n<.> Second explanation\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source);
        AsciiDocCalloutCatalog catalog = AsciiDocCalloutCatalog.Create(document);
        Assert.Empty(catalog.Diagnostics);
        AsciiDocCalloutGroup group = Assert.Single(catalog.Groups);
        Assert.Equal(new[] { 1, 2 }, group.Items.Select(item => item.Number));
        Assert.Equal(2, group.Items[0].MarkerOffsets.Count);
        Assert.Single(group.Items[1].MarkerOffsets);
        group.Items[1].Explanation.Text = "Changed explanation";
        string written = document.ToAsciiDoc(new AsciiDocWriterOptions { Mode = AsciiDocWriterMode.Canonical });
        Assert.Contains("<.> Changed explanation", written);
        Assert.Equal(AsciiDocListKind.Callout, AsciiDocDocument.Parse(written).BlocksOfType<AsciiDocListBlock>().Single().Kind);
    }
    [Fact]
    public void MissingCalloutMarkersAndExplanationsAreReported() {
        AsciiDocDocument document = AsciiDocDocument.Parse("----\nCode <1>\n----\n\n<2> Unmatched explanation\n<2> Duplicate\n");
        AsciiDocCalloutCatalog catalog = AsciiDocCalloutCatalog.Create(document);
        Assert.Contains(catalog.Diagnostics, diagnostic => diagnostic.Code == "ADOCCO002");
        Assert.Contains(catalog.Diagnostics, diagnostic => diagnostic.Code == "ADOCCO003");
        Assert.Contains(catalog.Diagnostics, diagnostic => diagnostic.Code == "ADOCCO004");
    }
}
