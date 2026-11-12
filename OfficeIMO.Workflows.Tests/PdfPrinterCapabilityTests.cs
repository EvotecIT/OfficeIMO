namespace OfficeIMO.Workflows.Tests;

public sealed class PdfPaperSourceTests {
    [Theory]
    [InlineData("InputSlot/Media Source: *Auto Tray1 Tray2\nDuplex: *None DuplexNoTumble", "InputSlot=Auto", "InputSlot=Tray1", "InputSlot=Tray2")]
    [InlineData("media-source/Source: *auto main alternate", "media-source=auto", "media-source=main", "media-source=alternate")]
    public void CupsSourcesRetainOnlyTheSupportedOption(string output, string first, string second, string third) {
        Assert.Equal(new[] { first, second, third }, CupsPdfPrinter.ParsePaperSources(output).Select(source => source.Id));
    }

    [Fact]
    public void CupsSourcesDoNotInventTraysOrAcceptOptionFragments() {
        Assert.Empty(CupsPdfPrinter.ParsePaperSources("PageSize/Media Size: *A4 Letter"));
        Assert.Equal(new[] { "InputSlot=Tray1" }, CupsPdfPrinter.ParsePaperSources("InputSlot/Source: Tray1 Tray1 bad=value /evil").Select(source => source.Id));
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("tray\n-o")]
    public void DeliveryRejectsMalformedPaperSource(string source) {
        var options = new PdfPrintDeliveryOptions { PrinterName = "queue", PaperSourceId = source };
        Assert.Throws<ArgumentException>(() => options.Snapshot());
    }
}
