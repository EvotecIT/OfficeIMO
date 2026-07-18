using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfGeneratedLayerTests {
    [Fact]
    public void Layer_GeneratesOptionalContentResourcesAndRoundTripsConfiguration() {
        byte[] bytes = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Always visible"))
            .Layer("Review notes", layer => {
                layer.Paragraph(paragraph => paragraph.Text("Layer page one"));
                layer.PageBreak();
                layer.Paragraph(paragraph => paragraph.Text("Layer page two"));
            }, new PdfLayerOptions {
                InitiallyVisible = false,
                Locked = true,
                VisibleWhenPrinting = false
            })
            .ToBytes();

        PdfReadDocument read = PdfReadDocument.Open(bytes);
        PdfOptionalContentGroup group = Assert.Single(read.OptionalContent!.Groups);
        Assert.Equal("Review notes", group.Name);
        Assert.False(group.IsInitiallyVisible);
        Assert.True(group.IsLocked);
        Assert.Equal(2, read.Pages.Count);

        string raw = PdfEncoding.Latin1GetString(bytes);
        Assert.StartsWith("%PDF-1.5", raw, StringComparison.Ordinal);
        Assert.Equal(2, raw.Split(new[] { "/OC /OC1 BDC" }, StringSplitOptions.None).Length - 1);
        Assert.Equal(2, raw.Split(new[] { "/Properties << /OC1" }, StringSplitOptions.None).Length - 1);
        Assert.Contains("/PrintState /OFF", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Layer_NestsWithIndependentOptionalContentGroups() {
        byte[] bytes = PdfDocument.Create()
            .Layer("Outer", outer => outer
                .Layer("Inner", inner => inner.Paragraph(paragraph => paragraph.Text("Nested"))))
            .ToBytes();

        PdfOptionalContentProperties optionalContent = PdfReadDocument.Open(bytes).OptionalContent!;
        Assert.Equal(new[] { "Outer", "Inner" }, optionalContent.Groups.Select(group => group.Name));
    }
}
