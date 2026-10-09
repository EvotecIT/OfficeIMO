using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFontFamilyTests {
    [Theory]
    [InlineData(PdfObjectSerializationMode.Buffered, PdfStandardFont.TimesRoman, false)]
    [InlineData(PdfObjectSerializationMode.ForwardOnly, PdfStandardFont.TimesRoman, false)]
    [InlineData(PdfObjectSerializationMode.Buffered, PdfStandardFont.HelveticaBold, false)]
    [InlineData(PdfObjectSerializationMode.ForwardOnly, PdfStandardFont.HelveticaBold, false)]
    [InlineData(PdfObjectSerializationMode.Buffered, PdfStandardFont.TimesRoman, true)]
    [InlineData(PdfObjectSerializationMode.ForwardOnly, PdfStandardFont.TimesRoman, true)]
    [InlineData(PdfObjectSerializationMode.Buffered, PdfStandardFont.HelveticaBold, true)]
    [InlineData(PdfObjectSerializationMode.ForwardOnly, PdfStandardFont.HelveticaBold, true)]
    public void AlternateStandardTextEmbedsOnlyUsedFontFaces(
        PdfObjectSerializationMode mode, PdfStandardFont bodyFont, bool regularHeader) {
        byte[] font = ManagedTextShapingTestAssets.CreateFont(' ', 'A', 'B');
        var options = new PdfOptions {
            FileVersion = PdfFileVersion.Pdf17,
            ObjectSerializationMode = mode,
            ShowHeader = regularHeader,
            HeaderFormat = "BA",
            HeaderFont = PdfStandardFont.Helvetica
        }.UseFontFamily(new PdfEmbeddedFontFamily("Resource Test", font, bold: font));
        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Font(bodyFont)
                .Bold(bodyFont == PdfStandardFont.HelveticaBold).Text("AB"))
            .ToBytes();

        using var independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        string text = string.Join(" ", independent.GetPages().Select(page => page.Text));
        Assert.Contains("AB", text, StringComparison.Ordinal);
        if (regularHeader) Assert.Contains("BA", text, StringComparison.Ordinal);

        PdfDocument document = PdfDocument.Load(bytes);
        PdfFontInventory inventory = document.Resources.Fonts();
        PdfRawDocumentView raw = document.Resources.RawStructure();
        Assert.False(raw.IsTruncated);
        int expectedEmbeddedFaces = (bodyFont == PdfStandardFont.HelveticaBold ? 1 : 0)
            + (regularHeader ? 1 : 0);
        Assert.Equal(expectedEmbeddedFaces, inventory.EmbeddedFontCount);
        Assert.Equal(expectedEmbeddedFaces, raw.Objects.Count(item =>
            item.Value.Entries.TryGetValue("Subtype", out PdfRawValue? subtype) && subtype.Text == "Type0"));
        Assert.Equal(expectedEmbeddedFaces, raw.Objects.Count(item => item.Value.Entries.ContainsKey("FontFile2")));
    }
    [Theory]
    [InlineData(PdfObjectSerializationMode.Buffered)]
    [InlineData(PdfObjectSerializationMode.ForwardOnly)]
    public void RegularFaceFirstUsedOnLaterPageRemainsEmbedded(PdfObjectSerializationMode mode) {
        byte[] font = ManagedTextShapingTestAssets.CreateFont(' ', 'A', 'B');
        var options = new PdfOptions {
            FileVersion = PdfFileVersion.Pdf17,
            ObjectSerializationMode = mode
        }.UseFontFamily(new PdfEmbeddedFontFamily("Later Resource", font, bold: font));
        byte[] bytes = PdfDocument.Create(options)
            .Page(page => page.Content(content => content.Item(item => item.Paragraph(p => p.Bold().Text("A")))))
            .Page(page => page.Content(content => content.Item(item => item.Paragraph(p => p.Text("B")))))
            .ToBytes();

        using var independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(2, independent.NumberOfPages);
        Assert.Equal("A", independent.GetPage(1).Text);
        Assert.Equal("B", independent.GetPage(2).Text);
        PdfDocument document = PdfDocument.Load(bytes);
        Assert.Equal(2, document.Resources.Fonts().EmbeddedFontCount);
        Assert.Equal(2, document.Resources.RawStructure().Objects.Count(item =>
            item.Value.Entries.ContainsKey("FontFile2")));
    }
}
