using System.IO;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFontFamilyTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CoverageStableFallbackPreservesMissingPhysicalItalic(bool bold) {
        byte[] regular = File.ReadAllBytes(Assert.IsType<string>(PdfComplianceTestFonts.FindBundledTrueTypeFont()));
        var family = new PdfEmbeddedFontFamily("Coverage Proof", regular).CreateCoverageStableFallbackSnapshot();
        var options = new PdfOptions { CompressContentStreams = false }.RegisterNamedFontFamily(family);
        byte[] bytes = PdfDocument.Create(options).Paragraph(p => p.FontFamily(family.FamilyName)
            .Bold(bold).Italic(true).Text("Coverage italic")).ToBytes();
        Assert.Contains("1 0 0.333 1", Encoding.ASCII.GetString(bytes));
        Assert.Contains("Coverage italic", PdfReadDocument.Open(bytes).ExtractText());
    }

    [Theory]
    [InlineData(OfficeFontSlant.Normal, true)]
    [InlineData(OfficeFontSlant.Italic, false)]
    [InlineData(OfficeFontSlant.Oblique, false)]
    public void SelectedNumericFacePreservesPhysicalSlant(OfficeFontSlant slant, bool needsSynthesis) {
        string regularPath = Assert.IsType<string>(PdfComplianceTestFonts.FindBundledTrueTypeFont());
        byte[] bytes = File.ReadAllBytes(slant == OfficeFontSlant.Normal ? regularPath : regularPath.Replace("-Regular.ttf", "-Italic.ttf"));
        var descriptor = new OfficeFontFaceDescriptor(450, 100, slant);
        var faces = new OfficeFontFaceCollection().Add("Numeric Proof", bytes, descriptor);
        Assert.True(PdfEmbeddedFontFamily.TrySelectSystemFace(faces, "Numeric Proof", descriptor, "Italic proof", out var selected));
        var options = new PdfOptions { CompressContentStreams = false }.RegisterNamedFontFamily(selected!);
        byte[] pdf = PdfDocument.Create(options).Paragraph(p => p.FontFamily(selected!.FamilyName)
            .Italic(true).Text("Italic proof")).ToBytes();
        string raw = Encoding.ASCII.GetString(pdf);
        Assert.Equal(needsSynthesis, raw.Contains("1 0 0.333 1"));
        Assert.Contains("Italic proof", PdfReadDocument.Open(pdf).ExtractText());
    }
}
