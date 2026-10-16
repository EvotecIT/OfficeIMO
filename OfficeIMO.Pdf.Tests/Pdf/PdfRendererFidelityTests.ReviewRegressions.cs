using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfRendererFidelityTests {
    [Fact]
    public void SearchFindsVisibleTextInsideAPartiallyClippedRun() {
        byte[] pdf = BuildPdf("q 0 0 40 100 re W n BT /F1 12 Tf 10 50 Td (ABC hidden) Tj ET Q",
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>");
        PdfDocument document = PdfDocument.Load(pdf);
        Assert.Single(document.Text.Find("ABC"));
        Assert.Empty(document.Text.Find("hidden"));
    }

    [Fact]
    public void MatteCorrectionAppliesDecodeBeforeUnpremultiplying() {
        var objects = new List<(string, byte[])> {
            (ImageDictionary(1, 1, "/DeviceRGB", 8, 3, "/SMask 6 0 R /Decode [1 0 1 0 1 0]"), new byte[] { 127, 255, 255 }),
            Mask(1, 1, 8, new byte[] { 128 }, "/Matte [0 0 0]")
        };
        OfficeRasterImage image = Render(BuildPdf("q 80 0 0 80 10 10 cm /Im1 Do Q", "<< /XObject << /Im1 5 0 R >> >>", objects: objects));
        AssertColor(image, 50, 50, 255, 128, 128, tolerance: 3);
    }

    [Theory]
    [InlineData(false, 255)]
    [InlineData(true, 223)]
    public void ResizedSoftMasksHonorTheirInterpolationFlag(bool interpolate, int expectedGreen) {
        OfficeRasterImage image = Image(4, 1, "/DeviceRGB", 8,
            new byte[] { 255, 0, 0, 255, 0, 0, 255, 0, 0, 255, 0, 0 },
            softMask: Mask(2, 1, 8, new byte[] { 0, 128 }, interpolate ? "/Interpolate true" : "/Interpolate false"));
        AssertColor(image, 40, 50, 255, expectedGreen, expectedGreen, tolerance: 3);
    }

    [Fact]
    public void NarrowSixteenBitColorKeysAreDiagnosedInsteadOfChangingMembership() {
        var objects = new List<(string, byte[])> {
            (ImageDictionary(2, 1, "/DeviceRGB", 16, 12, "/Mask [1 1 0 0 0 0]"), new byte[] { 0, 1, 0, 0, 0, 0, 0, 2, 0, 0, 0, 0 })
        };
        byte[] pdf = BuildPdf("q 80 0 0 80 10 10 cm /Im1 Do Q", "<< /XObject << /Im1 5 0 R >> >>", objects: objects);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedImageId);
    }

    [Theory]
    [InlineData("1 J", 14, 44)]
    [InlineData("1 j", 55, 75)]
    [InlineData("2 j", 55, 75)]
    public void StrokeAdjustmentPreservesRoundCapsAndNonMiterJoins(string style, int x, int y) {
        OfficeRasterImage image = RenderPage(style + " 12 w 20 50 m 80 50 l S 20 20 m 50 20 l 50 5 l S");
        AssertColor(image, x, y, 255, 255, 255);
        AssertColor(image, 50, 50, 0, 0, 0);
    }

    [Fact]
    public void ChoiceAppearancesUseInheritedDisplayValues() {
        var objects = new List<(string, byte[])> {
            ("<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [10 60 90 80] /P 3 0 R >>", Array.Empty<byte>()),
            ("<< /FT /Ch /T (country) /V (US) /Opt [[(US) (United States)]] /Kids [5 0 R] >>", Array.Empty<byte>())
        };
        byte[] pdf = BuildPdf(string.Empty, "<< >>", objects: objects,
            catalogEntries: "/AcroForm << /Fields [6 0 R] /NeedAppearances true >>", pageEntries: "/Annots [5 0 R]");
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        PdfDictionary widget = Assert.IsType<PdfDictionary>(document.Objects[5].Value);
        Assert.True(PdfFormFiller.TryCreateTransientWidgetAppearance(document.Objects, widget, null, out PdfStream appearance));
        Assert.Contains("556E6974656420537461746573", Encoding.ASCII.GetString(appearance.Data));
        Assert.DoesNotContain("(US)", Encoding.ASCII.GetString(appearance.Data));
    }

    [Fact]
    public void GenericType3GlyphNamesDoNotBecomeDingbats() {
        byte[] pdf = BuildPdf("BT /F1 12 Tf 10 50 Td (A) Tj ET",
            "<< /Font << /F1 << /Type /Font /Subtype /Type3 /BaseFont /TeX /FontBBox [0 0 1000 1000] " +
            "/FontMatrix [.001 0 0 .001 0 0] /FirstChar 65 /LastChar 65 /Widths [600] " +
            "/Encoding << /BaseEncoding /WinAnsiEncoding /Differences [65 /a65] >> /CharProcs << >> >> >> >>");
        Assert.Equal("A", PdfReadDocument.Open(pdf).ExtractText().Trim());
    }

    [Fact]
    public void IncompleteFreeFormMeshIsOmittedWithADiagnostic() {
        var objects = new List<(string, byte[])> {
            ("<< /ShadingType 4 /ColorSpace /DeviceRGB /BitsPerCoordinate 8 /BitsPerComponent 8 " +
             "/BitsPerFlag 8 /Decode [0 100 0 100 0 1 0 1 0 1] /Length 6 >>", new byte[] { 1, 10, 10, 255, 0, 0 })
        };
        byte[] pdf = BuildPdf("/Sh1 sh", "<< /Shading << /Sh1 5 0 R >> >>", objects: objects);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        Assert.NotEmpty(result.CapabilityDiagnostics);
    }

    [Fact]
    public void MeshSamplingChargesItsCandidateSearch() {
        var x = new List<double> { 0, 100, 0 };
        var y = new List<double> { 0, 0, 100 };
        var triangles = Enumerable.Range(0, 500).SelectMany(_ => new[] { 0, 1, 2 }).ToList();
        var mesh = new PdfMeshShadingSampler(x, y, new List<double> { 0, 0, 0 }, 1, triangles, null);
        Assert.True(mesh.GetEvaluationCost(90, 90) >= 500);
    }
}
