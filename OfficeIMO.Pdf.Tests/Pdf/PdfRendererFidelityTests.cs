using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>Pixel-level regressions for managed-rasterizer conventions that follow the PDF imaging model.</summary>
public sealed partial class PdfRendererFidelityTests {
    [Fact]
    public void PackedTwoBitRgbSamplesExpandToFullRange() {
        // (3,0,1) and (0,2,3) packed as 2-bit components across a byte boundary.
        OfficeRasterImage image = Image(2, 1, "/DeviceRGB", 2, new byte[] { 0xC4, 0xB0 });
        AssertColor(image, 30, 50, 255, 0, 85);
        AssertColor(image, 70, 50, 0, 170, 255);
    }

    [Fact]
    public void SixteenBitRgbSamplesUseTheirHighPrecisionValue() {
        OfficeRasterImage image = Image(1, 1, "/DeviceRGB", 16, new byte[] { 0xFF, 0xFF, 0x80, 0x00, 0x00, 0x00 });
        AssertColor(image, 50, 50, 255, 128, 0);
    }

    [Fact]
    public void FourBitCmykSamplesUseTheSharedCmykModel() {
        OfficeRasterImage image = Image(1, 1, "/DeviceCMYK", 4, new byte[] { 0x0F, 0xF0 });
        OfficeColor expected = OfficeColorSpaceConverter.FromCmyk(0D, 1D, 1D, 0D);
        AssertColor(image, 50, 50, expected.R, expected.G, expected.B);
    }

    [Fact]
    public void MatteSoftMaskColorsAreUnpremultipliedBeforeCompositing() {
        // Red at 50% alpha pre-blended with a black matte is stored as (128,0,0); on white it composites to (255,128,128).
        OfficeRasterImage image = Image(1, 1, "/DeviceRGB", 8, new byte[] { 128, 0, 0 },
            softMask: Mask(1, 1, 8, new byte[] { 128 }, "/Matte [0 0 0]"));
        AssertColor(image, 50, 50, 255, 128, 128, tolerance: 3);
    }

    [Fact]
    public void SoftMasksOfAnotherSizeAndDepthAreResampledToTheImage() {
        // A 4x2 one-bit mask keeps the red left half of a 2x1 image and hides the blue right half.
        OfficeRasterImage image = Image(2, 1, "/DeviceRGB", 8, new byte[] { 255, 0, 0, 0, 0, 255 },
            softMask: Mask(4, 2, 1, new byte[] { 0xC0, 0xC0 }));
        AssertColor(image, 30, 50, 255, 0, 0);
        AssertColor(image, 70, 50, 255, 255, 255);
    }

    [Theory]
    [InlineData("ColorDodge", "0 0 0", "1 1 1", 0)]
    [InlineData("ColorDodge", "0.5 0.5 0.5", "0.5 0.5 0.5", 255)]
    [InlineData("ColorBurn", "1 1 1", "0 0 0", 255)]
    [InlineData("ColorBurn", "0.5 0.5 0.5", "0.5 0.5 0.5", 0)]
    public void SeparableDodgeAndBurnHonourBackdropExtremes(string mode, string backdrop, string source, int expected) {
        string content = backdrop + " rg 10 10 80 80 re f /G1 gs " + source + " rg 10 10 80 80 re f";
        OfficeRasterImage image = RenderPage(content, "/ExtGState << /G1 << /BM /" + mode + " >> >>");
        AssertColor(image, 50, 50, expected, expected, expected);
    }

    [Theory]
    [InlineData("/TR 6 0 R", 255, 64, 64)]
    [InlineData("/TR /Identity", 255, 191, 191)]
    public void SoftMaskTransferFunctionsMapMaskValues(string transfer, int r, int g, int b) {
        // A luminosity mask of 0.25; the transfer function 1 - x turns it into 0.75 coverage of red over white.
        byte[] group = Encoding.ASCII.GetBytes("0.25 g 0 0 100 100 re f");
        var objects = new List<(string, byte[])> {
            ("<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Group << /S /Transparency /CS /DeviceRGB >> /Length " +
             group.Length.ToString(CultureInfo.InvariantCulture) + " >>", group),
            ("<< /FunctionType 2 /Domain [0 1] /C0 [1] /C1 [0] /N 1 >>", Array.Empty<byte>())
        };
        byte[] pdf = BuildPdf("/G1 gs 1 0 0 rg 10 10 80 80 re f",
            "<< /ExtGState << /G1 << /SMask << /Type /Mask /S /Luminosity /G 5 0 R " + transfer + " >> >> >> >>", objects: objects);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf, options: new PdfPageRenderOptions { Dpi = 72D }));
        Assert.DoesNotContain(result.CapabilityDiagnostics, static diagnostic => diagnostic.Code == PdfRenderCapabilities.UnsupportedSoftMaskId);
        AssertColor(PdfRendererQualitySupport.Decode(result.Bytes!), 50, 50, r, g, b, tolerance: 3);
    }

    [Fact]
    public void KnockoutGroupsReportTheirOrdinaryCompositing() {
        byte[] form = Encoding.ASCII.GetBytes("0 0 1 rg 10 10 50 50 re f 1 0 0 rg 30 30 50 50 re f");
        var objects = new List<(string, byte[])> {
            ("<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Group << /S /Transparency /I true /K true >> /Length " +
             form.Length.ToString(CultureInfo.InvariantCulture) + " >>", form)
        };
        byte[] pdf = BuildPdf("/Fm1 Do", "<< /XObject << /Fm1 5 0 R >> >>", objects: objects);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf, options: new PdfPageRenderOptions { Dpi = 72D }));
        Assert.Contains(result.CapabilityDiagnostics, static diagnostic =>
            diagnostic.Code == PdfRenderCapabilities.KnockoutGroupId && diagnostic.Subject == "Fm1");
    }

    [Fact]
    public void StrokesDefaultToButtCapsAndMiterJoins() {
        OfficeRasterImage image = RenderPage("10 w 20 50 m 80 50 l S 10 w 20 20 m 50 20 l 50 5 l S");
        AssertColor(image, 15, 50, 255, 255, 255);
        AssertColor(image, 50, 50, 0, 0, 0);
        // The miter join fills the outer corner square that a round or bevel join would cut.
        AssertColor(image, 54, 76, 0, 0, 0);
    }

    [Theory]
    [InlineData(10, true)]
    [InlineData(1, false)]
    public void MiterLimitSwitchesSharpJoinsToBevels(int miterLimit, bool expectTip) {
        OfficeRasterImage image = RenderPage(miterLimit.ToString(CultureInfo.InvariantCulture) + " M 20 w 10 30 m 50 70 l 90 30 l S");
        double luma = Luma(image.GetPixel(50, 100 - 81));
        if (expectTip) Assert.True(luma < 64D, "Expected the miter tip at (50,81), luma " + luma);
        else Assert.True(luma > 192D, "Expected a bevel without a tip at (50,81), luma " + luma);
    }

    [Theory]
    [InlineData("0 w 10 50.3 m 90 50.3 l S")]
    [InlineData("0.05 w 10 50.3 m 90 50.3 l S")]
    [InlineData("q 0.1 0 0 0.1 0 0 cm 1 w 100 503 m 900 503 l S Q")]
    public void HairlinesPaintOneSolidDevicePixel(string content) {
        OfficeRasterImage image = RenderPage(content);
        int inked = 0;
        double darkest = 255D;
        for (int y = 40; y < 60; y++) {
            double luma = Luma(image.GetPixel(50, y));
            if (luma < 250D) inked++;
            darkest = Math.Min(darkest, luma);
        }
        Assert.Equal(1, inked);
        Assert.True(darkest < 8D, "Hairline luma " + darkest);
    }

    [Fact]
    public void PageSizesThatAreWholePixelsDoNotGainAnExtraRow() {
        // A 792pt page at 150 DPI is 1650.0000000000002 pixels in binary arithmetic, and exactly 1650 rows.
        byte[] pdf = BuildPdf("0 0 1 rg 0 0 612 792 re f", "<< >>", mediaBox: "0 0 612 792");
        OfficeImageExportResult result = PdfReadDocument.Open(pdf).Pages[0]
            .ExportImage(OfficeImageExportFormat.Png, new PdfImageExportOptions { TargetDpi = 150D });
        OfficeRasterImage image = PdfRendererQualitySupport.Decode(result.Bytes!);
        Assert.Equal((1275, 1650), (image.Width, image.Height));
        AssertColor(image, 1274, 1649, 0, 0, 255);
    }

    [Fact]
    public void SymbolFontTextDecodesThroughItsBuiltInEncoding() {
        byte[] pdf = BuildPdf("BT /F1 12 Tf 10 50 Td (abg\\245) Tj ET",
            "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Symbol >> >> >>");
        Assert.Contains("αβγ∞", PdfReadDocument.Open(pdf).Pages[0].ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void WidgetsWithoutAppearancesAreSynthesizedFromTheirFields() {
        PdfPageRenderResult result = RenderWidgets(needAppearances: false,
            "<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /V (Hello) /Rect [10 60 90 80] /P 3 0 R " +
            "/DA (/Helv 12 Tf 0 0 1 rg) /MK << /BG [1 1 0] /BC [1 0 0] >> >>",
            "<< /Type /Annot /Subtype /Widget /FT /Btn /T (check) /V /Yes /AS /Yes /Rect [10 10 40 40] /P 3 0 R /MK << /BC [0 0 0] >> >>");
        OfficeRasterImage image = PdfRendererQualitySupport.Decode(result.Bytes!);

        AssertColor(image, 85, 100 - 63, 255, 255, 0);
        Assert.True(CountPixels(image, 12, 20, 60, 40, color => color.B > 128 && color.R < 128) > 20, "Expected blue field text.");
        Assert.True(CountPixels(image, 14, 62, 36, 86, color => Luma(color) < 96D) > 20, "Expected a check mark.");
        Assert.Contains(result.CapabilityDiagnostics, static diagnostic =>
            diagnostic.Code == PdfRenderCapabilities.SynthesizedAnnotationAppearanceId && diagnostic.Subject == "Widget[0]");
        Assert.DoesNotContain(result.CapabilityDiagnostics, static diagnostic => diagnostic.Code == PdfRenderCapabilities.AnnotationAppearanceId);
    }

    [Fact]
    public void NeedAppearancesRegeneratesStaleTextWidgetAppearances() {
        PdfPageRenderResult result = RenderWidgets(needAppearances: true,
            "<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /V (Hello) /Rect [10 60 90 80] /P 3 0 R " +
            "/DA (/Helv 12 Tf 0 g) /AP << /N 7 0 R >> >>");
        OfficeRasterImage image = PdfRendererQualitySupport.Decode(result.Bytes!);

        AssertColor(image, 85, 100 - 63, 255, 255, 255);
        Assert.True(CountPixels(image, 12, 20, 60, 40, color => Luma(color) < 96D) > 20, "Expected regenerated field text.");
        Assert.Contains(result.CapabilityDiagnostics, static diagnostic => diagnostic.Code == PdfRenderCapabilities.SynthesizedAnnotationAppearanceId);
    }

    private static PdfPageRenderResult RenderWidgets(bool needAppearances, params string[] widgets) {
        var objects = new List<(string, byte[])>();
        var references = new StringBuilder();
        foreach (string widget in widgets) {
            references.Append(' ').Append((5 + objects.Count).ToString(CultureInfo.InvariantCulture)).Append(" 0 R");
            objects.Add((widget, Array.Empty<byte>()));
        }
        // A stale appearance that would paint the whole field red.
        while (objects.Count < 2) objects.Add(("<< >>", Array.Empty<byte>()));
        byte[] stale = Encoding.ASCII.GetBytes("1 0 0 rg 0 0 80 20 re f");
        objects.Add(("<< /Type /XObject /Subtype /Form /BBox [0 0 80 20] /Length " + stale.Length.ToString(CultureInfo.InvariantCulture) + " >>", stale));
        string acroForm = "/AcroForm << /Fields [" + references + " ] /DA (/Helv 0 Tf 0 g) " +
            "/DR << /Font << /Helv << /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >> >> >>" +
            (needAppearances ? " /NeedAppearances true" : string.Empty) + " >>";
        byte[] pdf = BuildPdf(string.Empty, "<< >>", objects: objects, catalogEntries: acroForm, pageEntries: "/Annots [" + references + " ]");
        return Assert.Single(PdfPageImageRenderer.RenderPages(pdf, options: new PdfPageRenderOptions { Format = global::OfficeIMO.Drawing.OfficeImageExportFormat.Png, Dpi = 72D }));
    }

    private static int CountPixels(OfficeRasterImage image, int left, int top, int right, int bottom, Func<OfficeColor, bool> predicate) {
        int count = 0;
        for (int y = top; y < bottom; y++) for (int x = left; x < right; x++) if (predicate(image.GetPixel(x, y))) count++;
        return count;
    }

    private static OfficeRasterImage Image(int width, int height, string colorSpace, int bits, byte[] data,
        (string Dictionary, byte[] Data)? softMask = null) {
        var objects = new List<(string, byte[])> {
            (ImageDictionary(width, height, colorSpace, bits, data.Length, softMask == null ? string.Empty : "/SMask 6 0 R"), data)
        };
        if (softMask != null) objects.Add(softMask.Value);
        return Render(BuildPdf("q 80 0 0 80 10 10 cm /Im1 Do Q", "<< /XObject << /Im1 5 0 R >> >>", objects: objects));
    }

    private static (string Dictionary, byte[] Data) Mask(int width, int height, int bits, byte[] data, string extra = "") =>
        (ImageDictionary(width, height, "/DeviceGray", bits, data.Length, extra), data);

    private static string ImageDictionary(int width, int height, string colorSpace, int bits, int length, string extra) =>
        string.Format(CultureInfo.InvariantCulture,
            "<< /Type /XObject /Subtype /Image /Width {0} /Height {1} /ColorSpace {2} /BitsPerComponent {3} {4} /Length {5} >>",
            width, height, colorSpace, bits, extra, length);

    private static OfficeRasterImage RenderPage(string content, string resources = "") =>
        Render(BuildPdf(content, "<< " + resources + " >>"));

    private static OfficeRasterImage Render(byte[] pdf) {
        OfficeImageExportResult result = PdfReadDocument.Open(pdf).Pages[0]
            .ExportImage(OfficeImageExportFormat.Png, new PdfImageExportOptions { TargetDpi = 72D });
        return PdfRendererQualitySupport.Decode(result.Bytes);
    }

    private static byte[] BuildPdf(string content, string resources, string mediaBox = "0 0 100 100",
        IReadOnlyList<(string Dictionary, byte[] Data)>? objects = null, string catalogEntries = "", string pageEntries = "") {
        using var pdf = new MemoryStream();
        void Ascii(string text) { byte[] bytes = Encoding.ASCII.GetBytes(text); pdf.Write(bytes, 0, bytes.Length); }
        Ascii("%PDF-1.4\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R " + catalogEntries + " >>\nendobj\n");
        Ascii("2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [" + mediaBox + "] >>\nendobj\n");
        Ascii("3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources " + resources + " /Contents 4 0 R " + pageEntries + " >>\nendobj\n");
        Ascii("4 0 obj\n<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + content + "\nendstream\nendobj\n");
        if (objects != null) {
            for (int index = 0; index < objects.Count; index++) {
                Ascii((5 + index).ToString(CultureInfo.InvariantCulture) + " 0 obj\n" + objects[index].Dictionary);
                if (objects[index].Data.Length > 0) {
                    Ascii("\nstream\n");
                    pdf.Write(objects[index].Data, 0, objects[index].Data.Length);
                    Ascii("\nendstream");
                }
                Ascii("\nendobj\n");
            }
        }
        Ascii("trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return pdf.ToArray();
    }

    private static void AssertColor(OfficeRasterImage image, int x, int y, int r, int g, int b, int tolerance = 2) {
        OfficeColor color = image.GetPixel(x, y);
        double alpha = color.A / 255D;
        int actualR = (int)Math.Round(color.R * alpha + 255D * (1D - alpha));
        int actualG = (int)Math.Round(color.G * alpha + 255D * (1D - alpha));
        int actualB = (int)Math.Round(color.B * alpha + 255D * (1D - alpha));
        Assert.True(Math.Abs(actualR - r) <= tolerance && Math.Abs(actualG - g) <= tolerance && Math.Abs(actualB - b) <= tolerance,
            $"Pixel ({x},{y}) is ({actualR},{actualG},{actualB}); expected ({r},{g},{b}).");
    }

    private static double Luma(OfficeColor color) {
        double alpha = color.A / 255D;
        return 0.299D * (color.R * alpha + 255D * (1D - alpha)) + 0.587D * (color.G * alpha + 255D * (1D - alpha)) +
            0.114D * (color.B * alpha + 255D * (1D - alpha));
    }
}
