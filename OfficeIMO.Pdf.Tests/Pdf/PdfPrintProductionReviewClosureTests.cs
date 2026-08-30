using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPrintProductionReviewClosureTests {
    [Theory]
    [InlineData("/ColorSpace /DeviceRGB", "/Matte [0 0 0]")]
    [InlineData("/ColorSpace /DeviceGray", "/Mask [0 1]")]
    public void ImageSoftMasksRequireGrayscaleAndCannotCarryNestedMasks(
        string softMaskColorSpace,
        string softMaskEntries) {
        byte[] pdf = BuildInspectionPdf(
            "/Im1 Do",
            resources: "/XObject << /Im1 5 0 R >>",
            extraObjects:
                ImageStream(5, "/ColorSpace /DeviceRGB /BitsPerComponent 8 /SMask 6 0 R", new byte[] { 1, 2, 3 }) +
                ImageStream(6, softMaskColorSpace + " /BitsPerComponent 8 " + softMaskEntries, new byte[] { 128 }));

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.True(evidence.UninspectableContentStreamCount > 0);
    }

    [Theory]
    [InlineData("/Mask [0 255]")]
    [InlineData("/Mask [0 255 0 255 0 255] /SMask 6 0 R")]
    [InlineData("/Mask 6 0 R")]
    public void ImageColorKeyMasksMatchTheImageComponentsAndAreExclusiveWithSoftMasks(string maskEntries) {
        byte[] pdf = BuildInspectionPdf(
            "/Im1 Do",
            resources: "/XObject << /Im1 5 0 R >>",
            extraObjects:
                ImageStream(5, "/ColorSpace /DeviceRGB /BitsPerComponent 8 " + maskEntries, new byte[] { 1, 2, 3 }) +
                ImageStream(6, "/ColorSpace /DeviceGray /BitsPerComponent 8", new byte[] { 128 }));

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
    }

    [Theory]
    [InlineData("/Alternate /DeviceRGB")]
    [InlineData("/Range [0 1]")]
    [InlineData("/Range [1 0 0 1 0 1 0 1]")]
    public void IccBasedColorSpacesRejectIncompatibleAlternateAndRangeEntries(string profileEntries) {
        byte[] profile = IccMabTestProfiles.CreateCmykLab8Bidirectional();
        byte[] pdf = BuildInspectionPdf(
            "/Im1 Do",
            resources: "/XObject << /Im1 5 0 R >>",
            extraObjects:
                ImageStream(5, "/ColorSpace [/ICCBased 6 0 R] /BitsPerComponent 8", new byte[] { 0, 0, 0, 0 }) +
                StreamObject(6, "/N 4 " + profileEntries, profile));

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
    }

    [Fact]
    public void ShadingPatternsInspectTheirGraphicsStateOpacity() {
        byte[] pdf = BuildInspectionPdf(
            "/Pattern cs /P1 scn 0 0 10 10 re f",
            resources: "/Pattern << /P1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /Pattern /PatternType 2 /ExtGState << /ca 0.5 >> " +
                "/Shading << /ShadingType 2 /ColorSpace /DeviceCMYK /Coords [0 0 100 0] " +
                "/Function << /FunctionType 2 /Domain [0 1] /C0 [0 0 0 0] /C1 [1 1 1 1] /N 1 >> >> >>\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.True(evidence.HasTransparency);
        Assert.Equal(1, evidence.NonOpaqueGraphicsStateCount);
    }

    [Fact]
    public void ShadingPatternsRejectMalformedMatrices() {
        byte[] pdf = BuildInspectionPdf(
            "/Pattern cs /P1 scn 0 0 10 10 re f",
            resources: "/Pattern << /P1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /Pattern /PatternType 2 /Matrix [1 0 0 1 0] " +
                "/Shading << /ShadingType 2 /ColorSpace /DeviceCMYK /Coords [0 0 100 0] " +
                "/Function << /FunctionType 2 /Domain [0 1] /C0 [0 0 0 0] /C1 [1 1 1 1] /N 1 >> >> >>\nendobj\n");

        Assert.False(PdfReadDocument.Open(pdf).InspectPrintProductionColors().IsComplete);
    }

    [Theory]
    [InlineData("/Background [0 0 0]")]
    [InlineData("/Background [0 0 0 1e309]")]
    public void ShadingsRequireOneFiniteBackgroundValuePerComponent(string background) {
        byte[] pdf = BuildInspectionPdf(
            "/Sh1 sh",
            resources: "/Shading << /Sh1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceCMYK " + background +
                " /Coords [0 0 100 0] /Function << /FunctionType 2 /Domain [0 1] " +
                "/C0 [0 0 0 0] /C1 [1 1 1 1] /N 1 >> >>\nendobj\n");

        Assert.False(PdfReadDocument.Open(pdf).InspectPrintProductionColors().IsComplete);
    }

    [Fact]
    public void SelectedType0FontsRequireTheirParentFontDictionaryType() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) return;
        var options = new PdfOptions()
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath), "Type0 validation font");
        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Selected Type0 font."))
            .ToBytes();
        ReplaceAsciiOnce(pdf, "/Type /Font /Subtype /Type0", "/Type /Null /Subtype /Type0");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.False(evidence.IsComplete);
    }

    private static byte[] BuildInspectionPdf(string content, string resources, string extraObjects) {
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << " + resources + " >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n" + extraObjects + "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static string ImageStream(int objectNumber, string entries, byte[] data) =>
        StreamObject(objectNumber, "/Type /XObject /Subtype /Image /Width 1 /Height 1 " + entries, data);

    private static string StreamObject(int objectNumber, string entries, byte[] data) {
        string payload = new(Array.ConvertAll(data, value => (char)value));
        return objectNumber + " 0 obj\n<< " + entries + " /Length " + data.Length + " >>\nstream\n" +
            payload + "\nendstream\nendobj\n";
    }

    private static void ReplaceAsciiOnce(byte[] bytes, string oldValue, string newValue) {
        Assert.Equal(oldValue.Length, newValue.Length);
        byte[] oldBytes = Encoding.ASCII.GetBytes(oldValue);
        int offset = bytes.AsSpan().IndexOf(oldBytes);
        Assert.True(offset >= 0, "Expected PDF token was not found.");
        Encoding.ASCII.GetBytes(newValue, 0, newValue.Length, bytes, offset);
    }

    private static void WriteAscii(Stream output, string text) {
        byte[] bytes = Encoding.ASCII.GetBytes(text);
        output.Write(bytes, 0, bytes.Length);
    }
}
