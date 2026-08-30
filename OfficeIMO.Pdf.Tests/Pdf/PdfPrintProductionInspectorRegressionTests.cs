using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfPrintProductionInspectorRegressionTests {
    [Fact]
    public void ColorInspectorFindsDirectResourceAndPatternShadingDictionaries() {
        byte[] pdf = BuildInspectionPdf(
            "/S1 sh /Pattern cs /P1 scn",
            resources:
                "/Shading << /S1 << /ShadingType 2 /ColorSpace /DeviceRGB >> >> " +
                "/Pattern << /P1 << /Type /Pattern /PatternType 2 /Shading << /ShadingType 3 /ColorSpace /DeviceCMYK >> >> >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Equal(1, evidence.DeviceRgbShadingCount);
        Assert.Equal(1, evidence.DeviceCmykShadingCount);
    }

    [Fact]
    public void ColorInspectorResolvesShadingColorSpaceFromInvokingResources() {
        byte[] pdf = BuildInspectionPdf(
            "/S1 sh",
            resources:
                "/ColorSpace << /CS1 /DeviceRGB >> " +
                "/Shading << /S1 << /ShadingType 2 /ColorSpace /CS1 /Coords [0 0 100 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [0 0 0] /C1 [1 0 0] /N 1 >> >> >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Equal(1, evidence.DeviceRgbShadingCount);
        Assert.True(evidence.HasDeviceRgbUsage);
    }

    [Fact]
    public void ColorInspectorTracksDeviceCmykSelectionsAndResourceAliases() {
        byte[] pdf = BuildInspectionPdf(
            "/DeviceCMYK cs 0 0 0 1 sc /PrintCmyk CS 0 0 0 1 SCN",
            resources: "/ColorSpace << /PrintCmyk /DeviceCMYK >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Equal(4, evidence.DeviceCmykOperatorCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorFramesRawInlineImagesUsingResourceAliasComponents() {
        byte[] pdf = BuildInspectionPdf(
            "BI /W 1 /H 1 /BPC 8 /CS /PrintCmyk ID abcd EI 0 0 0 1 k",
            resources: "/ColorSpace << /PrintCmyk /DeviceCMYK >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(1, evidence.DeviceCmykImageCount);
        Assert.Equal(1, evidence.DeviceCmykOperatorCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("[99 0 R /Normal]")]
    [InlineData("[[/Normal] /Multiply]")]
    [InlineData("[1 /Normal]")]
    public void ColorInspectorRejectsMalformedBlendModeArrayEntries(string blendMode) {
        byte[] pdf = BuildInspectionPdf(
            "/Blend gs",
            resources: "/ExtGState << /Blend << /BM " + blendMode + " >> >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("/F1 12 Tf (A) Tj")]
    [InlineData("BT /F1 12 Tf ET (A) Tj")]
    [InlineData("BT BT /F1 12 Tf (A) Tj ET")]
    [InlineData("BT /F1 12 Tf (A) Tj")]
    public void StructureInspectorRejectsTextOperatorsOutsideBalancedTextObjects(string content) {
        byte[] pdf = BuildInspectionPdf(
            content,
            resources: "/Font << /F1 5 0 R >>",
            pageEntries: "/TrimBox [10 10 90 90]",
            extraObjects: "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.False(evidence.IsComplete);
        Assert.True(evidence.UninspectableFontResourceCount > 0);
    }

    [Fact]
    public void StructureInspectorFramesRawInlineImagesUsingResourceAliasComponents() {
        byte[] pdf = BuildInspectionPdf(
            "BI /W 1 /H 1 /BPC 8 /CS /PrintCmyk ID abcd EI BT /F1 12 Tf (A) Tj ET",
            resources: "/ColorSpace << /PrintCmyk /DeviceCMYK >> /Font << /F1 5 0 R >>",
            pageEntries: "/TrimBox [10 10 90 90]",
            extraObjects: "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Fact]
    public void ColorInspectorPreservesSelectedColorSpaceAcrossOrderedPageContentStreams() {
        byte[] pdf = BuildInspectionPdf(
            "/DeviceRGB cs ",
            contents: "[4 0 R 5 0 R]",
            extraObjects:
                "5 0 obj\n<< /Length 8 >>\nstream\n1 0 0 sc\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(2, evidence.DeviceRgbOperatorCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorInterpretsOrderedPageContentStreamsAsOneTokenStream() {
        byte[] pdf = BuildInspectionPdf(
            "/DeviceRGB",
            contents: "[4 0 R 5 0 R]",
            extraObjects:
                "5 0 obj\n<< /Length 12 >>\nstream\n cs 1 0 0 sc\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(2, evidence.DeviceRgbOperatorCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorPreservesRepeatedPageContentStreamReferences() {
        byte[] pdf = BuildInspectionPdf("1 0 0 rg ", contents: "[4 0 R 4 0 R]");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(2, evidence.DeviceRgbOperatorCount);
    }

    [Fact]
    public void StructureInspectorInterpretsOrderedPageContentStreamsAsOneTokenStream() {
        byte[] pdf = BuildInspectionPdf(
            "BT /F1",
            resources: "/Font << /F1 6 0 R >>",
            pageEntries: "/TrimBox [10 10 90 90]",
            contents: "[4 0 R 5 0 R]",
            extraObjects:
                "5 0 obj\n<< /Length 9 >>\nstream\n 12 Tf ET\nendstream\nendobj\n" +
                "6 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Theory]
    [InlineData("k")]
    [InlineData("/Bogus k")]
    [InlineData("1 0 rg")]
    [InlineData("1 0 /Bogus RG")]
    [InlineData("g")]
    [InlineData("/DeviceRGB cs 1 0 sc")]
    [InlineData("/DeviceCMYK CS 0 0 0 /Bogus SCN")]
    [InlineData("/Pattern cs /P1 sc")]
    public void ColorInspectorRejectsMalformedColorSettingOperands(string content) {
        byte[] pdf = BuildInspectionPdf(content);

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("q 99 Q 0.5 g")]
    [InlineData("99 q 0.5 g Q")]
    public void ColorInspectorRejectsGraphicsStateOperatorsWithOperands(string content) {
        byte[] pdf = BuildInspectionPdf(content);

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorRejectsMalformedArrayColorSpaceAliasWithoutClassifyingNestedNames() {
        byte[] pdf = BuildInspectionPdf(
            "/CS1 cs 0 0 0 1 sc",
            resources: "/ColorSpace << /CS1 [/Bogus /DeviceCMYK] >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(0, evidence.DeviceCmykOperatorCount);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorRejectsPatternAsTheBaseOfAnUncoloredPatternColorSpace() {
        byte[] pdf = BuildInspectionPdf(
            "/CS1 cs /P1 scn",
            resources: "/ColorSpace << /CS1 [/Pattern /Pattern] >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorRejectsUndersizedIndexedLookupPayload() {
        byte[] pdf = BuildInspectionPdf(
            "/CS1 cs 0 sc",
            resources: "/ColorSpace << /CS1 [/Indexed /DeviceCMYK 255 <00>] >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(0, evidence.DeviceCmykOperatorCount);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorDecodesCompleteIndexedLookupStreams() {
        const string lookup = "00000000FFFFFFFF>";
        byte[] pdf = BuildInspectionPdf(
            "/CS1 cs 0 sc",
            resources: "/ColorSpace << /CS1 [/Indexed /DeviceCMYK 1 5 0 R] >>",
            extraObjects:
                "5 0 obj\n<< /Filter /ASCIIHexDecode /Length " + lookup.Length + " >>\nstream\n" +
                lookup + "\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(2, evidence.DeviceCmykOperatorCount);
    }

    [Theory]
    [InlineData("[/Separation /Spot /DeviceCMYK <<>>]", "0.5 sc")]
    [InlineData("[/DeviceN [/Cyan /Spot] /DeviceCMYK <<>>]", "0.5 0.5 sc")]
    public void ColorInspectorRejectsMalformedSpotTintTransforms(string colorSpace, string paint) {
        byte[] pdf = BuildInspectionPdf(
            "/CS1 cs " + paint,
            resources: "/ColorSpace << /CS1 " + colorSpace + " >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(0, evidence.DeviceCmykOperatorCount);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("<< /Subtype 1 >>")]
    [InlineData("<< /Subtype /Bogus >>")]
    [InlineData("<< /Colorants 1 >>")]
    [InlineData("<< /Colorants << /Cyan [/Separation /Magenta /DeviceCMYK 5 0 R] >> >>")]
    [InlineData("<< /Process << /ColorSpace /DeviceCMYK >> >>")]
    [InlineData("<< /Process << /ColorSpace /DeviceCMYK /Components [/Cyan] >> >>")]
    [InlineData("<< /MixingHints << /Solidities << /Cyan 0.8 >> >> >>")]
    [InlineData("<< /MixingHints << /Solidities << /Cyan 1.1 >> /PrintingOrder [/Cyan] >> >>")]
    [InlineData("<< /MixingHints << /PrintingOrder [/Spot] >> >>")]
    [InlineData("<< /Unvalidated true >>")]
    public void ColorInspectorRejectsMalformedDeviceNAttributes(string attributes) {
        byte[] pdf = BuildInspectionPdf(
            "/CS1 cs 0.5 sc",
            resources: "/ColorSpace << /CS1 [/DeviceN [/Cyan] /DeviceCMYK 5 0 R " + attributes + "] >>",
            extraObjects: BuildDeviceNTintFunctionObject());

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(0, evidence.DeviceCmykOperatorCount);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("<< /Subtype /DeviceN >>")]
    [InlineData("<< /Subtype /NChannel /Colorants << /Cyan [/Separation /Cyan /DeviceCMYK 5 0 R] >> >>")]
    [InlineData("<< /Subtype /NChannel /Process << /ColorSpace /DeviceCMYK /Components [/Cyan /Magenta /Yellow /Black] >> >>")]
    [InlineData("<< /MixingHints << /Solidities << /Cyan 0.8 /Default 0 >> /PrintingOrder [/Cyan] >> >>")]
    public void ColorInspectorAcceptsValidDeviceNAttributes(string attributes) {
        byte[] pdf = BuildInspectionPdf(
            "/CS1 cs 0.5 sc",
            resources: "/ColorSpace << /CS1 [/DeviceN [/Cyan] /DeviceCMYK 5 0 R " + attributes + "] >>",
            extraObjects: BuildDeviceNTintFunctionObject());

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(2, evidence.DeviceCmykOperatorCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorAppliesValidDefaultRgbIccProfile() {
        byte[] pdf = BuildInspectionPdf(
            "1 0 0 rg /DeviceRGB cs 1 0 0 sc",
            resources: "/ColorSpace << /DefaultRGB [/ICCBased 6 0 R] >>",
            extraObjects: BuildIccProfileObject(6, 3, PdfIccProfiles.SrgbIec6196621));

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(0, evidence.DeviceRgbOperatorCount);
        Assert.Equal(3, evidence.DeviceIndependentColorUsageCount);
    }

    [Fact]
    public void ColorInspectorFailsClosedOnInvalidDefaultDeviceColorSpace() {
        byte[] pdf = BuildInspectionPdf(
            "0 0 0 1 k",
            resources: "/ColorSpace << /DefaultCMYK /Bogus >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("/DefaultRGB [/ICCBased 99 0 R]", "")]
    [InlineData("/DefaultRGB [/ICCBased 6 0 R]", "6 0 obj\n<< /N 4 /Length 0 >>\nstream\n\nendstream\nendobj\n")]
    [InlineData("/DefaultRGB [/ICCBased 6 0 R]", "6 0 obj\n<< /N 3 /Length 0 >>\nstream\n\nendstream\nendobj\n")]
    public void ColorInspectorRejectsUnresolvedOrWrongComponentDefaultRgbProfiles(
        string defaultColorSpace,
        string extraObjects) {
        byte[] pdf = BuildInspectionPdf(
            "1 0 0 rg",
            resources: "/ColorSpace << " + defaultColorSpace + " >>",
            extraObjects: extraObjects);

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("1 0 0 rg", "/DefaultRGB /DeviceCMYK")]
    [InlineData("0 0 0 1 k", "/DefaultCMYK /DeviceGray")]
    [InlineData("0.5 g", "/DefaultGray /DeviceRGB")]
    public void ColorInspectorRejectsDeviceColorSpacesAsDefaultReplacements(
        string content,
        string defaultColorSpace) {
        byte[] pdf = BuildInspectionPdf(
            content,
            resources: "/ColorSpace << " + defaultColorSpace + " >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("ICCBased")]
    [InlineData("CalRGB")]
    [InlineData("CalGray")]
    [InlineData("Lab")]
    public void ColorInspectorRejectsArrayColorSpaceFamiliesUsedAsDirectNames(string family) {
        byte[] pdf = BuildInspectionPdf("/" + family + " cs 0 sc");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
        Assert.Equal(0, evidence.DeviceIndependentColorUsageCount);
    }

    [Theory]
    [InlineData("[/CalGray <<>>]")]
    [InlineData("[/CalGray << /WhitePoint [0.95 1 1.09] /Gamma 0 >>]")]
    [InlineData("[/CalRGB << /WhitePoint [0.95 1 1.09] /Gamma [1 1] >>]")]
    [InlineData("[/CalRGB << /WhitePoint [0.95 1 1.09] /Matrix [1 0] >>]")]
    [InlineData("[/Lab << /WhitePoint [0.95 1 1.09] /Range [100 -100 -100 100] >>]")]
    [InlineData("[/Lab << /WhitePoint [0.95 1 1.09] /BlackPoint [0 -1 0] >>]")]
    public void ColorInspectorFailsClosedOnMalformedCalibratedColorDictionaries(string colorSpace) {
        byte[] pdf = BuildInspectionPdf(
            "/CS1 cs 0 0 0 sc",
            resources: "/ColorSpace << /CS1 " + colorSpace + " >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
        Assert.Equal(0, evidence.DeviceIndependentColorUsageCount);
    }

    [Fact]
    public void ColorInspectorAppliesDefaultRgbToReachableImagesAndShadings() {
        byte[] pdf = BuildInspectionPdf(
            "/Im1 Do /S1 sh",
            resources:
                "/ColorSpace << /DefaultRGB [/ICCBased 6 0 R] >> " +
                "/XObject << /Im1 5 0 R >> " +
                "/Shading << /S1 << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 10 0] " +
                "/Function << /FunctionType 2 /Domain [0 1] /C0 [0 0 0] /C1 [1 1 1] /N 1 >> >> >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB " +
                "/BitsPerComponent 8 /Length 3 >>\nstream\nrgb\nendstream\nendobj\n" +
                BuildIccProfileObject(6, 3, PdfIccProfiles.SrgbIec6196621));

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(0, evidence.DeviceRgbImageCount);
        Assert.Equal(0, evidence.DeviceRgbShadingCount);
        Assert.Equal(2, evidence.DeviceIndependentColorUsageCount);
    }

    [Fact]
    public void ColorInspectorAppliesInvokingResourcesToFormsWithoutOwnResources() {
        const string formContent = "/PrintRgb cs 1 0 0 sc";
        byte[] pdf = BuildInspectionPdf(
            "/Fm Do",
            resources: "/ColorSpace << /PrintRgb /DeviceRGB >> /XObject << /Fm 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length " +
                formContent.Length +
                " >>\nstream\n" + formContent + "\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Equal(2, evidence.DeviceRgbOperatorCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("")]
    [InlineData("/BBox [10 0 0 10]")]
    [InlineData("/BBox [0 0 10]")]
    public void ColorInspectorRejectsMalformedReachableFormGeometry(string formGeometry) {
        byte[] pdf = BuildInspectionPdf(
            "/Fm Do",
            resources: "/XObject << /Fm 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form " + formGeometry +
                " /Length 0 >>\nstream\n\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();
        PdfPrintProductionStructureEvidence structure = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
        Assert.False(structure.IsComplete);
        Assert.True(structure.UninspectableFontResourceCount > 0);
    }

    [Fact]
    public void ColorInspectorAppliesInvokingColorStateToForms() {
        const string formContent = "1 0 0 sc";
        byte[] pdf = BuildInspectionPdf(
            "/DeviceRGB cs /Fm Do",
            resources: "/XObject << /Fm 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length " +
                formContent.Length +
                " >>\nstream\n" + formContent + "\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(2, evidence.DeviceRgbOperatorCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorInspectsRepeatedFormInvocationsUnderEachColorState() {
        const string formContent = "1 0 0 sc";
        byte[] pdf = BuildInspectionPdf(
            "/DeviceRGB cs /Fm Do /DeviceCMYK cs /Fm Do",
            resources: "/XObject << /Fm 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length " +
                formContent.Length +
                " >>\nstream\n" + formContent + "\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorAcceptsReachableXObjectsWithoutTheOptionalType() {
        byte[] pdf = BuildInspectionPdf(
            "/Im1 Do",
            resources: "/XObject << /Im1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Subtype /Image /Width 1 /Height 1 " +
                "/ColorSpace /DeviceCMYK /BitsPerComponent 8 /Length 4 >>\nstream\ncmyk\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(1, evidence.DeviceCmykImageCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorRejectsReachableXObjectsWithTheWrongType() {
        byte[] pdf = BuildInspectionPdf(
            "/Im1 Do",
            resources: "/XObject << /Im1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /Pattern /Subtype /Image /Width 1 /Height 1 " +
                "/ColorSpace /DeviceCMYK /BitsPerComponent 8 /Length 4 >>\nstream\ncmyk\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorAppliesPatternResourcesToNestedForms() {
        const string patternContent = "/Fm Do";
        const string formContent = "/PrintRgb cs 1 0 0 sc";
        byte[] pdf = BuildInspectionPdf(
            "/Pattern cs /P1 scn 0 0 10 10 re f",
            resources: "/Pattern << /P1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 " +
                "/Resources << /ColorSpace << /PrintRgb /DeviceRGB >> /XObject << /Fm 6 0 R >> >> /Length " +
                patternContent.Length + " >>\nstream\n" + patternContent + "\nendstream\nendobj\n" +
                "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length " +
                formContent.Length + " >>\nstream\n" + formContent + "\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Equal(2, evidence.DeviceRgbOperatorCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorAppliesGraphicsStateResourcesToSoftMaskForms() {
        const string formContent = "/PrintRgb cs 1 0 0 sc";
        byte[] pdf = BuildInspectionPdf(
            "/GS1 gs",
            resources:
                "/ColorSpace << /PrintRgb /DeviceRGB >> " +
                "/ExtGState << /GS1 << /SMask << /S /Luminosity /G 5 0 R >> >> >>",
            extraObjects:
                "5 0 obj\n<< /Subtype /Form /BBox [0 0 10 10] /Group << /S /Transparency /CS /DeviceRGB >> /Length " +
                formContent.Length + " >>\nstream\n" + formContent + "\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Equal(2, evidence.DeviceRgbOperatorCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorRejectsSoftMaskFormsWithTheWrongXObjectType() {
        byte[] pdf = BuildInspectionPdf(
            "/GS1 gs",
            resources: "/ExtGState << /GS1 << /SMask << /S /Alpha /G 5 0 R >> >> >>",
            extraObjects:
                "5 0 obj\n<< /Type /Pattern /Subtype /Form /BBox [0 0 10 10] " +
                "/Group << /S /Transparency >> /Length 0 >>\nstream\n\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("/G 5 0 R", "/Group << /S /Transparency >>")]
    [InlineData("/S /Bogus /G 5 0 R", "/Group << /S /Transparency >>")]
    [InlineData("/S /Alpha /G 5 0 R", "")]
    [InlineData("/S /Luminosity /G 5 0 R", "/Group << /S /Bogus >>")]
    [InlineData("/S /Luminosity /G 5 0 R", "/Group << /S /Transparency >>")]
    [InlineData("/S /Alpha /G 5 0 R /TR <<>>", "/Group << /S /Transparency /CS /DeviceGray >>")]
    [InlineData("/S /Luminosity /G 5 0 R /BC [0 0]", "/Group << /S /Transparency /CS /DeviceRGB >>")]
    public void ColorInspectorFailsClosedOnMalformedReachableSoftMask(
        string softMaskEntries,
        string formGroup) {
        byte[] pdf = BuildInspectionPdf(
            "/GS1 gs",
            resources: "/ExtGState << /GS1 << /SMask << " + softMaskEntries + " >> >> >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] " + formGroup +
                " /Length 0 >>\nstream\n\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorAcceptsValidSoftMaskTransferAndBackdropEntries() {
        byte[] pdf = BuildInspectionPdf(
            "/GS1 gs",
            resources: "/ExtGState << /GS1 << /SMask << /S /Luminosity /G 5 0 R /TR /Identity /BC [0 0 0] >> >> >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] " +
                "/Group << /S /Transparency /CS /DeviceRGB >> /Length 0 >>\nstream\n\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorResolvesLuminosityGroupColorSpaceFromGroupResources() {
        byte[] pdf = BuildInspectionPdf(
            "/GS1 gs",
            resources: "/ColorSpace << /BlendRgb /DeviceGray >> " +
                       "/ExtGState << /GS1 << /SMask << /S /Luminosity /G 5 0 R /BC [0 0 0] >> >> >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] " +
                "/Resources << /ColorSpace << /BlendRgb /DeviceRGB >> >> " +
                "/Group << /S /Transparency /CS /BlendRgb >> /Length 0 >>\nstream\n\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorIgnoresBackdropForAlphaSoftMasks() {
        byte[] pdf = BuildInspectionPdf(
            "/GS1 gs",
            resources: "/ExtGState << /GS1 << /SMask << /S /Alpha /G 5 0 R /BC /Ignored >> >> >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] " +
                "/Group << /S /Transparency >> /Length 0 >>\nstream\n\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorRecognizesInlineImageColorSpaceAbbreviations() {
        byte[] pdf = BuildInspectionPdf(
            "BI /W 1 /H 1 /BPC 8 /CS /RGB ID abc EI\n" +
            "BI /W 1 /H 1 /BPC 8 /CS /CMYK ID abcd EI\n" +
            "BI /W 1 /H 1 /BPC 8 /CS /G ID a EI");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Equal(1, evidence.DeviceRgbImageCount);
        Assert.Equal(1, evidence.DeviceCmykImageCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorTreatsImageSoftMaskNoneAsOpaque() {
        byte[] pdf = BuildInspectionPdf(
            string.Empty,
            extraObjects: "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 /SMask /None /Length 1 >>\nstream\na\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Equal(0, evidence.TransparentImageCount);
    }

    [Fact]
    public void ColorInspectorAcceptsReachableGrayscaleImageSoftMask() {
        byte[] pdf = BuildInspectionPdf(
            "/Im1 Do",
            resources: "/XObject << /Im1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceCMYK /BitsPerComponent 8 /SMask 6 0 R /Length 4 >>\nstream\ncmyk\nendstream\nendobj\n" +
                "6 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 /Length 1 >>\nstream\ng\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(1, evidence.TransparentImageCount);
        Assert.Equal(1, evidence.DeviceCmykImageCount);
        Assert.Equal(0, evidence.DeviceRgbImageCount);
    }

    [Fact]
    public void ColorInspectorFailsClosedOnUninspectableImageSoftMask() {
        byte[] pdf = BuildInspectionPdf(
            "/Im1 Do",
            resources: "/XObject << /Im1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceCMYK /BitsPerComponent 8 /SMask 99 0 R /Length 4 >>\nstream\ncmyk\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.TransparentImageCount);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorResolvesImageColorSpaceFromInvokingPageResources() {
        byte[] pdf = BuildInspectionPdf(
            "/Im1 Do",
            resources: "/ColorSpace << /CS1 /DeviceRGB >> /XObject << /Im1 5 0 R >>",
            extraObjects: "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /CS1 /BitsPerComponent 8 /Length 3 >>\nstream\nrgb\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Equal(1, evidence.DeviceRgbImageCount);
        Assert.True(evidence.HasDeviceRgbUsage);
    }

    [Fact]
    public void ColorInspectorResolvesImageColorSpaceFromInvokingFormResources() {
        const string formContent = "/Im1 Do";
        byte[] pdf = BuildInspectionPdf(
            "/Fm Do",
            resources: "/XObject << /Fm 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] " +
                "/Resources << /ColorSpace << /CS1 /DeviceRGB >> /XObject << /Im1 6 0 R >> >> /Length " +
                formContent.Length + " >>\nstream\n" + formContent + "\nendstream\nendobj\n" +
                "6 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /CS1 /BitsPerComponent 8 /Length 3 >>\nstream\nrgb\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Equal(1, evidence.DeviceRgbImageCount);
        Assert.True(evidence.HasDeviceRgbUsage);
    }

    [Theory]
    [InlineData("/Bogus")]
    [InlineData("/ICCBased")]
    [InlineData("/CalRGB")]
    [InlineData("/CalGray")]
    [InlineData("/Lab")]
    [InlineData("[/Indexed /Bogus 1 <00>]")]
    public void ColorInspectorFailsClosedOnUnclassifiedReachableImageColorSpace(string colorSpace) {
        byte[] pdf = BuildInspectionPdf(
            "/Im1 Do",
            resources: "/XObject << /Im1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace " + colorSpace +
                " /BitsPerComponent 8 /Length 1 >>\nstream\na\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorFailsClosedOnUnclassifiedInlineImageColorSpace() {
        byte[] pdf = BuildInspectionPdf("BI /W 1 /H 1 /BPC 8 /CS /Bogus ID a EI");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorIgnoresUnusedImageAndFormXObjects() {
        const string unusedFormContent = "/DeviceRGB cs 1 0 0 sc";
        const string unusedPatternContent = "/DeviceRGB cs 1 0 0 sc";
        const string unusedSoftMaskContent = "/DeviceRGB cs 1 0 0 sc";
        const string unusedCharProcContent = "/DeviceRGB cs 1 0 0 sc";
        byte[] pdf = BuildInspectionPdf(
            string.Empty,
            resources:
                "/XObject << /UnusedImage 5 0 R /UnusedForm 6 0 R >> " +
                "/Shading << /UnusedShading << /ShadingType 2 /ColorSpace /DeviceRGB >> >> " +
                "/Pattern << /UnusedPattern 7 0 R >> " +
                "/ExtGState << /UnusedState << /ca 0.5 /SMask << /S /Luminosity /G 8 0 R >> >> >> " +
                "/Font << /UnusedFont 9 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\nrgb\nendstream\nendobj\n" +
                "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length " +
                unusedFormContent.Length + " >>\nstream\n" + unusedFormContent + "\nendstream\nendobj\n" +
                "7 0 obj\n<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Length " +
                unusedPatternContent.Length + " >>\nstream\n" + unusedPatternContent + "\nendstream\nendobj\n" +
                "8 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Group << /S /Transparency >> /Length " +
                unusedSoftMaskContent.Length + " >>\nstream\n" + unusedSoftMaskContent + "\nendstream\nendobj\n" +
                "9 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] " +
                "/CharProcs << /A 10 0 R >> /Encoding << /Type /Encoding /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] >>\nendobj\n" +
                "10 0 obj\n<< /Length " + unusedCharProcContent.Length + " >>\nstream\n" + unusedCharProcContent + "\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Equal(0, evidence.DeviceRgbImageCount);
        Assert.Equal(0, evidence.DeviceRgbOperatorCount);
        Assert.Equal(0, evidence.DeviceRgbShadingCount);
        Assert.Equal(0, evidence.NonOpaqueGraphicsStateCount);
        Assert.Equal(0, evidence.TransparencyGroupCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void StructureInspectorAcceptsArtBoxWithoutExplicitBleedBox() {
        byte[] pdf = BuildInspectionPdf(
            string.Empty,
            pageEntries: "/ArtBox [10 10 90 90]");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.ValidProductionPageBoxCount);
        Assert.Equal(0, evidence.InvalidProductionPageBoxCount);
    }

    [Fact]
    public void StructureInspectorRejectsProductionBoxWithExtraCoordinates() {
        byte[] pdf = BuildInspectionPdf(
            string.Empty,
            pageEntries: "/TrimBox [10 10 90 90 999]");
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        PdfPrintProductionStructureEvidence evidence = document.InspectPrintProductionStructure();

        Assert.Null(document.Pages[0].GetGeometry().TrimBox);
        Assert.Equal(0, evidence.ValidProductionPageBoxCount);
        Assert.Equal(1, evidence.InvalidProductionPageBoxCount);
    }

    [Fact]
    public void StructureInspectorRejectsProductionBoxesInheritedFromPageTree() {
        byte[] pdf = BuildInspectionPdf(
            string.Empty,
            pageTreeEntries: "/TrimBox [10 10 90 90]");
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        PdfPrintProductionStructureEvidence evidence = document.InspectPrintProductionStructure();

        Assert.Null(document.Pages[0].GetGeometry().TrimBox);
        Assert.Equal(0, evidence.ValidProductionPageBoxCount);
        Assert.Equal(1, evidence.InvalidProductionPageBoxCount);
    }

    [Fact]
    public void StructureInspectorRejectsNonemptyInvalidEmbeddedFontStream() {
        byte[] pdf = BuildInspectionPdf(
            "BT /F1 12 Tf (A) Tj ET",
            resources: "/Font << /F1 5 0 R >>",
            pageEntries: "/TrimBox [10 10 90 90]",
            extraObjects:
                "5 0 obj\n<< /Type /Font /Subtype /TrueType /BaseFont /Fixture /FontDescriptor 6 0 R >>\nendobj\n" +
                "6 0 obj\n<< /Type /FontDescriptor /FontName /Fixture /Flags 32 /FontBBox [0 0 500 700] " +
                "/ItalicAngle 0 /Ascent 700 /Descent -200 /CapHeight 700 /StemV 80 /FontFile2 7 0 R >>\nendobj\n" +
                "7 0 obj\n<< /Length 10 >>\nstream\nnot-a-font\nendstream\nendobj\n");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Theory]
    [InlineData("/FontName /Fixture /Flags 32 /FontBBox [0 0 500 700] /ItalicAngle 0 /Ascent 700 /Descent -200 /StemV 80 /FontFile2 7 0 R")]
    [InlineData("/Type /FontDescriptor /FontName /Fixture /Flags 32 /FontBBox [0 0 500 700] /ItalicAngle 0 /Ascent 700 /Descent -200 /CapHeight 700 /FontFile2 7 0 R")]
    [InlineData("/Type /FontDescriptor /FontName /Fixture /Flags 32 /FontBBox [0 0 500 700] /ItalicAngle 0 /Ascent 700 /Descent -200 /StemV 80 /FontFile2 7 0 R")]
    public void StructureInspectorRejectsIncompleteFontDescriptors(string descriptorEntries) {
        byte[] pdf = BuildTrueTypeInspectionPdf(
            BuildMinimalTrueTypeProgram(validLoca: true),
            descriptorEntries);

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Theory]
    [InlineData("/Subtype /TrueType /BaseFont /Fixture")]
    [InlineData("/Type /Font /Subtype /TrueType")]
    public void StructureInspectorRejectsIncompleteSimpleFontDictionaries(string fontEntries) {
        byte[] pdf = BuildTrueTypeInspectionPdf(
            BuildMinimalTrueTypeProgram(validLoca: true),
            fontEntries: fontEntries);

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Theory]
    [InlineData("/DecodeParms [null]")]
    [InlineData("/DecodeParms [<< /ColorTransform 1 >>]")]
    public void ColorInspectorAcceptsSingletonDctFilterAndParameterArrays(string decodeParameters) {
        byte[] pdf = BuildBinaryImageInspectionPdf(
            "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter [/DCTDecode] " + decodeParameters,
            CreateMinimalJpeg(1, 1));

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.Equal(1, evidence.DeviceRgbImageCount);
        Assert.Equal(0, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void StructureInspectorRejectsMultipleEmbeddedProgramEntries() {
        byte[] pdf = BuildTrueTypeInspectionPdf(
            BuildMinimalTrueTypeProgram(validLoca: true),
            "/Type /FontDescriptor /FontName /Fixture /Flags 32 /FontBBox [0 0 500 700] " +
            "/ItalicAngle 0 /Ascent 700 /Descent -200 /CapHeight 700 /StemV 80 /FontFile 7 0 R /FontFile2 7 0 R");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Theory]
    [InlineData("/F1 /Bogus Tf")]
    [InlineData("/F1 [12] Tf")]
    public void StructureInspectorRejectsMalformedFontSizeOperands(string selectFont) {
        byte[] pdf = BuildInspectionPdf(
            "BT " + selectFont + " (A) Tj ET",
            resources: "/Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >>");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(0, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UninspectableFontResourceCount);
    }

    [Fact]
    public void ColorInspectorDoesNotTraverseType3GlyphsAfterMalformedFontSizeOperand() {
        const string charProc = "/DeviceRGB cs 1 0 0 sc 0 0 500 700 re f";
        byte[] pdf = BuildInspectionPdf(
            "BT /F1 /Bogus Tf (A) Tj ET",
            resources: "/Font << /F1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] " +
                "/CharProcs << /A 6 0 R >> /Encoding << /Type /Encoding /Differences [65 /A] >> " +
                "/FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj\n" +
                "6 0 obj\n<< /Length " + charProc.Length + " >>\nstream\n" + charProc + "\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(0, evidence.DeviceRgbOperatorCount);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData(true, 0)]
    [InlineData(false, 1)]
    public void StructureInspectorValidatesTrueTypeGlyphLocations(bool validLoca, int expectedUnembedded) {
        byte[] pdf = BuildTrueTypeInspectionPdf(BuildMinimalTrueTypeProgram(validLoca));

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(expectedUnembedded, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Fact]
    public void StructureInspectorRejectsTruncatedTrueTypeSimpleGlyph() {
        byte[] pdf = BuildTrueTypeInspectionPdf(BuildMinimalTrueTypeProgram(validLoca: true, glyphData: new byte[10]));

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void StructureInspectorRejectsInvalidTrueTypeCompositeReferences(int componentGlyph) {
        byte[] pdf = BuildTrueTypeInspectionPdf(BuildMinimalTrueTypeProgram(
            validLoca: true,
            glyphData: BuildCompositeGlyph(componentGlyph)));

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Fact]
    public void ContentStreamSequenceDecoderPreservesTokenBytesAcrossStreamBoundaries() {
        var streams = new[] {
            new PdfStream(new PdfDictionary(), Encoding.ASCII.GetBytes("1")),
            new PdfStream(new PdfDictionary(), Encoding.ASCII.GetBytes("0 0 m"))
        };

        Assert.True(PdfContentStreamSequenceDecoder.TryDecode(
            streams,
            new Dictionary<int, PdfIndirectObject>(),
            new PdfReadLimits(),
            enforcePageContentLimit: true,
            out string content));
        Assert.Equal("10 0 m", content);
    }

    [Fact]
    public void StructureInspectorRejectsType0FontWithoutValidDescendant() {
        const string type1Program = "%!PS-AdobeFont fixture eexec";
        byte[] pdf = BuildInspectionPdf(
            "BT /F1 12 Tf (A) Tj ET",
            resources: "/Font << /F1 5 0 R >>",
            pageEntries: "/TrimBox [10 10 90 90]",
            extraObjects:
                "5 0 obj\n<< /Type /Font /Subtype /Type0 /BaseFont /Fixture /Encoding /Identity-H /FontDescriptor 6 0 R >>\nendobj\n" +
                "6 0 obj\n<< /Type /FontDescriptor /FontName /Fixture /Flags 32 /FontBBox [0 0 500 700] " +
                "/ItalicAngle 0 /Ascent 700 /Descent -200 /CapHeight 700 /StemV 80 /FontFile 7 0 R >>\nendobj\n" +
                "7 0 obj\n<< /Length " + type1Program.Length + " >>\nstream\n" + type1Program + "\nendstream\nendobj\n");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Theory]
    [InlineData(false, 1)]
    [InlineData(true, 0)]
    public void StructureInspectorRequiresACompleteEncryptedType1Program(bool complete, int expectedUnembedded) {
        string type1Program = complete
            ? BuildValidType1Pfa()
            : "%!PS-AdobeFont-1.0: Fixture 1.0\ncurrentfile eexec\ncleartomark\n";
        byte[] pdf = BuildInspectionPdf(
            "BT /F1 12 Tf (A) Tj ET",
            resources: "/Font << /F1 5 0 R >>",
            pageEntries: "/TrimBox [10 10 90 90]",
            extraObjects:
                "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Fixture /FontDescriptor 6 0 R >>\nendobj\n" +
                "6 0 obj\n<< /Type /FontDescriptor /FontName /Fixture /Flags 32 /FontBBox [0 0 500 700] " +
                "/ItalicAngle 0 /Ascent 700 /Descent -200 /CapHeight 700 /StemV 80 /FontFile 7 0 R >>\nendobj\n" +
                "7 0 obj\n<< /Length " + type1Program.Length + " >>\nstream\n" + type1Program + "\nendstream\nendobj\n");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(expectedUnembedded, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Fact]
    public void StructureInspectorAcceptsACompleteBinaryType1Program() {
        byte[] pdf = BuildType1InspectionPdf(BuildValidType1Pfb());

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(0, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Fact]
    public void StructureInspectorRejectsTokenOnlyType1PrivatePrograms() {
        byte[] privateProgram = Encoding.ASCII.GetBytes(
            "seed/Private 1 dict dup begin\n/CharStrings 1 dict dup begin\n/.notdef 1 RD x ND\nend\nend\n");
        byte[] pdf = BuildType1InspectionPdf(BuildType1Pfa(privateProgram));

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Fact]
    public void StructureInspectorSharesType1OperationBudgetAcrossSubroutineCalls() {
        byte[] pdf = BuildType1InspectionPdf(BuildType1PfaWithRepeatedSubroutineWork(
            callerCount: 100,
            subroutineWorkItems: 400));

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Fact]
    public void StructureInspectorRejectsCff2OpenTypeProgramsForPdfXEmbeddingProof() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCff2Font();
        Assert.NotNull(fontPath);
        byte[] fontProgram = File.ReadAllBytes(fontPath!);
        Assert.True(HasSfntTable(fontProgram, "CFF2"));
        Assert.False(HasSfntTable(fontProgram, "CFF "));
        byte[] pdf = BuildOpenTypeCffInspectionPdf(fontProgram);

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Fact]
    public void ColorInspectorRejectsTruncatedMeshShadingRecords() {
        byte[] pdf = BuildInspectionPdf(
            "/S1 sh",
            resources: "/Shading << /S1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /ShadingType 5 /ColorSpace /DeviceCMYK /BitsPerCoordinate 8 " +
                "/BitsPerComponent 8 /VerticesPerRow 2 /Decode [0 1 0 1 0 1 0 1 0 1 0 1] /Length 1 >>\n" +
                "stream\nx\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("/TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>")]
    [InlineData("/PaintType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>")]
    [InlineData("/PaintType 1 /TilingType 1 /XStep 5 /YStep 5 /Resources << >>")]
    [InlineData("/PaintType 1 /TilingType 1 /BBox [0 0 5 5] /YStep 5 /Resources << >>")]
    [InlineData("/PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /Resources << >>")]
    [InlineData("/PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5")]
    public void ColorInspectorRejectsIncompleteTilingPatternDictionaries(string entries) {
        const string patternContent = "0 0 0 k";
        byte[] pdf = BuildInspectionPdf(
            "/Pattern cs /P1 scn",
            resources: "/Pattern << /P1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /Pattern /PatternType 1 " + entries + " /Length " +
                patternContent.Length + " >>\nstream\n" + patternContent + "\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("")]
    [InlineData("/Type /XObject")]
    public void ColorInspectorRejectsReachablePatternsWithoutThePatternType(string typeEntry) {
        const string patternContent = "0 0 0 k";
        byte[] pdf = BuildInspectionPdf(
            "/Pattern cs /P1 scn",
            resources: "/Pattern << /P1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< " + typeEntry + " /PatternType 1 /PaintType 1 /TilingType 1 " +
                "/BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >> /Length " +
                patternContent.Length + " >>\nstream\n" + patternContent + "\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("/Height 1 /ColorSpace /DeviceCMYK /BitsPerComponent 8")]
    [InlineData("/Width 1 /ColorSpace /DeviceCMYK /BitsPerComponent 8")]
    [InlineData("/Width 1 /Height 1 /ColorSpace /DeviceCMYK")]
    [InlineData("/Width 1 /Height 1 /ColorSpace /DeviceCMYK /BitsPerComponent 8")]
    [InlineData("/Width 1 /Height 1 /ColorSpace /DeviceCMYK /BitsPerComponent 8 /Filter /Unsupported")]
    public void ColorInspectorRejectsMalformedOrTruncatedImageXObjects(string entries) {
        byte[] pdf = BuildInspectionPdf(
            "/Im1 Do",
            resources: "/XObject << /Im1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Image " + entries +
                " /Length 1 >>\nstream\nx\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void MetadataInspectorResolvesIndirectInfoValues() {
        byte[] pdf = BuildIndirectInfoPdf();

        PdfMetadata metadata = PdfReadDocument.Open(pdf).Metadata;

        Assert.Equal("Indirect title", metadata.Title);
        Assert.Equal("Indirect author", metadata.Author);
        Assert.Equal("Indirect subject", metadata.Subject);
        Assert.Equal("alpha, beta", metadata.Keywords);
        Assert.Equal(PdfTrappingStatus.False, metadata.TrappingStatus);
        Assert.Equal(new DateTimeOffset(2026, 8, 26, 7, 0, 0, TimeSpan.Zero), metadata.CreationDate);
        Assert.Equal(new DateTimeOffset(2026, 8, 26, 7, 5, 0, TimeSpan.Zero), metadata.ModificationDate);
        Assert.Equal("PDF/X-4", metadata.PdfXVersion);
        Assert.Equal("PDF/X-4", metadata.PdfXConformance);
    }

    [Fact]
    public void StructureInspectorIgnoresUnusedFontResources() {
        const string charProc = "500 0 0 0 500 700 d1 0 0 500 700 re f";
        byte[] pdf = BuildInspectionPdf(
            "BT /F1 12 Tf (A) Tj ET",
            resources: "/Font << /F1 5 0 R /Unused 7 0 R >>",
            pageEntries: "/TrimBox [10 10 90 90]",
            extraObjects:
                "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] " +
                "/CharProcs << /A 6 0 R >> /Encoding << /Type /Encoding /Differences [65 /A] >> " +
                "/FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj\n" +
                "6 0 obj\n<< /Length " + charProc.Length + " >>\nstream\n" + charProc + "\nendstream\nendobj\n" +
                "7 0 obj\n<< /Type /Font /Subtype /TrueType /BaseFont /Unused >>\nendobj\n");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(0, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
        Assert.True(evidence.IsComplete);
    }

    [Fact]
    public void StructureInspectorFollowsFontSelectionsInInvokedForms() {
        const string formContent = "BT /F2 12 Tf (A) Tj ET";
        const string charProc = "500 0 0 0 500 700 d1 0 0 500 700 re f";
        byte[] pdf = BuildInspectionPdf(
            "/Fm Do",
            resources: "/Font << /Unused 9 0 R >> /XObject << /Fm 5 0 R >>",
            pageEntries: "/TrimBox [10 10 90 90]",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /Font << /F2 6 0 R >> >> /Length " +
                formContent.Length + " >>\nstream\n" + formContent + "\nendstream\nendobj\n" +
                "6 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] " +
                "/CharProcs << /A 7 0 R >> /Encoding << /Type /Encoding /Differences [65 /A] >> " +
                "/FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj\n" +
                "7 0 obj\n<< /Length " + charProc.Length + " >>\nstream\n" + charProc + "\nendstream\nendobj\n" +
                "9 0 obj\n<< /Type /Font /Subtype /TrueType /BaseFont /Unused >>\nendobj\n");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(0, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Fact]
    public void StructureInspectorFollowsFontSelectionsInPatternsAndSoftMasks() {
        const string selectedText = "BT /F1 12 Tf (A) Tj ET";
        const string charProc = "500 0 0 0 500 700 d1 0 0 500 700 re f";
        string type3Font =
            "<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] " +
            "/CharProcs << /A {0} 0 R >> /Encoding << /Type /Encoding /Differences [65 /A] >> " +
            "/FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>";
        byte[] pdf = BuildInspectionPdf(
            "/Pattern cs /P1 scn /GS1 gs",
            resources: "/Pattern << /P1 5 0 R >> /ExtGState << /GS1 << /SMask << /S /Luminosity /G 8 0 R >> >> >>",
            pageEntries: "/TrimBox [10 10 90 90]",
            extraObjects:
                "5 0 obj\n<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 " +
                "/Resources << /Font << /F1 6 0 R >> >> /Length " + selectedText.Length + " >>\nstream\n" + selectedText + "\nendstream\nendobj\n" +
                "6 0 obj\n" + string.Format(type3Font, 7) + "\nendobj\n" +
                "7 0 obj\n<< /Length " + charProc.Length + " >>\nstream\n" + charProc + "\nendstream\nendobj\n" +
                "8 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Group << /S /Transparency >> " +
                "/Resources << /Font << /F1 9 0 R >> >> /Length " + selectedText.Length + " >>\nstream\n" + selectedText + "\nendstream\nendobj\n" +
                "9 0 obj\n" + string.Format(type3Font, 10) + "\nendobj\n" +
                "10 0 obj\n<< /Length " + charProc.Length + " >>\nstream\n" + charProc + "\nendstream\nendobj\n");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(2, evidence.FontResourceCount);
        Assert.Equal(0, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Theory]
    [InlineData("")]
    [InlineData("/Type /XObject")]
    public void StructureInspectorRejectsReachableTilingPatternsWithoutThePatternType(string typeEntry) {
        byte[] pdf = BuildInspectionPdf(
            "/Pattern cs /P1 scn",
            resources: "/Pattern << /P1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< " + typeEntry + " /PatternType 1 /PaintType 1 /TilingType 1 " +
                "/BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << >> /Length 0 >>\n" +
                "stream\n\nendstream\nendobj\n");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.True(evidence.UninspectableFontResourceCount > 0);
    }

    [Fact]
    public void StructureInspectorBoundsIndirectFontResourceGraphTraversal() {
        const int firstObject = 5;
        const int lastObject = 40;
        var extraObjects = new StringBuilder();
        for (int objectNumber = firstObject; objectNumber <= lastObject; objectNumber++) {
            extraObjects.Append(objectNumber).Append(" 0 obj\n");
            if (objectNumber < lastObject) {
                extraObjects.Append(objectNumber + 1).Append(" 0 R");
            } else {
                extraObjects.Append("<< /Type /Font /Subtype /TrueType /BaseFont /Deep >>");
            }
            extraObjects.Append("\nendobj\n");
        }
        byte[] pdf = BuildInspectionPdf(
            "BT /F1 12 Tf (A) Tj ET",
            resources: "/Font << /F1 5 0 R >>",
            extraObjects: extraObjects.ToString());
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxObjectNestingDepth = 8 }
        };
        PdfReadDocument document = PdfReadDocument.Open(pdf, options);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            document.InspectPrintProductionStructure());

        Assert.Equal(PdfReadLimitKind.ObjectNestingDepth, exception.Kind);
        Assert.Equal(8, exception.Limit);
        Assert.Equal(9, exception.Actual);
    }

    [Fact]
    public void ColorInspectorFailsClosedWhenReachableResourceTraversalExceedsObjectDepth() {
        const int firstObject = 5;
        const int lastObject = 40;
        var extraObjects = new StringBuilder();
        for (int objectNumber = firstObject; objectNumber <= lastObject; objectNumber++) {
            extraObjects.Append(objectNumber).Append(" 0 obj\n");
            if (objectNumber < lastObject) {
                extraObjects.Append(objectNumber + 1).Append(" 0 R");
            } else {
                extraObjects.Append("<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length 0 >>\nstream\n\nendstream");
            }
            extraObjects.Append("\nendobj\n");
        }
        byte[] pdf = BuildInspectionPdf(
            "/Deep Do",
            resources: "/XObject << /Deep 5 0 R >>",
            extraObjects: extraObjects.ToString());
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxObjectNestingDepth = 8 }
        };

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf, options).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("99 0 R")]
    [InlineData("[4 0 R 99 0 R]")]
    [InlineData("[4 0 R 5]")]
    public void PrintProductionInspectorsFailClosedOnMalformedPageContents(string contents) {
        byte[] pdf = BuildInspectionPdf(string.Empty, contents: contents);
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        PdfPrintProductionColorEvidence colors = document.InspectPrintProductionColors();
        PdfPrintProductionStructureEvidence structure = document.InspectPrintProductionStructure();

        Assert.False(colors.IsComplete);
        Assert.True(colors.UninspectableContentStreamCount > 0);
        Assert.False(structure.IsComplete);
        Assert.True(structure.UninspectableFontResourceCount > 0);
    }

    [Fact]
    public void ColorInspectorFailsClosedOnUnterminatedContentString() {
        byte[] pdf = BuildInspectionPdf("(unterminated");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("/ca /Bogus")]
    [InlineData("/CA 2")]
    [InlineData("/ca -0.1")]
    public void ColorInspectorFailsClosedOnMalformedGraphicsStateAlpha(string graphicsState) {
        byte[] pdf = BuildInspectionPdf(
            "/GS1 gs",
            resources: "/ExtGState << /GS1 << " + graphicsState + " >> >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorClassifiesTransparencyGroupColorSpaceAliases() {
        byte[] pdf = BuildInspectionPdf(
            "/Fm1 Do",
            resources: "/ColorSpace << /BlendRgb /DeviceRGB >> /XObject << /Fm1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] " +
                "/Group << /S /Transparency /CS /BlendRgb >> /Resources << /ColorSpace << /BlendRgb /DeviceRGB >> >> " +
                "/Length 0 >>\nstream\n\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.True(evidence.IsComplete);
        Assert.True(evidence.HasDeviceRgbUsage);
        Assert.Equal(1, evidence.DeviceRgbTransparencyGroupCount);
        Assert.Equal(1, evidence.TransparencyGroupCount);
    }

    [Fact]
    public void ColorInspectorFailsClosedOnUnknownPageGroupSubtype() {
        byte[] pdf = BuildInspectionPdf(
            string.Empty,
            pageEntries: "/Group << /S /Unknown >>");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void ColorInspectorFailsClosedOnUnknownFormGroupSubtype() {
        byte[] pdf = BuildInspectionPdf(
            "/Fm1 Do",
            resources: "/XObject << /Fm1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Group << /S /Unknown >> /Length 0 >>\nstream\n\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Theory]
    [InlineData("/Type /Bogus")]
    [InlineData("/I 1")]
    [InlineData("/K /False")]
    public void ColorInspectorFailsClosedOnMalformedTransparencyGroupOptions(string groupEntries) {
        byte[] pdf = BuildInspectionPdf(
            "/Fm1 Do",
            resources: "/XObject << /Fm1 5 0 R >>",
            extraObjects:
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] " +
                "/Group << /S /Transparency " + groupEntries + " >> /Length 0 >>\nstream\n\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.UninspectableContentStreamCount);
    }

    [Fact]
    public void StructureInspectorInspectsOnlyPaintedType3CharacterProcedures() {
        const string painted = "500 0 0 0 500 700 d1 0 0 500 700 re f";
        byte[] pdf = BuildInspectionPdf(
            "BT /F1 12 Tf (A) Tj ET",
            resources: "/Font << /F1 5 0 R >>",
            pageEntries: "/TrimBox [10 10 90 90]",
            extraObjects:
                "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] " +
                "/CharProcs << /A 6 0 R /B 7 0 R >> /Encoding << /Type /Encoding /Differences [65 /A /B] >> " +
                "/FirstChar 65 /LastChar 66 /Widths [500 500] /Resources << >> >>\nendobj\n" +
                "6 0 obj\n<< /Length " + painted.Length + " >>\nstream\n" + painted + "\nendstream\nendobj\n" +
                "7 0 obj\n<< /Filter /Unsupported /Length 3 >>\nstream\nbad\nendstream\nendobj\n");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.True(evidence.IsComplete);
        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(0, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Theory]
    [InlineData("")]
    [InlineData("0 0 m 10 10 l S")]
    [InlineData("0 0 500 d1")]
    public void StructureInspectorRejectsSelectedType3GlyphWithoutValidInitialMetrics(string glyphContent) {
        byte[] pdf = BuildInspectionPdf(
            "BT /F1 12 Tf (A) Tj ET",
            resources: "/Font << /F1 5 0 R >>",
            pageEntries: "/TrimBox [10 10 90 90]",
            extraObjects:
                "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] " +
                "/CharProcs << /A 6 0 R >> /Encoding << /Type /Encoding /Differences [65 /A] >> " +
                "/FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj\n" +
                "6 0 obj\n<< /Length " + glyphContent.Length + " >>\nstream\n" + glyphContent + "\nendstream\nendobj\n");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.False(evidence.IsComplete);
        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
    }

    [Theory]
    [InlineData("")]
    [InlineData("/FontBBox [0 0 500]")]
    [InlineData("/FontMatrix [0 0 0 0 0 0]")]
    [InlineData("/Widths [500 500]")]
    public void StructureInspectorRejectsMalformedType3FontDictionary(string replacement) {
        const string glyphContent = "500 0 0 0 500 700 d1 0 0 500 700 re f";
        string requiredEntries =
            "/FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /Widths [500]";
        if (replacement.Length == 0) {
            requiredEntries = "/FontMatrix [0.001 0 0 0.001 0 0] /Widths [500]";
        } else if (replacement.StartsWith("/FontBBox", StringComparison.Ordinal)) {
            requiredEntries = replacement + " /FontMatrix [0.001 0 0 0.001 0 0] /Widths [500]";
        } else if (replacement.StartsWith("/FontMatrix", StringComparison.Ordinal)) {
            requiredEntries = "/FontBBox [0 0 500 700] " + replacement + " /Widths [500]";
        } else {
            requiredEntries = "/FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] " + replacement;
        }
        byte[] pdf = BuildInspectionPdf(
            "BT /F1 12 Tf (A) Tj ET",
            resources: "/Font << /F1 5 0 R >>",
            pageEntries: "/TrimBox [10 10 90 90]",
            extraObjects:
                "5 0 obj\n<< /Type /Font /Subtype /Type3 " + requiredEntries +
                " /CharProcs << /A 6 0 R >> /Encoding << /Type /Encoding /Differences [65 /A] >> " +
                "/FirstChar 65 /LastChar 65 /Resources << >> >>\nendobj\n" +
                "6 0 obj\n<< /Length " + glyphContent.Length + " >>\nstream\n" + glyphContent + "\nendstream\nendobj\n");

        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionStructure();

        Assert.Equal(1, evidence.UnembeddedFontResourceCount);
    }

    [Fact]
    public void XmpInspectorReadsSimplePropertiesFromRdfDescriptionAttributes() {
        const string xmp = """
            <?xpacket begin="﻿"?>
            <x:xmpmeta xmlns:x="adobe:ns:meta/">
              <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
                <rdf:Description
                  xmlns:pdfxid="http://www.npes.org/pdfx/ns/id/"
                  xmlns:xmp="http://ns.adobe.com/xap/1.0/"
                  xmlns:xmpMM="http://ns.adobe.com/xap/1.0/mm/"
                  xmlns:pdf="http://ns.adobe.com/pdf/1.3/"
                  pdfxid:GTS_PDFXVersion="PDF/X-4"
                  pdfxid:GTS_PDFXConformance="PDF/X-4"
                  xmp:CreateDate="2026-08-25T10:00:00Z"
                  xmp:ModifyDate="2026-08-25T10:05:00Z"
                  xmp:MetadataDate="2026-08-25T10:05:00Z"
                  xmpMM:DocumentID="uuid:11111111-1111-1111-1111-111111111111"
                  xmpMM:InstanceID="uuid:22222222-2222-2222-2222-222222222222"
                  xmpMM:VersionID="7"
                  xmpMM:RenditionClass="proof"
                  pdf:Trapped="False" />
              </rdf:RDF>
            </x:xmpmeta>
            <?xpacket end="w"?>
            """;
        byte[] pdf = BuildXmpInspectionPdf(xmp);

        PdfXmpMetadataInfo metadata = Assert.IsType<PdfXmpMetadataInfo>(PdfReadDocument.Open(pdf).XmpMetadata);

        Assert.Equal("PDF/X-4", metadata.PdfXVersion);
        Assert.Equal("PDF/X-4", metadata.PdfXConformance);
        Assert.Equal(new DateTimeOffset(2026, 8, 25, 10, 0, 0, TimeSpan.Zero), metadata.CreationDate);
        Assert.Equal(new DateTimeOffset(2026, 8, 25, 10, 5, 0, TimeSpan.Zero), metadata.ModificationDate);
        Assert.Equal(metadata.ModificationDate, metadata.MetadataDate);
        Assert.Equal("uuid:11111111-1111-1111-1111-111111111111", metadata.DocumentId);
        Assert.Equal("uuid:22222222-2222-2222-2222-222222222222", metadata.InstanceId);
        Assert.Equal("7", metadata.VersionId);
        Assert.Equal("proof", metadata.RenditionClass);
        Assert.Equal(PdfTrappingStatus.False, metadata.TrappingStatus);
    }

    [Fact]
    public void XmpInspectorRejectsNonXmpDateLexicalForms() {
        const string xmp = """
            <?xpacket begin="﻿"?>
            <x:xmpmeta xmlns:x="adobe:ns:meta/">
              <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
                <rdf:Description xmlns:xmp="http://ns.adobe.com/xap/1.0/"
                  xmp:CreateDate="08/26/2026 12:00:00 +00:00"
                  xmp:ModifyDate="2026/08/26 12:05:00Z"
                  xmp:MetadataDate="26 Aug 2026 12:05:00 GMT" />
              </rdf:RDF>
            </x:xmpmeta>
            <?xpacket end="w"?>
            """;

        PdfXmpMetadataInfo metadata = Assert.IsType<PdfXmpMetadataInfo>(
            PdfReadDocument.Open(BuildXmpInspectionPdf(xmp)).XmpMetadata);

        Assert.Null(metadata.CreationDate);
        Assert.Null(metadata.ModificationDate);
        Assert.Null(metadata.MetadataDate);
    }

    private static byte[] BuildInspectionPdf(
        string content,
        string resources = "",
        string pageEntries = "",
        string extraObjects = "",
        string pageTreeEntries = "",
        string contents = "4 0 R") {
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] " + pageTreeEntries + " >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << " + resources + " >> " + pageEntries + " /Contents " + contents + " >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n" + extraObjects + "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static string BuildIccProfileObject(int objectNumber, int components, byte[] profile) {
        string hex = BitConverter.ToString(profile).Replace("-", string.Empty) + ">";
        return objectNumber + " 0 obj\n<< /N " + components +
            " /Filter /ASCIIHexDecode /Length " + hex.Length + " >>\nstream\n" +
            hex + "\nendstream\nendobj\n";
    }

    private static string BuildDeviceNTintFunctionObject() =>
        "5 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0 0] /C1 [1 0 0 0] /N 1 >>\nendobj\n";

    private static byte[] BuildType1InspectionPdf(byte[] type1Program) {
        byte[] content = Encoding.ASCII.GetBytes("BT /F1 12 Tf (A) Tj ET");
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /TrimBox [10 10 90 90] /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + content.Length + " >>\nstream\n");
        output.Write(content, 0, content.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Fixture /FontDescriptor 6 0 R >>\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /Type /FontDescriptor /FontName /Fixture /Flags 32 /FontBBox [0 0 500 700] /ItalicAngle 0 /Ascent 700 /Descent -200 /CapHeight 700 /StemV 80 /FontFile 7 0 R >>\nendobj\n");
        WriteAscii(output, "7 0 obj\n<< /Length " + type1Program.Length + " >>\nstream\n");
        output.Write(type1Program, 0, type1Program.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildTrueTypeInspectionPdf(
        byte[] trueTypeProgram,
        string? descriptorEntries = null,
        string? fontEntries = null) {
        byte[] content = Encoding.ASCII.GetBytes("BT /F1 12 Tf (A) Tj ET");
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /TrimBox [10 10 90 90] /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + content.Length + " >>\nstream\n");
        output.Write(content, 0, content.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        fontEntries ??= "/Type /Font /Subtype /TrueType /BaseFont /Fixture";
        WriteAscii(output, "5 0 obj\n<< " + fontEntries + " /FontDescriptor 6 0 R >>\nendobj\n");
        descriptorEntries ??= "/Type /FontDescriptor /FontName /Fixture /Flags 32 /FontBBox [0 0 500 700] /ItalicAngle 0 /Ascent 700 /Descent -200 /CapHeight 700 /StemV 80 /FontFile2 7 0 R";
        WriteAscii(output, "6 0 obj\n<< " + descriptorEntries + " >>\nendobj\n");
        WriteAscii(output, "7 0 obj\n<< /Length " + trueTypeProgram.Length + " >>\nstream\n");
        output.Write(trueTypeProgram, 0, trueTypeProgram.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildBinaryImageInspectionPdf(string imageEntries, byte[] imageData) {
        const string content = "/Im1 Do";
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 " + imageEntries + " /Length " + imageData.Length + " >>\nstream\n");
        output.Write(imageData, 0, imageData.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] CreateMinimalJpeg(int width, int height) {
        byte[] rgba = new byte[checked(width * height * 4)];
        for (int offset = 0; offset < rgba.Length; offset += 4) {
            rgba[offset] = 255;
            rgba[offset + 3] = 255;
        }
        return OfficeJpegCodec.Encode(
            OfficeRasterImage.FromRgba32(width, height, rgba),
            new OfficeJpegEncodeOptions {
                Quality = 100,
                Subsampling = OfficeJpegSubsampling.Y444
            });
    }

    private static byte[] BuildOpenTypeCffInspectionPdf(byte[] openTypeProgram) {
        byte[] content = Encoding.ASCII.GetBytes("BT /F1 12 Tf (A) Tj ET");
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /TrimBox [10 10 90 90] /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + content.Length + " >>\nstream\n");
        output.Write(content, 0, content.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Fixture /FontDescriptor 6 0 R >>\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /Type /FontDescriptor /FontName /Fixture /Flags 32 /FontBBox [0 0 500 700] /ItalicAngle 0 /Ascent 700 /Descent -200 /CapHeight 700 /StemV 80 /FontFile3 7 0 R >>\nendobj\n");
        WriteAscii(output, "7 0 obj\n<< /Subtype /OpenType /Length " + openTypeProgram.Length + " >>\nstream\n");
        output.Write(openTypeProgram, 0, openTypeProgram.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static bool HasSfntTable(byte[] data, string tag) {
        if (data.Length < 12 || tag.Length != 4) return false;
        int tableCount = (data[4] << 8) | data[5];
        if (tableCount <= 0 || 12L + tableCount * 16L > data.Length) return false;
        for (int index = 0; index < tableCount; index++) {
            int record = 12 + index * 16;
            if (data[record] == tag[0] &&
                data[record + 1] == tag[1] &&
                data[record + 2] == tag[2] &&
                data[record + 3] == tag[3]) return true;
        }
        return false;
    }

    private static byte[] BuildMinimalTrueTypeProgram(bool validLoca, byte[]? glyphData = null) {
        glyphData ??= BuildValidSimpleGlyph();
        const int directoryLength = 12 + 4 * 16;
        const int headOffset = directoryLength;
        const int maxpOffset = 132;
        const int locaOffset = 140;
        int locaLength = validLoca ? 8 : 4;
        int glyfOffset = locaOffset + locaLength;
        var data = new byte[glyfOffset + glyphData.Length];
        WriteUInt32BigEndian(data, 0, 0x00010000U);
        WriteUInt16BigEndian(data, 4, 4);
        WriteTableRecord(data, 12, "head", headOffset, 54);
        WriteTableRecord(data, 28, "maxp", maxpOffset, 6);
        WriteTableRecord(data, 44, "loca", locaOffset, locaLength);
        WriteTableRecord(data, 60, "glyf", glyfOffset, glyphData.Length);
        WriteUInt16BigEndian(data, headOffset + 50, 1); // Long loca offsets.
        WriteUInt16BigEndian(data, maxpOffset + 4, 1); // One glyph plus two loca entries.
        if (validLoca) WriteUInt32BigEndian(data, locaOffset + 4, checked((uint)glyphData.Length));
        Buffer.BlockCopy(glyphData, 0, data, glyfOffset, glyphData.Length);
        return data;
    }

    private static byte[] BuildValidSimpleGlyph() {
        var glyph = new byte[16];
        WriteUInt16BigEndian(glyph, 0, 1); // One contour.
        WriteUInt16BigEndian(glyph, 10, 0); // The contour ends at the first point.
        WriteUInt16BigEndian(glyph, 12, 0); // No instructions.
        glyph[14] = 0x31; // On-curve point with unchanged X/Y coordinates.
        return glyph;
    }

    private static byte[] BuildCompositeGlyph(int componentGlyph) {
        var glyph = new byte[16];
        WriteUInt16BigEndian(glyph, 0, -1);
        WriteUInt16BigEndian(glyph, 10, 0); // Final component, byte-sized arguments, no transform.
        WriteUInt16BigEndian(glyph, 12, componentGlyph);
        return glyph;
    }

    private static void WriteTableRecord(byte[] data, int offset, string tag, int tableOffset, int tableLength) {
        for (int index = 0; index < 4; index++) data[offset + index] = (byte)tag[index];
        WriteUInt32BigEndian(data, offset + 8, checked((uint)tableOffset));
        WriteUInt32BigEndian(data, offset + 12, checked((uint)tableLength));
    }

    private static void WriteUInt16BigEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

    private static void WriteUInt32BigEndian(byte[] data, int offset, uint value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static byte[] BuildXmpInspectionPdf(string xmp) {
        byte[] metadataBytes = Encoding.UTF8.GetBytes(xmp);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /Metadata 5 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /Metadata /Subtype /XML /Length " + metadataBytes.Length + " >>\nstream\n");
        output.Write(metadataBytes, 0, metadataBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildIndirectInfoPdf() {
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n");
        WriteAscii(output, "5 2 obj\n<< /Title 6 0 R /Author 7 0 R /Subject 8 0 R /Keywords 9 0 R /Trapped 10 0 R " +
            "/CreationDate 11 0 R /ModDate 12 0 R /GTS_PDFXVersion 13 0 R /GTS_PDFXConformance 14 0 R >>\nendobj\n");
        WriteAscii(output, "6 0 obj\n(Indirect title)\nendobj\n7 0 obj\n(Indirect author)\nendobj\n" +
            "8 0 obj\n(Indirect subject)\nendobj\n9 0 obj\n(alpha, beta)\nendobj\n10 0 obj\n/False\nendobj\n" +
            "11 0 obj\n(D:20260826070000Z)\nendobj\n12 0 obj\n(D:20260826070500Z)\nendobj\n" +
            "13 0 obj\n(PDF/X-4)\nendobj\n14 0 obj\n(PDF/X-4)\nendobj\n");
        WriteAscii(output, "trailer\n<< /Root 1 0 R /Info 5 2 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static string BuildValidType1Pfa() {
        byte[] encrypted = EncryptType1PrivateProgram();
        string hex = BitConverter.ToString(encrypted).Replace("-", string.Empty);
        return "%!PS-AdobeFont-1.0: Fixture 1.0\ncurrentfile eexec\n" + hex + "\ncleartomark\n";
    }

    private static byte[] BuildType1Pfa(byte[] privateProgram) {
        byte[] encrypted = EncryptType1Eexec(privateProgram);
        string hex = BitConverter.ToString(encrypted).Replace("-", string.Empty);
        return Encoding.ASCII.GetBytes(
            "%!PS-AdobeFont-1.0: Fixture 1.0\ncurrentfile eexec\n" + hex + "\ncleartomark\n");
    }

    private static byte[] BuildType1PfaWithRepeatedSubroutineWork(int callerCount, int subroutineWorkItems) {
        using var subroutine = new MemoryStream();
        for (int index = 0; index < subroutineWorkItems; index++) {
            subroutine.WriteByte(139);
            subroutine.WriteByte(139);
            subroutine.WriteByte(1);
        }
        subroutine.WriteByte(11);

        using var glyph = new MemoryStream();
        for (int index = 0; index < callerCount; index++) {
            glyph.WriteByte(139);
            glyph.WriteByte(10);
        }
        glyph.WriteByte(14);

        using var privateProgram = new MemoryStream();
        WriteAscii(privateProgram, "seed/Private 1 dict dup begin\n/lenIV -1 def\n/Subrs 1 array\ndup 0 " + subroutine.Length + " RD ");
        subroutine.Position = 0;
        subroutine.CopyTo(privateProgram);
        WriteAscii(privateProgram, " NP\n/CharStrings 1 dict dup begin\n/.notdef " + glyph.Length + " RD ");
        glyph.Position = 0;
        glyph.CopyTo(privateProgram);
        WriteAscii(privateProgram, " ND\nend\nend\n");
        return BuildType1Pfa(privateProgram.ToArray());
    }

    private static byte[] BuildValidType1Pfb() {
        byte[] header = Encoding.ASCII.GetBytes("%!PS-AdobeFont-1.0: Fixture 1.0\ncurrentfile eexec\n");
        byte[] encrypted = EncryptType1PrivateProgram();
        byte[] trailer = Encoding.ASCII.GetBytes("\ncleartomark\n");
        using var output = new MemoryStream();
        WritePfbSegment(output, 1, header);
        WritePfbSegment(output, 2, encrypted);
        WritePfbSegment(output, 1, trailer);
        output.WriteByte(0x80);
        output.WriteByte(0x03);
        return output.ToArray();
    }

    private static byte[] EncryptType1PrivateProgram() {
        byte[] charString = EncryptType1CharString(new byte[] {
            0, 0, 0, 0,
            139,
            248, 136,
            13,
            14
        });
        using var privateProgramStream = new MemoryStream();
        WriteAscii(
            privateProgramStream,
            "seed/Private 1 dict dup begin\n/lenIV 4 def\n/CharStrings 1 dict dup begin\n/.notdef " +
            charString.Length + " RD ");
        privateProgramStream.Write(charString, 0, charString.Length);
        WriteAscii(privateProgramStream, " ND\nend\nend\n");
        byte[] privateProgram = privateProgramStream.ToArray();
        return EncryptType1Eexec(privateProgram);
    }

    private static byte[] EncryptType1Eexec(byte[] privateProgram) {
        var encrypted = new byte[privateProgram.Length];
        ushort state = 55665;
        for (int index = 0; index < privateProgram.Length; index++) {
            byte cipher = (byte)(privateProgram[index] ^ (state >> 8));
            encrypted[index] = cipher;
            state = unchecked((ushort)((cipher + state) * 52845 + 22719));
        }
        return encrypted;
    }

    private static byte[] EncryptType1CharString(byte[] plainText) {
        var encrypted = new byte[plainText.Length];
        ushort state = 4330;
        for (int index = 0; index < plainText.Length; index++) {
            byte cipher = (byte)(plainText[index] ^ (state >> 8));
            encrypted[index] = cipher;
            state = unchecked((ushort)((cipher + state) * 52845 + 22719));
        }
        return encrypted;
    }

    private static void WritePfbSegment(Stream output, byte type, byte[] data) {
        output.WriteByte(0x80);
        output.WriteByte(type);
        uint length = checked((uint)data.Length);
        output.WriteByte((byte)length);
        output.WriteByte((byte)(length >> 8));
        output.WriteByte((byte)(length >> 16));
        output.WriteByte((byte)(length >> 24));
        output.Write(data, 0, data.Length);
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }
}
