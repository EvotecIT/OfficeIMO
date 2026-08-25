using System.Text;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfPrintProductionInspectorRegressionTests {
    [Fact]
    public void ColorInspectorFindsDirectResourceAndPatternShadingDictionaries() {
        byte[] pdf = BuildInspectionPdf(
            "/S1 sh",
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

    [Fact]
    public void ColorInspectorAppliesPatternResourcesToNestedForms() {
        const string patternContent = "/Fm Do";
        const string formContent = "/PrintRgb cs 1 0 0 sc";
        byte[] pdf = BuildInspectionPdf(
            string.Empty,
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
                "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length " +
                formContent.Length + " >>\nstream\n" + formContent + "\nendstream\nendobj\n");

        PdfPrintProductionColorEvidence evidence = PdfReadDocument.Open(pdf).InspectPrintProductionColors();

        Assert.Equal(2, evidence.DeviceRgbOperatorCount);
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

    [Fact]
    public void StructureInspectorIgnoresUnusedFontResources() {
        const string charProc = "0 0 500 700 d1 0 0 500 700 re f";
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
        const string charProc = "0 0 500 700 d1 0 0 500 700 re f";
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
        const string charProc = "0 0 500 700 d1 0 0 500 700 re f";
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

    private static byte[] BuildInspectionPdf(
        string content,
        string resources = "",
        string pageEntries = "",
        string extraObjects = "") {
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 100 100] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << " + resources + " >> " + pageEntries + " /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n" + extraObjects + "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
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

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }
}
