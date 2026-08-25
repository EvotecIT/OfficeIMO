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
    public void StructureInspectorBoundsIndirectFontResourceGraphTraversal() {
        const int firstObject = 5;
        const int lastObject = 40;
        var extraObjects = new StringBuilder();
        for (int objectNumber = firstObject; objectNumber <= lastObject; objectNumber++) {
            extraObjects.Append(objectNumber).Append(" 0 obj\n<<");
            if (objectNumber < lastObject) {
                extraObjects.Append(" /Next ").Append(objectNumber + 1).Append(" 0 R");
            }
            extraObjects.Append(" >>\nendobj\n");
        }
        byte[] pdf = BuildInspectionPdf(string.Empty, pageEntries: "/Audit 5 0 R", extraObjects: extraObjects.ToString());
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
