using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void RecognizedSvgContentOverridesAMisleadingPdfExtension() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\"><metadata>" +
            Encoding.UTF8.GetString(CreateXmpPacket()) +
            "</metadata></svg>");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "misleading.pdf");

        Assert.True(result.Before.HasGenerativeAiDeclaration);
        Assert.True(result.WasChanged);
        Assert.False(result.After.HasGenerativeAiDeclaration);
    }

    [Fact]
    public void XmpNodeBudgetFailuresAreNotReportedAsParseMisses() {
        var options = new OfficeProvenanceOptions { MaxContainerEntries = 1 };

        Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceXml.TryLoadDocument(CreateXmpPacket(), options, out _));
    }

    [Fact]
    public void TiffPreservesXmpThatOverlapsSubIfdImageData() {
        byte[] xmp = CreateXmpPacket();
        const int primaryIfdOffset = 8;
        const int subIfdOffset = 38;
        const int xmpOffset = 68;
        byte[] tiff = new byte[xmpOffset + xmp.Length];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = primaryIfdOffset;
        tiff[primaryIfdOffset] = 2;
        WriteLittleEndianEntry(tiff, primaryIfdOffset + 2, 330, 4, 1, subIfdOffset);
        WriteLittleEndianEntry(tiff, primaryIfdOffset + 14, 700, 1, xmp.Length, xmpOffset);
        tiff[subIfdOffset] = 2;
        WriteLittleEndianEntry(tiff, subIfdOffset + 2, 273, 4, 1, xmpOffset);
        WriteLittleEndianEntry(tiff, subIfdOffset + 14, 279, 4, 1, xmp.Length);
        Buffer.BlockCopy(xmp, 0, tiff, xmpOffset, xmp.Length);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tif");

        Assert.False(result.WasChanged);
        Assert.Equal(tiff, result.ToArray());
        Assert.True(result.After.HasGenerativeAiDeclaration);
    }
}
