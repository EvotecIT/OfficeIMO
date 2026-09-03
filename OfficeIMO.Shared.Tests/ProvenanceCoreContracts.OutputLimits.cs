using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void XmpRewriteStopsAtTheOutputBoundaryDuringSerialization() {
        byte[] xmp = CreateXmpPacket();
        byte[] header = Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE1, Join(header, xmp)),
            new byte[] { 0xFF, 0xD9 });
        var options = new OfficeProvenanceRemovalOptions {
            MaxOutputBytes = 64
        };
        options.Limits.MaxAssetBytes = jpeg.Length;
        options.Limits.MaxManifestBytes = jpeg.Length;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg", options));

        Assert.True(OfficeProvenanceLimitException.IsOutput(exception));
    }
}
