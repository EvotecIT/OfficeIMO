using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
    [Fact]
    public void GifProcessesXmpPacketSplitAcrossDataSubBlocks() {
        byte[] xmp = Encoding.UTF8.GetBytes(
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
            "<rdf:Description xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\" " +
            "iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/>" +
            "</rdf:RDF></x:xmpmeta>");
        byte[] gif = CreateGifWithSubBlockedXmp(xmp);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(gif, "fixture.gif");

        Assert.Single(result.Before.Evidence);
        Assert.True(result.WasChanged);
        Assert.True(result.WasReserialized);
        Assert.Empty(result.After.Evidence);
        Assert.Equal((byte)0x3B, result.ToArray()[result.ToArray().Length - 1]);
    }

    private static byte[] CreateGifWithSubBlockedXmp(byte[] packet) {
        using var payload = new MemoryStream();
        for (int offset = 0; offset < packet.Length;) {
            int length = Math.Min(127, packet.Length - offset);
            payload.WriteByte((byte)length);
            payload.Write(packet, offset, length);
            offset += length;
        }
        byte[] trailer = new byte[258];
        trailer[0] = 0x01;
        for (int index = 1; index <= 255; index++) trailer[index] = checked((byte)(256 - index));
        return Join(
            Encoding.ASCII.GetBytes("GIF89a"),
            new byte[] { 1, 0, 1, 0, 0, 0, 0 },
            new byte[] { 0x21, 0xFF, 0x0B },
            Encoding.ASCII.GetBytes("XMP DataXMP"),
            payload.ToArray(),
            trailer,
            new byte[] { 0x3B });
    }
}
