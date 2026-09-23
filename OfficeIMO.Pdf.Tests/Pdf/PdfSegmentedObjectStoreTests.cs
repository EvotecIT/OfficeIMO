using OfficeIMO.Pdf;
using System.Security.Cryptography;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSegmentedObjectStoreTests {
    [Fact]
    public void SegmentedObjectRoundTripsInMemoryAndAfterSpilling() {
        using var store = new PdfObjectStore(memoryLimitBytes: 10);
        store.AddSegments(new byte[] { 1, 2 }, new byte[] { 3 }, new byte[0]);
        Assert.False(store.IsSpilled);
        Assert.Equal(new byte[] { 1, 2, 3 }, store[0]);
        Assert.Equal(0, store.IndexOf(new byte[] { 1, 2, 3 }));

        store.AddSegments(new byte[] { 4, 5, 6, 7 }, new byte[] { 8, 9, 10, 11 });
        Assert.True(store.IsSpilled);
        using var sha = SHA256.Create();
        using var destination = new MemoryStream();
        store.CopyTo(0, destination, sha);
        store.CopyTo(1, destination, sha);
        Assert.Equal(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 }, destination.ToArray());
        Assert.Equal(new byte[] { 1, 2, 3 }, store[0]);

        store[0] = new byte[] { 9 };
        Assert.Equal(new byte[] { 9 }, store[0]);
    }

    [Fact]
    public void SegmentedObjectHashesAndCopiesLikeTheJoinedBytes() {
        using var segmented = new PdfObjectStore();
        segmented.AddSegments(new byte[] { 1, 2 }, new byte[] { 3, 4 });
        using var joined = new PdfObjectStore();
        joined.AddSegments(new byte[] { 1, 2 }, new byte[] { 3, 4 });
        joined[0] = joined[0];

        using var segmentedHash = SHA256.Create();
        using var joinedHash = SHA256.Create();
        using var segmentedCopy = new MemoryStream();
        using var joinedCopy = new MemoryStream();
        segmented.CopyTo(0, segmentedCopy, segmentedHash);
        joined.CopyTo(0, joinedCopy, joinedHash);
        segmentedHash.TransformFinalBlock(new byte[0], 0, 0);
        joinedHash.TransformFinalBlock(new byte[0], 0, 0);

        Assert.Equal(joinedHash.Hash, segmentedHash.Hash);
        Assert.Equal(joinedCopy.ToArray(), segmentedCopy.ToArray());
    }

    [Fact]
    public void OutputDoesNotDependOnWhereTheObjectStoreSpills() {
        byte[] reference = Render(PdfObjectStore.DefaultMemoryLimitBytes);
        Assert.Equal(reference, Render(PdfObjectStore.DefaultMemoryLimitBytes));
        foreach (long limit in new long[] { 0, 3000, 20000, 60000 }) {
            Assert.Equal(reference, Render(limit));
        }
    }

    private static byte[] Render(long objectBufferLimit) {
        var document = PdfDocument.Create(new PdfOptions { ObjectBufferMemoryLimitBytes = objectBufferLimit });
        for (int i = 0; i < 40; i++) {
            document.Paragraph(p => p.Text("Segmented object store transition paragraph " + i + new string('x', 400)));
        }
        document.Table(Enumerable.Range(0, 200).Select(r => new[] { "r" + r, "value " + r, new string('y', r % 50) }).ToArray());
        return document.ToBytes();
    }
}
