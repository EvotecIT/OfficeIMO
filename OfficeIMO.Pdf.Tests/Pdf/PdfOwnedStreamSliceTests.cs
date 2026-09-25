using System.Collections.Concurrent;
using System.IO.Compression;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfOwnedStreamSliceTests {
    [Fact]
    public void FlateDecodeConsumesOwnedSourceSliceWithoutMaterializingIt() {
        byte[] expected = Enumerable.Repeat((byte)'x', 4096).ToArray();
        byte[] compressed = Compress(expected);
        const int prefixLength = 17;
        var source = new byte[prefixLength + compressed.Length + 23];
        Buffer.BlockCopy(compressed, 0, source, prefixLength, compressed.Length);
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        PdfStream stream = PdfStream.FromOwnedSource(dictionary, source, prefixLength, compressed.Length);

        int validatedLength = StreamDecoder.ValidateRequired(stream, maxOutputBytes: expected.Length);
        byte[] decoded = StreamDecoder.DecodeRequired(stream, maxOutputBytes: expected.Length);
        stream.GetDataSegment(out byte[] retained, out int offset, out int length);

        Assert.Equal(expected.Length, validatedLength);
        Assert.Equal(expected, decoded);
        Assert.Same(source, retained);
        Assert.Equal(prefixLength, offset);
        Assert.Equal(compressed.Length, length);
    }

    [Fact]
    public void ValidationOnlyDecodePreservesLimitsAndMalformedDataFailures() {
        byte[] expanded = Enumerable.Repeat((byte)'y', 1024).ToArray();
        byte[] compressed = Compress(expanded);
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        var stream = new PdfStream(dictionary, compressed);
        var malformed = new PdfStream(dictionary, Encoding.ASCII.GetBytes("not-valid-flate-data"));

        Assert.Throws<PdfReadLimitException>(() =>
            StreamDecoder.ValidateRequired(stream, maxOutputBytes: expanded.Length - 1));
        Assert.Throws<InvalidDataException>(() =>
            StreamDecoder.ValidateRequired(malformed, maxOutputBytes: expanded.Length));
    }

    [Fact]
    public void ExplicitDecodeParametersUseTheFullCanonicalDecoder() {
        byte[] expected = { 1, 2, 3, 4 };
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        var decodeParameters = new PdfDictionary();
        decodeParameters.Items["Predictor"] = new PdfNumber(1);
        dictionary.Items["DecodeParms"] = decodeParameters;
        var stream = new PdfStream(dictionary, Compress(expected));

        int validatedLength = StreamDecoder.ValidateRequired(stream, maxOutputBytes: expected.Length);

        Assert.Equal(expected.Length, validatedLength);
        Assert.Equal(expected, StreamDecoder.DecodeRequired(stream, maxOutputBytes: expected.Length));
    }

    [Fact]
    public void UnfilteredOwnedSourceSliceReturnsOnlyItsPayload() {
        byte[] source = { 91, 92, 1, 2, 3, 93 };
        PdfStream stream = PdfStream.FromOwnedSource(new PdfDictionary(), source, 2, 3);

        byte[] decoded = StreamDecoder.DecodeRequired(stream, maxOutputBytes: 3);

        Assert.Equal(new byte[] { 1, 2, 3 }, decoded);
        Assert.NotSame(source, decoded);
    }

    [Fact]
    public void LenientDecodeChecksSourceLimitBeforeMaterializingOwnedSlice() {
        byte[] source = { 91, 92, 1, 2, 3, 4, 93 };
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("UnsupportedFilter");
        PdfStream stream = PdfStream.FromOwnedSource(dictionary, source, 2, 4);

        Assert.Throws<PdfReadLimitException>(() => StreamDecoder.Decode(stream, maxOutputBytes: 3));
        stream.GetDataSegment(out byte[] retained, out int offset, out int length);
        Assert.Same(source, retained);
        Assert.Equal(2, offset);
        Assert.Equal(4, length);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, StreamDecoder.Decode(stream, maxOutputBytes: 4));
    }

    [Fact]
    public void ConcurrentMaterializationPublishesOneStablePayload() {
        byte[] source = Enumerable.Range(0, 1024).Select(static value => (byte)value).ToArray();
        PdfStream stream = PdfStream.FromOwnedSource(new PdfDictionary(), source, 100, 700);
        var observed = new ConcurrentBag<byte[]>();

        Parallel.For(0, 64, _ => observed.Add(stream.Data));

        byte[] published = Assert.Single(observed.Distinct());
        Assert.Equal(source.Skip(100).Take(700), published);
        Assert.Same(published, stream.Data);
    }

    [Fact]
    public void LoadedDocumentRetainsItsCallerSnapshotAfterInputMutation() {
        byte[] callerBytes = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Caller snapshot"))
            .ToBytes();
        byte[] expected = (byte[])callerBytes.Clone();
        PdfDocument document = PdfDocument.Load(callerBytes);

        Array.Clear(callerBytes, 0, callerBytes.Length);

        Assert.Equal(expected, document.ToBytes());
        Assert.Contains("Caller snapshot", document.Reader.Text(), StringComparison.Ordinal);
    }

    private static byte[] Compress(byte[] bytes) {
        using var output = new MemoryStream();
        using (var compressor = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            compressor.Write(bytes, 0, bytes.Length);
        }
        return output.ToArray();
    }
}
