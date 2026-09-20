using System.Collections.Generic;
using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfIndirectObjectSerializationTests {
    [Fact]
    public void DirectStreamSerializationPreservesIndirectObjectBytes() {
        var sourceObjects = new Dictionary<int, PdfIndirectObject> {
            [7] = new PdfIndirectObject(7, 1, new PdfName("Example"))
        };
        var numberMap = new Dictionary<int, int> { [7] = 2 };
        var context = new PdfPageExtractor.SerializationContext(
            numberMap,
            pagesObjectId: 0,
            new Dictionary<int, Dictionary<string, PdfObject>>(),
            sourceObjects);
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Length"] = new PdfNumber(1);
        dictionary.Items["Related"] = new PdfReference(7, 1);
        var stream = new PdfStream(dictionary, new byte[] { 0, 10, 13, 37, 128, 255 });

        byte[] expected = PdfPageExtractor.WrapObject(4, PdfPageExtractor.SerializeObject(stream, context));
        byte[] actual = PdfPageExtractor.SerializeIndirectObject(4, stream, context);

        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(81937)]
    public void SegmentedStreamAssemblyMatchesBufferedAssemblyExactly(int payloadLength) {
        var numberMap = new Dictionary<int, int>();
        var context = new PdfPageExtractor.SerializationContext(
            numberMap,
            pagesObjectId: 0,
            new Dictionary<int, Dictionary<string, PdfObject>>());
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("Catalog");
        dictionary.Items["Length"] = new PdfNumber(1);
        var ownedSource = new byte[payloadLength + 2];
        for (int index = 0; index < payloadLength; index++) {
            ownedSource[index + 1] = (byte)(index * 31);
        }
        PdfStream stream = PdfStream.FromOwnedSource(dictionary, ownedSource, 1, payloadLength);

        byte[] bufferedObject = PdfPageExtractor.SerializeIndirectObject(1, stream, context);
        PdfSerializedObject segmentedObject = PdfPageExtractor.SerializeIndirectObjectForAssembly(1, stream, context);
        byte[] secondObject = PdfPageExtractor.WrapObject(2, PdfEncoding.Latin1GetBytes("<< /Type /Example >>\n"));

        byte[] bufferedFile = PdfFileAssembler.Assemble(new[] { bufferedObject, secondObject }, 1, 0);
        using var cancellation = new CancellationTokenSource();
        byte[] segmentedFile = PdfFileAssembler.Assemble(
            new[] { segmentedObject, PdfSerializedObject.FromBytes(secondObject) },
            1,
            0,
            cancellationToken: cancellation.Token);

        Assert.Equal(bufferedObject.LongLength, segmentedObject.Length);
        Assert.Equal(bufferedFile.LongLength, PdfFileAssembler.GetAssembledLength(
            new[] { segmentedObject, PdfSerializedObject.FromBytes(secondObject) },
            1,
            0));
        Assert.Equal(bufferedFile, segmentedFile);
    }
}
