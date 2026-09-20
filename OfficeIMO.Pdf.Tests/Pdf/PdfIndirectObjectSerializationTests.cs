using System.Collections.Generic;
using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfIndirectObjectSerializationTests {
    [Fact]
    public void ObjectSerializationEscapesDictionaryKeysAndNamesWithoutChangingBytes() {
        var context = new PdfPageExtractor.SerializationContext(
            new Dictionary<int, int>(),
            pagesObjectId: 0,
            new Dictionary<int, Dictionary<string, PdfObject>>());
        var dictionary = new PdfDictionary();
        dictionary.Items["Ordinary"] = new PdfName("Value");
        dictionary.Items["Space Key"] = new PdfName("Hash#Value");
        dictionary.Items["Unicode\u0141"] = new PdfName("Slash/Value");
        dictionary.Items["Emoji\U0001F600"] = new PdfName("Caf\u00E9");
        dictionary.Items["Integer"] = new PdfNumber(-42);
        dictionary.Items["Real"] = new PdfNumber(1.23456);
        dictionary.Items["Reference"] = new PdfReference(7, 0);
        context.NumberMap[7] = 42;

        byte[] serialized = PdfPageExtractor.SerializeObject(dictionary, context);

        Assert.Equal(
            "<< /Ordinary /Value /Space#20Key /Hash#23Value /Unicode#C5#81 /Slash#2FValue /Emoji#F0#9F#98#80 /Caf#C3#A9 /Integer -42 /Real 1.235 /Reference 42 0 R >>\n",
            PdfEncoding.Latin1GetString(serialized));
        Assert.Equal("Unicode\u0141", PdfSyntax.DecodeName("Unicode#C5#81"));
        Assert.Equal("Emoji\U0001F600", PdfSyntax.DecodeName("Emoji#F0#9F#98#80"));
        Assert.Equal("Latin\u00E9", PdfSyntax.DecodeName("Latin#E9"));
        PdfPageExtractor.EnsureSerializedObjectWithinLimit(dictionary, context, serialized.LongLength);
        Assert.Throws<InvalidDataException>(() =>
            PdfPageExtractor.EnsureSerializedObjectWithinLimit(dictionary, context, serialized.LongLength - 1L));
    }

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
