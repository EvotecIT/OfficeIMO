using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstPropertyContextReaderTests {
    [Fact]
    public void DecodesVariableMultiValuedProperties() {
        byte[] unicode = CreateVariableValues(
            Encoding.Unicode.GetBytes("alpha\0"),
            Encoding.Unicode.GetBytes("beta\0"));
        byte[] binary = CreateVariableValues(
            new byte[] { 1, 2, 3 },
            Array.Empty<byte>(),
            new byte[] { 4, 5 });

        Assert.Equal(new[] { "alpha", "beta" },
            Assert.IsType<string[]>(PstPropertyContextReader.DecodeVariable(
                MapiPropertyType.MultipleUnicode, unicode)));
        byte[][] values = Assert.IsType<byte[][]>(PstPropertyContextReader.DecodeVariable(
            MapiPropertyType.MultipleBinary, binary));
        Assert.Equal(3, values.Length);
        Assert.Equal(new byte[] { 1, 2, 3 }, values[0]);
        Assert.Empty(values[1]);
        Assert.Equal(new byte[] { 4, 5 }, values[2]);
    }

    [Fact]
    public void DecodesEveryFixedMultiValuedPropertyFamily() {
        var guid1 = new Guid("8e14d68e-2598-43ec-9bf9-135d878b78d7");
        var guid2 = new Guid("c4e255f8-e880-43fc-b26c-b35540fa07dc");
        var guidBytes = guid1.ToByteArray().Concat(guid2.ToByteArray()).ToArray();
        var doubles = new[] { 1.25, -8.5 };
        byte[] doubleBytes = doubles.SelectMany(BitConverter.GetBytes).ToArray();
        var singles = new[] { 2.5f, -3.75f };
        byte[] singleBytes = singles.SelectMany(BitConverter.GetBytes).ToArray();

        Assert.Equal(singles, Assert.IsType<float[]>(PstPropertyContextReader.DecodeVariable(
            MapiPropertyType.MultipleFloating32, singleBytes)));
        Assert.Equal(doubles, Assert.IsType<double[]>(PstPropertyContextReader.DecodeVariable(
            MapiPropertyType.MultipleFloating64, doubleBytes)));
        Assert.Equal(new[] { guid1, guid2 }, Assert.IsType<Guid[]>(PstPropertyContextReader.DecodeVariable(
            MapiPropertyType.MultipleGuid, guidBytes)));
    }

    [Fact]
    public void RejectsMalformedMultiValuedPropertyLayouts() {
        Assert.Throws<InvalidDataException>(() => PstPropertyContextReader.DecodeVariable(
            MapiPropertyType.MultipleInteger32, new byte[] { 1, 2, 3 }));
        Assert.Throws<InvalidDataException>(() => PstPropertyContextReader.DecodeVariable(
            MapiPropertyType.MultipleUnicode,
            new byte[] { 1, 0, 0, 0, 3, 0, 0, 0 }));
    }

    private static byte[] CreateVariableValues(params byte[][] values) {
        int headerLength = checked(4 + values.Length * 4);
        var result = new byte[checked(headerLength + values.Sum(value => value.Length))];
        WriteUInt32(result, 0, (uint)values.Length);
        int cursor = headerLength;
        for (int index = 0; index < values.Length; index++) {
            WriteUInt32(result, 4 + index * 4, (uint)cursor);
            Buffer.BlockCopy(values[index], 0, result, cursor, values[index].Length);
            cursor += values[index].Length;
        }
        return result;
    }

    private static void WriteUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }
}
