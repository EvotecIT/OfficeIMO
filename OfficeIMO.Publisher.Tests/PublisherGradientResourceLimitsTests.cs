using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherGradientResourceLimitsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeShapeGradientStopsShareAPreallocationLimit(bool inherited) {
        byte[] bytes = Shapes(inherited);
        var accepted = Read(bytes, 6);
        Assert.Equal(2, accepted.Shapes.Count);
        Assert.All(accepted.Shapes, shape => {
            Assert.Equal(3, shape.Style.FillGradientStops.Count);
            Assert.Equal(new[] { 0D, .5D, 1D }, shape.Style.FillGradientStops.Select(stop => stop.Position));
        });

        var error = Assert.Throws<InvalidDataException>(() => Read(bytes, 5));
        Assert.Contains("gradient stop", error.Message);
    }

    private static PublisherEscherData Read(byte[] bytes, int maximumItems) =>
        new PublisherEscherReader(new PublisherBinaryData(bytes, "Escher/EscherStm"), null,
            new PublisherParseContext(new PublisherReadOptions {
                Limits = new OfficeLegacyImportLimits { MaxItems = maximumItems }
            }, default)).Read();

    private static byte[] Shapes(bool inherited) {
        byte[] gradient = Payload(writer => {
            writer.Write((ushort)0x8197); writer.Write(30U);
            writer.Write((ushort)3); writer.Write((ushort)3); writer.Write((ushort)8);
            for (int index = 0; index < 3; index++) {
                writer.Write((uint)(index * 0x404040)); writer.Write((uint)(index * 32768));
            }
        });
        byte[] properties = Record(0x0013, 0xF00B, gradient);
        var records = new List<byte[]>();
        if (inherited) {
            records.Add(Record(0x000F, 0xF000, properties));
            records.Add(new byte[4]); // Publisher's top-level drawing-instance separator.
        }
        for (uint id = 1; id <= 2; id++) {
            byte[] descriptor = Record(0x0012, 0xF00A, Payload(writer => { writer.Write(id); writer.Write(0U); }));
            byte[] client = Record(0, 0xF011, Payload(writer => { writer.Write(10U); writer.Write((ushort)0x6801); writer.Write(id); }));
            byte[] body = descriptor.Concat(inherited ? Array.Empty<byte>() : properties).Concat(client).ToArray();
            records.Add(Record(0x000F, 0xF004, body));
        }
        return records.SelectMany(record => record).ToArray();
    }

    private static byte[] Record(ushort initial, ushort kind, byte[] body) => Payload(writer => {
        writer.Write(initial); writer.Write(kind); writer.Write(body.Length); writer.Write(body);
    });

    private static byte[] Payload(Action<BinaryWriter> write) {
        using var output = new MemoryStream();
        using (var writer = new BinaryWriter(output, System.Text.Encoding.UTF8, true)) write(writer);
        return output.ToArray();
    }
}
