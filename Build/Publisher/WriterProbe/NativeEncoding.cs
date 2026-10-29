using System.Text;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Small binary-envelope builders for an unqualified creation experiment.</summary>
internal static class NativeEncoding {
    internal static byte[] Bytes(Action<BinaryWriter> action) {
        using var buffer = new MemoryStream();
        using var writer = new BinaryWriter(buffer, Encoding.UTF8, leaveOpen: true);
        action(writer);
        return buffer.ToArray();
    }

    internal static byte[] Join(params byte[][] parts) => Bytes(writer => {
        foreach (byte[] part in parts) writer.Write(part);
    });

    internal static byte[] Scalar(byte id, byte type, uint value) => Bytes(writer => {
        writer.Write(id); writer.Write(type); writer.Write(value);
    });

    internal static byte[] ShortScalar(byte id, byte type, ushort value) => Bytes(writer => {
        writer.Write(id); writer.Write(type); writer.Write(value);
    });

    internal static byte[] Marker(byte id, byte type) => new[] { id, type };

    internal static byte[] Variable(byte id, byte type, params byte[][] fields) => Bytes(writer => {
        byte[] body = Join(fields);
        writer.Write(id); writer.Write(type); writer.Write(body.Length + 4); writer.Write(body);
    });

    internal static byte[] Chunk(params byte[][] fields) => Bytes(writer => {
        byte[] body = Join(fields);
        writer.Write(body.Length + 4); writer.Write(body);
    });

    internal static byte[] Art(ushort initial, ushort type, params byte[][] fields) => Bytes(writer => {
        byte[] body = Join(fields);
        writer.Write(initial); writer.Write(type); writer.Write(body.Length); writer.Write(body);
    });

    internal static byte[] Client(params (ushort Id, uint Value)[] properties) => Bytes(writer => {
        writer.Write(4 + properties.Length * 6);
        foreach (var property in properties) { writer.Write(property.Id); writer.Write(property.Value); }
    });
}
