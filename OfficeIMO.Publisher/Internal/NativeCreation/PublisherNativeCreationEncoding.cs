using System.Text;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Assembles fresh creation buffers and the OfficeArt/client envelopes specific to this native profile.</summary>
internal static class PublisherNativeCreationEncoding {
    internal static byte[] Bytes(Action<BinaryWriter> action) {
        using var buffer = new MemoryStream();
        using var writer = new BinaryWriter(buffer, Encoding.UTF8, leaveOpen: true);
        action(writer);
        return buffer.ToArray();
    }

    internal static byte[] Join(params byte[][] parts) => Bytes(writer => {
        foreach (byte[] part in parts) writer.Write(part);
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