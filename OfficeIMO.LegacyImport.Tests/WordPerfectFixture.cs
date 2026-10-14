using System;
using System.Linq;
using System.Text;

namespace OfficeIMO.LegacyImport.Tests;

/// <summary>Small specification-built records; independent files are kept separately.</summary>
internal static class WordPerfectFixture {
    internal static byte[] Text(string text) => Encoding.ASCII.GetBytes(text);
    internal static byte[] Join(params byte[][] parts) => parts.SelectMany(part => part).ToArray();
    internal static byte[] Number(int value) => new[] { (byte)value, (byte)(value >> 8) };

    internal static byte[] Function6(byte code, byte sub, byte[]? data = null, params int[] packets) {
        data ??= Array.Empty<byte>();
        int size = 10 + data.Length + (packets.Length > 0 ? 1 + packets.Length * 2 : 0);
        return Join(new[] { code, sub }, Number(size), new[] { packets.Length > 0 ? (byte)0x80 : (byte)0 },
            packets.Length > 0 ? Join(new[] { (byte)packets.Length }, packets.SelectMany(Number).ToArray()) : Array.Empty<byte>(),
            Number(data.Length), data, Number(size), new[] { code });
    }

    internal static byte[] Function5(byte code, byte sub, byte[]? data = null) {
        data ??= Array.Empty<byte>();
        return Join(new[] { code, sub }, Number(4 + data.Length), data, Number(4 + data.Length), new[] { sub, code });
    }

    internal static byte[] TextPacket(byte[] text) => Join(Number(1), BitConverter.GetBytes(10), BitConverter.GetBytes(text.Length), text);

    internal static byte[] Document6(byte[] body, params (byte Type, byte[] Data)[] packets) {
        int count = packets.Length + 1, offset = 512 + count * 14;
        var bytes = new byte[offset + packets.Sum(packet => packet.Data.Length) + body.Length];
        Header(bytes, 2); Array.Copy(Number(512), 0, bytes, 14, 2);
        bytes[512] = 2; Array.Copy(Number(count), 0, bytes, 514, 2);
        for (int i = 0; i < packets.Length; i++) {
            int entry = 512 + (i + 1) * 14;
            bytes[entry + 1] = packets[i].Type;
            Array.Copy(Number(1), 0, bytes, entry + 2, 2);
            Array.Copy(BitConverter.GetBytes(packets[i].Data.Length), 0, bytes, entry + 6, 4);
            Array.Copy(BitConverter.GetBytes(offset), 0, bytes, entry + 10, 4);
            Array.Copy(packets[i].Data, 0, bytes, offset, packets[i].Data.Length);
            offset += packets[i].Data.Length;
        }
        Array.Copy(BitConverter.GetBytes(offset), 0, bytes, 4, 4);
        Array.Copy(BitConverter.GetBytes(bytes.Length), 0, bytes, 20, 4);
        Array.Copy(body, 0, bytes, offset, body.Length);
        return bytes;
    }

    internal static byte[] Document5(byte[] body) {
        var bytes = new byte[66 + body.Length]; Header(bytes, 0);
        Array.Copy(BitConverter.GetBytes(66), 0, bytes, 4, 4);
        Array.Copy(Number(0xfffb), 0, bytes, 16, 2);
        Array.Copy(Number(5), 0, bytes, 18, 2); Array.Copy(Number(50), 0, bytes, 20, 2);
        Array.Copy(body, 0, bytes, 66, body.Length);
        return bytes;
    }

    private static void Header(byte[] bytes, byte version) {
        bytes[0] = 0xff; bytes[1] = 0x57; bytes[2] = 0x50; bytes[3] = 0x43;
        bytes[8] = 1; bytes[9] = 10; bytes[10] = version; bytes[11] = 1;
    }
}
