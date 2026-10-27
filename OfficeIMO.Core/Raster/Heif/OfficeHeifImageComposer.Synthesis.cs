using System;
using System.IO;
using System.Text;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeHeifImageComposer {
    /// <summary>Wraps one unchanged coded item in a container with no presentation or color properties.</summary>
    private byte[] Synthesize(OfficeHeifItemInfo item, int width, int height) {
        byte[] payload = ReadPayload(item);
        byte[] config = new byte[item.CodecConfigurationBytes.Count];
        for (int i = 0; i < config.Length; i++) config[i] = item.CodecConfigurationBytes[i];
        byte[] ftyp = Box("ftyp", Join(Ascii("heic"), Number(0), Ascii("mif1heic")));
        byte[] meta = MakeMeta(config, width, height, 0, payload.Length);
        int offset = checked(ftyp.Length + meta.Length + 8);
        meta = MakeMeta(config, width, height, offset, payload.Length);
        byte[] result = Join(ftyp, meta, Box("mdat", payload));
        Require(result.Length <= _options.MaximumEncodedBytes, "Synthesized HEIC item exceeds the encoded-byte limit.");
        return result;
    }

    private static byte[] MakeMeta(byte[] config, int width, int height, int offset, int length) {
        byte[] hdlr = Box("hdlr", Join(Number(0), Number(0), Ascii("pict"), new byte[13]));
        byte[] pitm = Box("pitm", Join(Number(0), Short(1)));
        byte[] infe = Box("infe", Join(Number(0x02000000), Short(1), Short(0), Ascii("hvc1"), new byte[1]));
        byte[] iinf = Box("iinf", Join(Number(0), Short(1), infe));
        byte[] iloc = Box("iloc", Join(Number(0), new byte[] { 0x44, 0 }, Short(1), Short(1), Short(0), Short(1), Number(offset), Number(length)));
        byte[] ispe = Box("ispe", Join(Number(0), Number(width), Number(height)));
        byte channels = (config[16] & 3) == 0 ? (byte)1 : (byte)3;
        byte depth = (byte)(8 + (config[17] & 7));
        byte[] depths = channels == 1 ? new[] { channels, depth } : new[] { channels, depth, depth, depth };
        byte[] ipco = Box("ipco", Join(Box("hvcC", config), ispe, Box("pixi", Join(Number(0), depths))));
        byte[] ipma = Box("ipma", Join(Number(0), Number(1), Short(1), new byte[] { 3, 0x81, 2, 3 }));
        return Box("meta", Join(Number(0), hdlr, pitm, iinf, iloc, Box("iprp", Join(ipco, ipma))));
    }

    private static byte[] Box(string type, byte[] payload) => Join(Number(checked(payload.Length + 8)), Ascii(type), payload);
    private static byte[] Ascii(string text) => Encoding.ASCII.GetBytes(text);
    private static byte[] Short(int value) => new[] { (byte)(value >> 8), (byte)value };
    private static byte[] Number(int value) => new[] { (byte)(value >> 24), (byte)(value >> 16), (byte)(value >> 8), (byte)value };
    private static byte[] Join(params byte[][] arrays) {
        int length = 0;
        foreach (byte[] array in arrays) length = checked(length + array.Length);
        var result = new byte[length];
        int offset = 0;
        foreach (byte[] array in arrays) { Buffer.BlockCopy(array, 0, result, offset, array.Length); offset += array.Length; }
        return result;
    }
}
