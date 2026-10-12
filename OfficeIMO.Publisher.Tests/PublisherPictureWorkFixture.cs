using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

// Replaces the first BLIP in an original publication, retaining its picture
// references and all other native streams. These are controlled record inputs.
internal static class PublisherPictureWorkFixture {
    internal static byte[] WithImage(byte[] publication, byte[] image) {
        Assert.True(OfficeCompoundFileReader.TryRead(publication, out OfficeCompoundFile? source, out string? error), error);
        bool replaced = false;
        byte[] Rewrite(byte[] bytes, int start, int end) {
            using var output = new MemoryStream(); using var writer = new BinaryWriter(output);
            for (int offset = start; offset < end;) {
                ushort initial = BitConverter.ToUInt16(bytes, offset), kind = BitConverter.ToUInt16(bytes, offset + 2);
                int content = offset + 8, boundary = content + checked((int)BitConverter.ToUInt32(bytes, offset + 4));
                byte[] body = bytes.Skip(content).Take(boundary - content).ToArray();
                if (kind == 0xF007 && !replaced) {
                    using var data = new MemoryStream(); using var imageWriter = new BinaryWriter(data);
                    byte[] metadata = new byte[36]; metadata[0] = metadata[1] = 6;
                    PublisherInputContractTests.WriteUInt32(metadata, 20, checked((uint)(image.Length + 25)));
                    PublisherInputContractTests.WriteUInt32(metadata, 24, 1);
                    imageWriter.Write(metadata); imageWriter.Write((ushort)0x6E00); imageWriter.Write((ushort)0xF01E);
                    imageWriter.Write(image.Length + 17); imageWriter.Write(new byte[16]); imageWriter.Write((byte)0xFF);
                    imageWriter.Write(image); body = data.ToArray(); initial = 0x62; replaced = true;
                } else if ((initial & 15) == 15) body = Rewrite(bytes, content, boundary);
                writer.Write(initial); writer.Write(kind); writer.Write(body.Length); writer.Write(body);
                offset = boundary;
                if (kind is 0xF000 or 0xF002 && boundary < end) { writer.Write(bytes, boundary, 4); offset += 4; }
            }
            return output.ToArray();
        }
        byte[] escher = source!.Streams["Escher/EscherStm"];
        byte[] replacement = Rewrite(escher, 0, escher.Length);
        Assert.True(replaced);
        return OfficeCompoundFileWriter.Rewrite(source, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = replacement });
    }

    internal static byte[] PaddedPng() {
        byte[] png = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(1, 1, OfficeColor.Red), OfficeImageExportFormat.Png);
        byte[] data = new byte[200_000], type = System.Text.Encoding.ASCII.GetBytes("npAD");
        using var output = new MemoryStream(); using var writer = new BinaryWriter(output);
        writer.Write(png, 0, 33); BigEndian(writer, (uint)data.Length); writer.Write(type); writer.Write(data);
        uint crc = uint.MaxValue;
        foreach (byte value in type.Concat(data)) {
            crc ^= value;
            for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320U : crc >> 1;
        }
        BigEndian(writer, ~crc); writer.Write(png, 33, png.Length - 33); return output.ToArray();
    }

    internal static byte[] TwoFrameGif() {
        byte[] header = { 71, 73, 70, 56, 57, 97, 1, 0, 1, 0, 0x80, 0, 0, 0, 0, 0, 255, 255, 255 };
        byte[] frame = { 0x2C, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 2, 0x4C, 1, 0 };
        return header.Concat(frame).Concat(frame).Concat(new byte[] { 0x3B }).ToArray();
    }

    private static void BigEndian(BinaryWriter writer, uint value) {
        writer.Write((byte)(value >> 24)); writer.Write((byte)(value >> 16)); writer.Write((byte)(value >> 8)); writer.Write((byte)value);
    }
}
