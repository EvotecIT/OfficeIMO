using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeIMO.Tests;

/// <summary>Small HEIF containers for composition tests. Fake samples are not HEVC bitstreams.</summary>
internal static class HeicTestFixture {
    internal sealed class Item {
        internal Item(string type, int width, int height, byte[] payload, params byte[][] properties) {
            Type = type; Width = width; Height = height; Payload = payload; Properties = properties;
        }
        internal string Type { get; }
        internal int Width { get; }
        internal int Height { get; }
        internal byte[] Payload { get; }
        internal byte[][] Properties { get; }
    }

    internal static Item Tile(int marker, int width = 3, int height = 2, params byte[][] properties) {
        byte[] config = new byte[23]; config[0] = 1; config[16] = 0xFD; config[17] = 0xF8; config[18] = 0xF8;
        return new Item("hvc1", width, height, new[] { (byte)marker }, new[] { Box("hvcC", config) }.Concat(properties).ToArray());
    }

    internal static byte[] Create(Item[] items, params (string Type, int From, int[] To)[] references) {
        var entries = new List<byte[]>(); var locations = new List<byte[]>();
        var properties = new List<byte[]>(); var associations = new List<byte[]>();
        int payloadOffset = 0;
        for (int i = 0; i < items.Length; i++) {
            Item item = items[i];
            entries.Add(Box("infe", Join(U32(0x02000000), U16(i + 1), U16(0), Text(item.Type), new byte[1])));
            locations.Add(Join(U16(i + 1), U16(1), U16(0), U16(1), U32(payloadOffset), U32(item.Payload.Length)));
            payloadOffset += item.Payload.Length;
            var ownProperties = new List<byte[]> { Box("ispe", Join(U32(0), U32(item.Width), U32(item.Height))) };
            ownProperties.AddRange(item.Properties);
            byte[] indexes = Enumerable.Range(properties.Count + 1, ownProperties.Count).Select(x => checked((byte)x)).ToArray();
            associations.Add(Join(U16(i + 1), new[] { checked((byte)indexes.Length) }, indexes));
            properties.AddRange(ownProperties);
        }
        byte[] iref = Box("iref", Join(U32(0), Join(references.Select(r => Box(r.Type, Join(U16(r.From), U16(r.To.Length), Join(r.To.Select(U16).ToArray())))).ToArray())));
        byte[] meta = Box("meta", Join(U32(0), Box("pitm", Join(U32(0), U16(1))),
            Box("iinf", Join(U32(0), U16(items.Length), Join(entries.ToArray()))),
            Box("iloc", Join(U32(0x01000000), new byte[] { 0x44, 0 }, U16(items.Length), Join(locations.ToArray()))),
            Box("iprp", Join(Box("ipco", Join(properties.ToArray())), Box("ipma", Join(U32(0), U32(items.Length), Join(associations.ToArray()))))),
            iref, Box("idat", Join(items.Select(i => i.Payload).ToArray()))));
        return Join(Box("ftyp", Join(Text("heic"), U32(0), Text("mif1heic"))), meta);
    }

    internal static byte[] Grid(int rows, int columns, int width, int height) => Join(new byte[] { 0, 0, (byte)(rows - 1), (byte)(columns - 1) }, U16(width), U16(height));
    internal static byte[] Rotation(int quarters) => Box("irot", new[] { (byte)quarters });
    internal static byte[] Mirror(int axis) => Box("imir", new[] { (byte)axis });
    internal static byte[] Alpha() => Box("auxC", Join(U32(0), Text("urn:mpeg:hevc:2015:auxid:1\0")));
    internal static byte[] Crop(int width, int height, int horizontalNumerator = 0, int horizontalDenominator = 1) =>
        Box("clap", Join(U32(width), U32(1), U32(height), U32(1), U32(horizontalNumerator), U32(horizontalDenominator), U32(0), U32(1)));
    internal static byte[] Box(string type, byte[] payload) => Join(U32(payload.Length + 8), Text(type), payload);
    internal static byte[] U16(int n) => new[] { (byte)(n >> 8), (byte)n };
    internal static byte[] U32(int n) => new[] { (byte)(n >> 24), (byte)(n >> 16), (byte)(n >> 8), (byte)n };
    internal static byte[] Text(string s) => Encoding.ASCII.GetBytes(s);
    internal static byte[] Join(params byte[][] arrays) {
        using var stream = new MemoryStream();
        foreach (byte[] a in arrays) stream.Write(a, 0, a.Length);
        return stream.ToArray();
    }
}
