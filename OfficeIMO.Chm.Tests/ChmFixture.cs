using System.Text;

namespace OfficeIMO.Chm.Tests;

internal static class ChmFixture {
    internal static string NativePath => Path.Combine(AppContext.BaseDirectory, "Fixtures", "svnbook-ms-htmlhelp.chm");
    internal static byte[] Html(string value) => Encoding.UTF8.GetBytes(value);
    internal static byte[] Book(uint version = 3) => Archive(new Dictionary<string, byte[]> {
        ["/#SYSTEM"] = SystemMetadata(), ["/#TOCIDX"] = Array.Empty<byte>(),
        ["/guide/contents.hhc"] = Html("<ul><li><object type='text/sitemap'><param name='Name' value='Guide'></object><ul><li><object type='text/sitemap'><param name='Name' value='Welcome'><param name='Local' value='../welcome.html'></object><li><object type='text/sitemap'><param name='Name' value='Details'><param name='Local' value='details.html#part'></object></ul></ul>"),
        ["/guide/index.hhk"] = Html("<ul><li><object type='text/sitemap'><param name='Name' value='Example'><param name='Local' value='../welcome.html'><param name='Name' value='Details'><param name='Local' value='details.html#part'></object><li><object type='text/sitemap'><param name='Name' value='Alias'><param name='See Also' value='Example'></object></ul>"),
        ["/welcome.html"] = Html("<!doctype html><html><head><meta charset='utf-8'><title>Welcome</title></head><body><h1 id='intro'>Welcome</h1><p>Café and help.</p><a href='guide/details.html#part'>Read details</a><img src='pixel.png' alt='White pixel'></body></html>"),
        ["/guide/details.html"] = Html("<html><head><meta charset='utf-8'></head><body><h1 id='part'>Details</h1><p>Second topic.</p><table><tr><th>Name</th><th>Value</th></tr><tr><td>Answer</td><td>42</td></tr></table><a href='../welcome.html#intro'>Return</a></body></html>"),
        ["/pixel.png"] = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=")
    }, version);

    internal static byte[] SystemMetadata(uint locale = 1033) {
        using var data = new MemoryStream(); using var writer = new BinaryWriter(data);
        writer.Write(3U);
        void Text(ushort code, string value) { byte[] bytes = Encoding.ASCII.GetBytes(value + '\0'); writer.Write(code); writer.Write((ushort)bytes.Length); writer.Write(bytes); }
        Text(0, "guide/contents.hhc"); Text(1, "guide/index.hhk"); Text(2, "welcome.html"); Text(3, "OfficeIMO help fixture"); Text(9, "OfficeIMO test-only uncompressed fixture builder");
        writer.Write((ushort)4); writer.Write((ushort)4); writer.Write(locale);
        return data.ToArray();
    }

    // Only builds section-zero ITSF fixtures. Production LZX acceptance uses an independently compiled archive.
    internal static byte[] Archive(IEnumerable<KeyValuePair<string, byte[]>> files, uint version = 3, int section = 0) {
        var entries = files.ToArray();
        using var records = new MemoryStream(); using var payload = new MemoryStream();
        foreach (var pair in entries) {
            byte[] name = Encoding.UTF8.GetBytes(pair.Key); EncInt(records, name.Length); records.Write(name, 0, name.Length);
            EncInt(records, section); EncInt(records, checked((int)payload.Length)); EncInt(records, pair.Value.Length);
            payload.Write(pair.Value, 0, pair.Value.Length);
        }
        int blockSize = 4096; while (blockSize < records.Length + 22) blockSize *= 2;
        int header = version == 2 ? 88 : 96, directoryLength = 84 + blockSize, content = header + directoryLength;
        var result = new byte[content + payload.Length];
        using var stream = new MemoryStream(result); using var writer = new BinaryWriter(stream);
        void At(int position, uint value) { stream.Position = position; writer.Write(value); }
        void LongAt(int position, ulong value) { stream.Position = position; writer.Write(value); }
        void Signature(int position, string value) { stream.Position = position; writer.Write(Encoding.ASCII.GetBytes(value)); }
        Signature(0, "ITSF"); At(4, version); At(8, (uint)header); At(20, 1033);
        LongAt(72, (ulong)header); LongAt(80, (ulong)directoryLength); if (version == 3) LongAt(88, (ulong)content);
        Signature(header, "ITSP"); At(header + 4, 1); At(header + 8, 84); At(header + 16, (uint)blockSize); At(header + 44, 1);
        int block = header + 84; Signature(block, "PMGL"); At(block + 4, (uint)(blockSize - 20 - records.Length));
        stream.Position = block + 20; writer.Write(records.ToArray()); stream.Position = block + blockSize - 2; writer.Write((ushort)entries.Length);
        stream.Position = content; writer.Write(payload.ToArray()); return result;
    }
    private static void EncInt(Stream stream, int value) {
        var bytes = new byte[5]; int count = 0;
        do { bytes[count++] = (byte)(value & 127); value >>= 7; } while (value != 0);
        while (count != 0) { count--; stream.WriteByte((byte)(bytes[count] | (count == 0 ? 0 : 128))); }
    }
}
