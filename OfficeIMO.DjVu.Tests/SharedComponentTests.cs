namespace OfficeIMO.DjVu.Tests;

public sealed class SharedComponentTests {
    private static string Folder => Path.Combine(AppContext.BaseDirectory, "Fixtures", "Authored", "Shared");

    [Fact]
    public void BundledAndExplicitIndirectDictionariesMatchIndependentPixels() {
        var bundled = DjVuDocument.Load(Path.Combine(Folder, "shared.djvu"));
        byte[] index = File.ReadAllBytes(Path.Combine(Folder, "Indirect", "index.djvu"));
        Assert.Throws<NotSupportedException>(() => DjVuDocument.Load(index));
        var components = Components();
        var indirect = DjVuDocument.Load(index, new DjVuReadOptions { ComponentResolver = (id, _) => components[id] });
        Assert.Equal(index.Length, indirect.SourceLengthBytes);
        foreach (byte[] bytes in components.Values) Array.Clear(bytes, 0, bytes.Length);
        Assert.Equal(2, indirect.Pages.Count);
        for (int i = 0; i < 2; i++) {
            byte[] expected = PageRenderTests.ReadPpm(Path.Combine(Folder, $"page-{i + 1}-reference.ppm")).GetPixels();
            Assert.Equal(expected, bundled.Pages[i].Render().Image.GetPixels());
            Assert.Equal(expected, indirect.Pages[i].Render().Image.GetPixels());
        }
    }

    [Fact]
    public void DuplicateIncludesAreIdempotentButCyclesAndForeignFormsAreRejected() {
        byte[] index = File.ReadAllBytes(Path.Combine(Folder, "Indirect", "index.djvu"));
        var components = Components();
        string dictionary = components.Keys.Single(k => k.EndsWith(".iff", StringComparison.Ordinal));
        components["shared-1.djvu"] = AppendChunk(components["shared-1.djvu"], "INCL", Encoding.UTF8.GetBytes(dictionary));
        var duplicate = DjVuDocument.Load(index, new DjVuReadOptions { ComponentResolver = (id, _) => components[id] });
        Assert.Equal(PageRenderTests.ReadPpm(Path.Combine(Folder, "page-1-reference.ppm")).GetPixels(), duplicate.Pages[0].Render().Image.GetPixels());
        components[dictionary] = AppendChunk(components[dictionary], "INCL", Encoding.UTF8.GetBytes(dictionary));
        Assert.Throws<InvalidDataException>(() => DjVuDocument.Load(index, new DjVuReadOptions { ComponentResolver = (id, _) => components[id] }));
        components = Components();
        components["shared-1.djvu"] = components[dictionary];
        Assert.Throws<InvalidDataException>(() => DjVuDocument.Load(index, new DjVuReadOptions { ComponentResolver = (id, _) => components[id] }));
    }

    [Fact]
    public void ResolvedBytesAndDictionaryCodecMemoryStayBounded() {
        byte[] index = File.ReadAllBytes(Path.Combine(Folder, "Indirect", "index.djvu"));
        var components = Components();
        Assert.Throws<DjVuResourceLimitException>(() => DjVuDocument.Load(index, new DjVuReadOptions {
            MaxSourceBytes = index.Length, ComponentResolver = (id, _) => components[id]
        }));
        var document = DjVuDocument.Load(index, new DjVuReadOptions { ComponentResolver = (id, _) => components[id] });
        Assert.Throws<DjVuResourceLimitException>(() => document.Pages[0].Render(new DjVuRenderOptions {
            Region = new DjVuRectangle(0, 0, 1, 1), MaxBytes = 512
        }));
    }

    internal static byte[] AppendChunk(byte[] source, string id, byte[] payload) {
        int form = Encoding.ASCII.GetString(source, 0, 4) == "AT&T" ? 4 : 0;
        int size = (source[form + 4] << 24) | (source[form + 5] << 16) | (source[form + 6] << 8) | source[form + 7];
        int start = form + 8 + size;
        var result = new byte[start + (start & 1) + 8 + payload.Length + (payload.Length & 1)];
        Buffer.BlockCopy(source, 0, result, 0, start);
        int at = start + (start & 1);
        Encoding.ASCII.GetBytes(id).CopyTo(result, at);
        WriteSize(result, at + 4, payload.Length);
        payload.CopyTo(result, at + 8);
        WriteSize(result, form + 4, result.Length - form - 8);
        return result;
    }

    private static void WriteSize(byte[] bytes, int offset, int value) {
        for (int i = 0; i < 4; i++) bytes[offset + i] = (byte)(value >> (24 - i * 8));
    }

    private static System.Collections.Generic.Dictionary<string, byte[]> Components() => Directory.GetFiles(Path.Combine(Folder, "Indirect"))
        .Where(p => Path.GetFileName(p) != "index.djvu").ToDictionary(p => Path.GetFileName(p), File.ReadAllBytes, StringComparer.Ordinal);
}
