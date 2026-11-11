using System.Text;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// The suite and corpus are UTF-8; these cover the encodings and byte-order marks a save must reproduce, through every
// load and save entry point.
public class YamlDocumentEncodingTests {
    private const string Text = "# naïve 🙂\r\nkey: \"wert\"\nlist: [1, 2]\n";

    public static TheoryData<string, bool> Encodings() => new() {
        { "utf-8", false },
        { "utf-8", true },
        { "utf-16LE", false },
        { "utf-16LE", true },
        { "utf-16BE", false },
        { "utf-16BE", true },
        { "utf-32LE", false },
        { "utf-32LE", true },
        { "utf-32BE", false },
        { "utf-32BE", true }
    };

    private static byte[] Encode(string encodingName, bool byteOrderMark) {
        Encoding encoding = encodingName switch {
            "utf-16LE" => new UnicodeEncoding(false, false),
            "utf-16BE" => new UnicodeEncoding(true, false),
            "utf-32LE" => new UTF32Encoding(false, false),
            "utf-32BE" => new UTF32Encoding(true, false),
            _ => new UTF8Encoding(false)
        };
        byte[] body = encoding.GetBytes(Text);
        if (!byteOrderMark) return body;
        byte[] preamble = encodingName switch {
            "utf-16LE" => new byte[] { 0xFF, 0xFE },
            "utf-16BE" => new byte[] { 0xFE, 0xFF },
            "utf-32LE" => new byte[] { 0xFF, 0xFE, 0, 0 },
            "utf-32BE" => new byte[] { 0, 0, 0xFE, 0xFF },
            _ => new byte[] { 0xEF, 0xBB, 0xBF }
        };
        return preamble.Concat(body).ToArray();
    }

    [Theory]
    [MemberData(nameof(Encodings))]
    public void StreamSaveReproducesEncodingAndByteOrderMark(string encodingName, bool byteOrderMark) {
        byte[] input = Encode(encodingName, byteOrderMark);
        YamlDocument document = YamlDocument.Load(new MemoryStream(input));
        Assert.Equal("wert", document.Select("key")!.AsString());
        Assert.Equal(Text, document.ToString());
        var output = new MemoryStream();
        document.Save(output);
        Assert.Equal(input, output.ToArray());
    }

    [Theory]
    [MemberData(nameof(Encodings))]
    public async Task FileAndAsyncSavesReproduceTheInput(string encodingName, bool byteOrderMark) {
        byte[] input = Encode(encodingName, byteOrderMark);
        string source = Path.Combine(Path.GetTempPath(), "officeimo-yaml-" + Guid.NewGuid().ToString("N") + ".yaml");
        string target = source + ".saved";
        try {
            File.WriteAllBytes(source, input);
            YamlDocument.Load(source).Save(target);
            Assert.Equal(input, File.ReadAllBytes(target));
            await (await YamlDocument.LoadAsync(source)).SaveAsync(target);
            Assert.Equal(input, File.ReadAllBytes(target));
        } finally {
            File.Delete(source);
            File.Delete(target);
        }
    }

    // FilePath follows the shared document lifecycle: Create(path) and loads set it, saves to a path move it, and
    // Save() writes to it. Text and streams have none.
    [Fact]
    public void FilePathComesFromCreateLoadAndPathSaves() {
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-yaml-" + Guid.NewGuid().ToString("N"));
        string created = Path.Combine(folder, "created.yaml");
        string copy = Path.Combine(folder, "copy.yaml");
        try {
            Assert.Null(YamlDocument.Parse("a: 1\n").FilePath);
            Assert.Throws<InvalidOperationException>(() => YamlDocument.Parse("a: 1\n").Save());

            YamlDocument document = YamlDocument.Create(created);
            Assert.Equal(created, document.FilePath);
            Assert.False(File.Exists(created));
            document.Set("a", 1);
            document.Save();
            Assert.Equal("a: 1\n", File.ReadAllText(created));

            YamlDocument loaded = YamlDocument.Load(created);
            Assert.Equal(created, loaded.FilePath);
            loaded.Save(copy);
            Assert.Equal(copy, loaded.FilePath);
            loaded.Set("a", 2);
            loaded.Save();
            Assert.Equal("a: 2\n", File.ReadAllText(copy));
            Assert.Equal("a: 1\n", File.ReadAllText(created));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void ParsedTextSavesAsUtf8AndKeepsALeadingByteOrderMarkCharacter() {
        const string text = "﻿key: v\n";
        YamlDocument document = YamlDocument.Parse(text);
        Assert.Equal(text, document.ToString());
        var output = new MemoryStream();
        document.Save(output);
        Assert.Equal(Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes("key: v\n")).ToArray(), output.ToArray());
    }

    [Fact]
    public void InvalidBytesFailWithTheEncodingCode() {
        YamlException error = Assert.Throws<YamlException>(() => YamlDocument.Load(new MemoryStream(new byte[] { (byte)'a', (byte)':', (byte)' ', 0xC3, 0x28 })));
        Assert.Equal(YamlErrorCodes.InvalidEncoding, error.Code);
    }
}
