using OfficeIMO.Reader;
using OfficeIMO.Reader.Markdown;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderMarkdownEncodingTests {
    private const string Source = "# Zażółć gęślą jaźń\n\nTreść akapitu.\n";

    public static TheoryData<string> Encodings => new() { "utf-8", "utf-8-bom", "utf-16le-bom", "utf-16be-bom" };

    [Theory]
    [MemberData(nameof(Encodings))]
    public void MarkdownHandler_DecodesTextTheSameFromBytesStreamsAndPaths(string encodingName) {
        byte[] payload = Encode(encodingName);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddMarkdownHandler().Build();
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".md");
        try {
            File.WriteAllBytes(path, payload);
            using var stream = new MemoryStream(payload, writable: false);

            string fromBytes = reader.ConvertToMarkdown(payload, "sample.md");
            string fromStream = reader.ConvertToMarkdown(stream, "sample.md");
            string fromPath = reader.ConvertToMarkdown(path);

            Assert.Contains("Zażółć gęślą jaźń", fromBytes, StringComparison.Ordinal);
            Assert.Contains("Treść akapitu.", fromBytes, StringComparison.Ordinal);
            Assert.DoesNotContain('﻿', fromBytes);
            Assert.Equal(fromBytes, fromStream);
            Assert.Equal(fromBytes, fromPath);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static byte[] Encode(string encodingName) {
        Encoding encoding = encodingName switch {
            "utf-8" => new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            "utf-8-bom" => new UTF8Encoding(encoderShouldEmitUTF8Identifier: true),
            "utf-16le-bom" => new UnicodeEncoding(bigEndian: false, byteOrderMark: true),
            "utf-16be-bom" => new UnicodeEncoding(bigEndian: true, byteOrderMark: true),
            _ => throw new ArgumentOutOfRangeException(nameof(encodingName), encodingName, null)
        };

        byte[] preamble = encoding.GetPreamble();
        byte[] body = encoding.GetBytes(Source);
        byte[] payload = new byte[preamble.Length + body.Length];
        preamble.CopyTo(payload, 0);
        body.CopyTo(payload, preamble.Length);
        return payload;
    }
}
