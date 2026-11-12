using System.Text;
using OfficeIMO.Yaml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// Stream and TextReader input go through the pooled sliding buffer, which the suite (small files read in one refill)
// never exercises. These tests compare buffered reading against in-memory reading of the same text.
public class YamlReaderInputTests {
    private static readonly Lazy<YamlTestSuite> Suite = new(() => YamlTestSuite.Load());

    // Every suite case read one character per refill: each token straddles refills at every possible offset.
    [Fact]
    public void ReadingOneCharacterPerRefillMatchesInMemoryReading() {
        foreach (YamlTestSuiteCase testCase in Suite.Value.Cases.Where(c => !c.IsError)) {
            string text = YamlTextTestHelpers.DecodeSuiteInput(testCase.Input);
            var options = new YamlReaderOptions { IncludeComments = true };
            string expected = YamlTextTestHelpers.Events(new YamlReader(text, options));
            string actual = YamlTextTestHelpers.Events(new YamlReader(new TrickleReader(text, 1), options));
            Assert.True(expected == actual, testCase.Id + " differs when read one character at a time.");
        }
    }

    // A long stream forces the window to compact between events and to grow while one scalar is longer than the buffer.
    [Fact]
    public void LargeStreamCompactsAndGrowsTheWindowWithoutChangingEvents() {
        var text = new StringBuilder();
        for (int i = 0; i < 400; i++) {
            text.Append("---\n# document ").Append(i).Append('\n');
            text.Append("name: item-").Append(i).Append("  # trailing\n");
            text.Append("quoted: \"line one\n  line \\u00e9 two\"\n");
            text.Append("tags: [a, b, {k: v}]\n");
            text.Append("text: |\n");
            int lines = i % 50 == 0 ? 2000 : 3;
            for (int line = 0; line < lines; line++) text.Append("  content line ").Append(line).Append('\n');
        }
        text.Append("long: ").Append('x', 70_000).Append('\n');
        string yaml = text.ToString();
        var options = new YamlReaderOptions { IncludeComments = true };
        string expected = YamlTextTestHelpers.Events(new YamlReader(yaml, options));
        Assert.Equal(expected, YamlTextTestHelpers.Events(new YamlReader(new TrickleReader(yaml, 4093), options)));
        Assert.Equal(expected, YamlTextTestHelpers.Events(new YamlReader(new MemoryStream(new UTF8Encoding(false).GetBytes(yaml)), options)));
    }

    // YAML 1.2 section 5.2: the encoding comes from a byte order mark or the pattern of zero bytes at the start.
    [Theory]
    [InlineData("utf-8", false)]
    [InlineData("utf-8", true)]
    [InlineData("utf-16LE", false)]
    [InlineData("utf-16LE", true)]
    [InlineData("utf-16BE", false)]
    [InlineData("utf-16BE", true)]
    [InlineData("utf-32LE", false)]
    [InlineData("utf-32LE", true)]
    [InlineData("utf-32BE", false)]
    [InlineData("utf-32BE", true)]
    public void StreamInputDetectsTheEncoding(string name, bool byteOrderMark) {
        const string yaml = "key: värde 😀\nlist: [1, 2]\n";
        Encoding encoding = name switch {
            "utf-16LE" => new UnicodeEncoding(false, false),
            "utf-16BE" => new UnicodeEncoding(true, false),
            "utf-32LE" => new UTF32Encoding(false, false),
            "utf-32BE" => new UTF32Encoding(true, false),
            _ => new UTF8Encoding(false)
        };
        byte[] body = encoding.GetBytes(yaml);
        byte[] bytes = byteOrderMark ? encoding.GetBytes("﻿").Concat(body).ToArray() : body;
        string expected = YamlTextTestHelpers.Events(new YamlReader(yaml));
        using var reader = new YamlReader(new MemoryStream(bytes));
        Assert.Equal(expected, YamlTextTestHelpers.Events(reader));
        Assert.Equal(byteOrderMark, reader.SourceHasByteOrderMark);
    }

    [Fact]
    public void InvalidUtf8IsRejectedWithTheEncodingCode() {
        using var reader = new YamlReader(new MemoryStream(new byte[] { (byte)'a', (byte)':', (byte)' ', 0xC3, 0x28, (byte)'\n' }));
        YamlException error = Assert.ThrowsAny<YamlException>(() => { while (reader.Read()) { } });
        Assert.Equal(YamlErrorCodes.InvalidEncoding, error.Code);
    }
}
