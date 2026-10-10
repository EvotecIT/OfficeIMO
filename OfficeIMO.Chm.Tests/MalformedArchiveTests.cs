using System.Text;

namespace OfficeIMO.Chm.Tests;

public sealed class MalformedArchiveTests {
    [Theory]
    [InlineData(12598 + 16, 3U, "CHM_COMPRESSION_METADATA")]
    [InlineData(12598 + 20, uint.MaxValue, "CHM_COMPRESSION_METADATA")]
    [InlineData(12626 + 4, uint.MaxValue, "CHM_BOUNDS")]
    [InlineData(12626 + 40, 1U, "CHM_COMPRESSION_METADATA")]
    public void InvalidNativeLzxMetadataIsRejectedBeforeDecoding(int offset, uint value, string code) {
        // Absolute offsets identify the pinned fixture's uncompressed ControlData and ResetTable.
        byte[] bytes = File.ReadAllBytes(ChmFixture.NativePath);
        Assert.Equal("LZXC", Encoding.ASCII.GetString(bytes, 12598 + 4, 4));
        SetUInt32(bytes, offset, value);
        Assert.Equal(code, Assert.Throws<ChmReadException>(() => ChmDocument.Load(bytes)).Code);
    }

    [Theory]
    [InlineData("cycle", "CHM_CONTENTS")]
    [InlineData("bounds", "CHM_TRUNCATED")]
    [InlineData("topic", "CHM_TOPIC_TABLE")]
    public void CompiledContentsRejectCyclesBoundsAndInvalidTopicReferences(string fault, string code) {
        ChmDocument source = ChmDocument.Load(ChmFixture.NativePath);
        var entries = new[] { "/#TOCIDX", "/#TOPICS", "/#STRINGS", "/#URLTBL", "/#URLSTR" }
            .ToDictionary(path => path, path => source.FindEntry(path)!.GetBytes());
        byte[] contents = entries["/#TOCIDX"];
        int first = BitConverter.ToInt32(contents, 0);
        if (fault == "cycle") SetUInt32(contents, first + 16, (uint)first);
        if (fault == "bounds") SetUInt32(contents, 0, (uint)contents.Length);
        if (fault == "topic") { SetUInt32(contents, first + 4, 8); SetUInt32(contents, first + 8, uint.MaxValue); }
        Assert.Equal(code, Assert.Throws<ChmReadException>(() => ChmDocument.Load(ChmFixture.Archive(entries))).Code);
    }

    [Theory]
    [InlineData(1041U, "shift_jis", "日本語")]
    [InlineData(1045U, "windows-1250", "Łódź")]
    public void LocaleSelectsLegacyEncodingForUndeclaredTopics(uint locale, string label, string text) {
        var options = new ChmReadOptions();
        Encoding encoding = options.EncodingProvider.ResolveLabel(label)!;
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/#SYSTEM"] = ChmFixture.SystemMetadata(locale), ["/topic.html"] = encoding.GetBytes("<p>" + text + "</p>")
        }), options);
        Assert.Contains(text, Assert.Single(book.Topics).ReadHtml());
        Assert.DoesNotContain(book.Diagnostics, item => item.Code == "CHM_ENCODING_FALLBACK");
    }

    private static void SetUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)value; bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16); bytes[offset + 3] = (byte)(value >> 24);
    }

    [Fact]
    public void CompiledTopicReferencesCannotAmplifyUnboundedStringsOrRecordCounts() {
        ChmDocument source = ChmDocument.Load(ChmFixture.NativePath);
        var entries = new[] { "/#TOPICS", "/#STRINGS", "/#URLTBL", "/#URLSTR" }
            .ToDictionary(path => path, path => source.FindEntry(path)!.GetBytes());
        Assert.Equal("CHM_NAVIGATION_LIMIT", Assert.Throws<ChmReadException>(() => ChmDocument.Load(ChmFixture.Archive(entries),
            new ChmReadOptions { MaxNavigationItems = 10 })).Code);
        entries["/#STRINGS"] = Encoding.ASCII.GetBytes(new string('X', 5000) + '\0');
        SetUInt32(entries["/#TOPICS"], 4, 0);
        Assert.Throws<ChmReadException>(() => ChmDocument.Load(ChmFixture.Archive(entries)));
    }
}
