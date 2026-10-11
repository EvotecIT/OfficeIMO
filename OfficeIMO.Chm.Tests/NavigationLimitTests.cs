using System.Text;

namespace OfficeIMO.Chm.Tests;

public sealed class NavigationLimitTests {
    [Fact]
    public void CompiledIndexTargetsUseOneAggregateReferenceBudget() {
        var entries = TopicTable(); entries["/$WWKeywordLinks/BTree"] = KeywordIndex(3, 8);
        byte[] archive = ChmFixture.Archive(entries);
        Assert.Equal("CHM_NAVIGATION_LIMIT", Assert.Throws<ChmReadException>(() => ChmDocument.Load(archive,
            new ChmReadOptions { MaxNavigationItems = 3, MaxNavigationReferences = 23 })).Code);
        ChmDocument book = ChmDocument.Load(archive, new ChmReadOptions { MaxNavigationItems = 3, MaxNavigationReferences = 24 });
        Assert.Equal(24, book.Index.Sum(item => item.Links.Count));
    }

    [Fact]
    public void ReusedCompiledTopicOffsetsCannotAmplifyNavigationText() {
        var entries = TopicTable(8, new string('T', 64));
        Assert.Equal("CHM_NAVIGATION_LIMIT", Assert.Throws<ChmReadException>(() => ChmDocument.Load(ChmFixture.Archive(entries),
            new ChmReadOptions { MaxNavigationItems = 8, MaxNavigationCharacters = 100 })).Code);
    }

    [Theory]
    [InlineData("Local")]
    [InlineData("Merge")]
    [InlineData("See Also")]
    public void ContentsAndIndexSitemapReferencesShareOneBudget(string kind) {
        string objectMarkup = "<object type='text/sitemap'><param name='Name' value='Topic'><param name='" + kind + "' value='a.html'></object>";
        var entries = new Dictionary<string, byte[]> {
            ["/a.html"] = ChmFixture.Html("<p>Topic</p>"),
            ["/contents.hhc"] = ChmFixture.Html("<ul><li>" + objectMarkup + "</ul>"),
            ["/index.hhk"] = ChmFixture.Html("<ul><li>" + objectMarkup + "</ul>")
        };
        Assert.Equal("CHM_NAVIGATION_LIMIT", Assert.Throws<ChmReadException>(() => ChmDocument.Load(ChmFixture.Archive(entries),
            new ChmReadOptions { MaxNavigationReferences = 1 })).Code);
    }

    [Theory]
    [InlineData("contents")]
    [InlineData("index")]
    [InlineData("sitemap")]
    public void NavigationHeadingsUseTheAggregateTextBudget(string format) {
        string heading = new string('H', 80);
        var entries = TopicTable();
        if (format == "contents") {
            entries["/#STRINGS"] = Encoding.ASCII.GetBytes("Topic\0" + heading + '\0');
            byte[] contents = new byte[36]; BitConverter.GetBytes(16U).CopyTo(contents, 0);
            BitConverter.GetBytes(1U).CopyTo(contents, 8); BitConverter.GetBytes(6U).CopyTo(contents, 24);
            entries["/#TOCIDX"] = contents;
        } else if (format == "index") entries["/$WWKeywordLinks/BTree"] = KeywordIndex(1, 0, heading);
        else entries["/contents.hhc"] = ChmFixture.Html("<object type='text/sitemap'><param name='Name' value='" + heading + "'></object>");
        Assert.Equal("CHM_NAVIGATION_LIMIT", Assert.Throws<ChmReadException>(() => ChmDocument.Load(ChmFixture.Archive(entries),
            new ChmReadOptions { MaxNavigationCharacters = 80 })).Code);
    }

    private static Dictionary<string, byte[]> TopicTable(int records = 1, string title = "Topic") => new Dictionary<string, byte[]> {
        ["/#TOPICS"] = new byte[records * 16], ["/#STRINGS"] = Encoding.ASCII.GetBytes(title + '\0'),
        ["/#URLTBL"] = new byte[12], ["/#URLSTR"] = new byte[8].Concat(Encoding.ASCII.GetBytes("a.html\0")).ToArray(),
        ["/a.html"] = ChmFixture.Html("<p>Topic</p>")
    };

    private static byte[] KeywordIndex(int recordCount, int pairs, string? heading = null) {
        const int blockLength = 512;
        using var records = new MemoryStream(); using var writer = new BinaryWriter(records);
        for (int i = 0; i < recordCount; i++) {
            writer.Write(Encoding.Unicode.GetBytes((heading ?? "Keyword" + i) + '\0'));
            writer.Write((ushort)0); writer.Write((ushort)0); writer.Write(0U); writer.Write(0U); writer.Write((uint)pairs);
            for (int pair = 0; pair < pairs; pair++) writer.Write(0U);
            writer.Write(0UL);
        }
        byte[] result = new byte[76 + blockLength];
        using var stream = new MemoryStream(result); using var output = new BinaryWriter(stream);
        output.Write((ushort)0x293B); stream.Position = 4; output.Write((ushort)blockLength);
        stream.Position = 38; output.Write(1U);
        stream.Position = 76; output.Write((ushort)(blockLength - 12 - records.Length)); output.Write((ushort)recordCount);
        stream.Position = 88; output.Write(records.ToArray()); return result;
    }
}
