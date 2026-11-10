using Xunit;

namespace OfficeIMO.Yaml.Tests;

// Comment attachment is a model rule with no suite equivalent; each case pins one clause of the documented rule.
public class YamlDocumentCommentTests {
    private const string Yaml = """
        # file header
        apiVersion: v1 # trailing on value
        # above metadata

        metadata:
          # above name
          name: web
          labels: # on labels key
            app: web
        items:
          # above first item
          - one
          - two # after two
          # above three
          - three
        script: | # header comment
          echo hi
        # after the script
        empty: # on empty value
        flow: [a, # after a
          b]
        # footer
        """;

    private static readonly YamlDocument Document = YamlDocument.Parse(Yaml + "\n");

    private static YamlNode Key(string path) {
        int dot = path.LastIndexOf('.');
        YamlMapping parent = dot < 0 ? Document.Root!.AsMapping() : Document.Select(path.Substring(0, dot))!.AsMapping();
        string name = dot < 0 ? path : path.Substring(dot + 1);
        foreach (YamlMappingEntry entry in parent) {
            if (entry.Key.AsString() == name) return entry.Key;
        }
        throw new KeyNotFoundException(path);
    }

    [Theory]
    [InlineData("apiVersion", "file header")]
    [InlineData("metadata", "above metadata")]
    [InlineData("metadata.name", "above name")]
    [InlineData("empty", "after the script")]
    public void CommentsOnTheirOwnLinesLeadTheNextKey(string key, string comment) {
        Assert.Equal(comment, Key(key).CommentBefore);
        Assert.Equal(comment, Document.Select(key)!.CommentBefore);
    }

    [Fact]
    public void FirstEntriesShareCommentsWithTheirSequenceItemOrRoot() {
        Assert.Equal("file header", Document.Root!.CommentBefore);
        Assert.Null(Document.Select("items")!.CommentBefore);
        Assert.Equal("above first item", Document.Select("items[0]")!.CommentBefore);
        Assert.Equal("above three", Document.Select("items[2]")!.CommentBefore);
        Assert.Null(Document.Select("items[1]")!.CommentBefore);
    }

    [Theory]
    [InlineData("apiVersion", "trailing on value")]
    [InlineData("items[1]", "after two")]
    [InlineData("script", "header comment")]
    [InlineData("empty", "on empty value")]
    [InlineData("flow[0]", "after a")]
    [InlineData("metadata.labels", "on labels key")]
    public void TrailingCommentsBelongToTheNodeEndingOnTheirLine(string path, string comment) {
        Assert.Equal(comment, Document.Select(path)!.Comment);
    }

    [Fact]
    public void KeysKeepTrailingCommentsOnlyWhenTheirValueStartsLater() {
        Assert.Equal("on labels key", Key("metadata.labels").Comment);
        Assert.Null(Key("empty").Comment);
        Assert.Null(Document.Select("flow[1]")!.Comment);
        Assert.Null(Document.Select("flow[1]")!.CommentBefore);
    }

    [Fact]
    public void DocumentMarkersBoundTheRootsComments() {
        YamlDocument document = YamlDocument.Parse("# before marker\n--- # on marker\n# above root\nkey: v\n");
        YamlNode key = document.Root!.AsMapping()[0].Key;
        Assert.Equal("above root", key.CommentBefore);
        Assert.Equal("before marker", document.Documents[0].CommentBefore);
    }

    [Fact]
    public void BlankLineSeparatesTheDocumentHeaderFromTheFirstEntry() {
        YamlDocument document = YamlDocument.Parse("# license\n# text\n\n# about a\na: 1\nb: 2\n# footer\n");
        Assert.Equal("license\ntext", document.Documents[0].CommentBefore);
        Assert.Equal("about a", document.Select("a")!.CommentBefore);
        Assert.True(document.Select("a")!.SpaceBefore);
        Assert.Equal("footer", document.Documents[0].CommentAfter);
    }

    [Fact]
    public void AnchoredBlockCollectionsKeepTheirFirstEntrysComments() {
        YamlDocument document = YamlDocument.Parse("base:\n  # about base\n  &b\n  # first\n  x: 1\n");
        Assert.Equal("first", document.Select("base")!.AsMapping()[0].Key.CommentBefore);
    }
}
