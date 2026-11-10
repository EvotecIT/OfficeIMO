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
    [InlineData("apiVersion", " file header")]
    [InlineData("metadata", " above metadata")]
    [InlineData("metadata.name", " above name")]
    [InlineData("empty", " after the script")]
    public void CommentsOnTheirOwnLinesLeadTheNextKey(string key, string comment) {
        Assert.Equal(new[] { comment }, Key(key).LeadingComments);
    }

    [Fact]
    public void BlockCollectionsWithoutPropertiesLeaveLeadingCommentsToTheirFirstChild() {
        Assert.Empty(Document.Root!.LeadingComments);
        Assert.Empty(Document.Select("items")!.LeadingComments);
        Assert.Equal(new[] { " above first item" }, Document.Select("items[0]")!.LeadingComments);
        Assert.Equal(new[] { " above three" }, Document.Select("items[2]")!.LeadingComments);
        Assert.Empty(Document.Select("items[1]")!.LeadingComments);
    }

    [Theory]
    [InlineData("apiVersion", " trailing on value")]
    [InlineData("items[1]", " after two")]
    [InlineData("script", " header comment")]
    [InlineData("empty", " on empty value")]
    [InlineData("flow[0]", " after a")]
    public void TrailingCommentsBelongToTheNodeEndingOnTheirLine(string path, string comment) {
        Assert.Equal(comment, Document.Select(path)!.TrailingComment);
    }

    [Fact]
    public void KeysKeepTrailingCommentsOnlyWhenTheirValueStartsLater() {
        Assert.Equal(" on labels key", Key("metadata.labels").TrailingComment);
        Assert.Null(Key("empty").TrailingComment);
        Assert.Null(Document.Select("metadata.labels")!.TrailingComment);
        Assert.Null(Document.Select("flow[1]")!.TrailingComment);
        Assert.Empty(Document.Select("flow[1]")!.LeadingComments);
    }

    [Fact]
    public void DocumentMarkersBoundTheRootsLeadingComments() {
        YamlDocument document = YamlDocument.Parse("# before marker\n--- # on marker\n# above root\nkey: v\n");
        YamlNode key = document.Root!.AsMapping()[0].Key;
        Assert.Equal(new[] { " on marker", " above root" }, key.LeadingComments);
    }

    [Fact]
    public void AnchoredBlockCollectionsOwnTheCommentsBeforeTheirProperties() {
        YamlDocument document = YamlDocument.Parse("base:\n  # about base\n  &b\n  # first\n  x: 1\n");
        YamlNode value = document.Select("base")!;
        Assert.Equal(new[] { " about base" }, value.LeadingComments);
        Assert.Equal(new[] { " first" }, value.AsMapping()[0].Key.LeadingComments);
    }
}
