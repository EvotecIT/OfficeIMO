using System.Text;
using Xunit;

namespace OfficeIMO.Toml.Tests;

// The pull reader's token contract and the writer's output; toml-test exercises neither API directly.
public class TomlReaderAndWriterTests {
    [Fact]
    public void TheReaderReportsHeadersKeysValuesAndPositions() {
        var reader = new TomlReader("# top\n[a.\"b c\"]\nx.y = [1, { z = 'q' }] # end\n", new TomlReaderOptions { IncludeComments = true });
        var tokens = new List<string>();
        while (reader.Read()) {
            string text = reader.TokenType is TomlTokenType.Key or TomlTokenType.String or TomlTokenType.Comment ? reader.GetString() : reader.ValueSpan.ToString();
            tokens.Add(reader.TokenType + ":" + text + "@" + reader.Start + "/" + reader.Depth);
        }
        Assert.Equal(new[] {
            "Comment: top@1:1/0",
            "TableHeaderStart:[@2:1/0", "Key:a@2:2/0", "Key:b c@2:4/0", "TableHeaderEnd:]@2:9/0",
            "Key:x@3:1/0", "Key:y@3:3/0", "StartArray:[@3:7/0", "Integer:1@3:8/1", "StartInlineTable:{@3:11/1",
            "Key:z@3:13/2", "String:q@3:17/2", "EndInlineTable:}@3:21/1", "EndArray:]@3:22/0", "Comment: end@3:24/0"
        }, tokens);
    }

    [Fact]
    public void TypedReaderAccessorsDecodeTheCurrentToken() {
        var reader = new TomlReader("a = 0b101\nb = -inf\nc = 1979-05-27T07:32Z\nd = true\n");
        var values = new List<object>();
        while (reader.Read()) {
            switch (reader.TokenType) {
                case TomlTokenType.Integer: values.Add(reader.GetInt64()); break;
                case TomlTokenType.Float: values.Add(reader.GetDouble()); break;
                case TomlTokenType.OffsetDateTime: values.Add(reader.GetDateTime().ToDateTimeOffset()); break;
                case TomlTokenType.Boolean: values.Add(reader.GetBoolean()); break;
            }
        }
        Assert.Equal(new object[] { 5L, double.NegativeInfinity, new DateTimeOffset(1979, 5, 27, 7, 32, 0, TimeSpan.Zero), true }, values);
    }

    [Fact]
    public void TheWriterProducesTomlThatReadsBackToTheSameValues() {
        var output = new StringWriter();
        using (var writer = new TomlWriter(output)) {
            writer.WriteComment("generated");
            writer.WriteKey("name");
            writer.WriteString("it's \"quoted\"\n", TomlStringStyle.Literal);
            writer.WriteKey("site", "url");
            writer.WriteString("https://example.com", TomlStringStyle.Literal);
            writer.WriteComment("tools");
            writer.WriteTableHeader("tool", "x y");
            writer.WriteKey("list");
            writer.WriteStartArray();
            writer.WriteInteger(1);
            writer.WriteFloat(0.5);
            writer.WriteStartInlineTable();
            writer.WriteKey("a");
            writer.WriteBoolean(false);
            writer.WriteEndInlineTable();
            writer.WriteEndArray();
            writer.WriteArrayOfTablesHeader("p");
            writer.WriteKey("when");
            writer.WriteDateTime(TomlDateTime.FromDate(2024, 2, 29));
            Assert.Throws<InvalidOperationException>(() => writer.WriteInteger(1));
        }
        Assert.Equal(
            "# generated\nname = \"it's \\\"quoted\\\"\\n\"\nsite.url = 'https://example.com'\n\n# tools\n[tool.\"x y\"]\nlist = [1, 0.5, { a = false }]\n\n[[p]]\nwhen = 2024-02-29\n",
            output.ToString());
        using TomlDocument doc = TomlDocument.Parse(output.ToString());
        Assert.Equal("it's \"quoted\"\n", doc.Select("name")!.AsString());
        Assert.False(doc.Select("tool[\"x y\"].list[2].a")!.GetValue<bool>());
        Assert.Equal(new[] { " tools" }, doc.Select("tool[\"x y\"]")!.LeadingComments);
    }
}
