using OfficeIMO.Markdown;
using OfficeIMO.Toml;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_FrontMatter_Toml_Tests {
    private const string Post = "+++\n# Site settings\ntitle   = \"Draft\"   # working title\ndate = 2024-01-02\ntags = ['a', \"b\"]\nweight = 0x10\n\n[author]\nname = \"Ada\"\n+++\n\n# Body\n";

    [Fact]
    public void Toml_Front_Matter_Reads_Nested_Values_Through_The_Front_Matter_Entries() {
        var document = MarkdownDocument.Parse(Post);
        FrontMatterBlock frontMatter = document.DocumentHeader!;

        Assert.Equal(MarkdownFrontMatterFormat.Toml, frontMatter.Format);
        Assert.NotNull(frontMatter.Toml);
        Assert.Null(frontMatter.Yaml);
        Assert.Equal(new[] { "title", "date", "tags", "weight", "author" }, frontMatter.Entries.Select(entry => entry.Key).ToArray());
        Assert.Equal("Ada", frontMatter.Select("author.name")!.AsString());
        Assert.Equal("b", frontMatter.Select("tags[1]")!.AsString());
        Assert.True(frontMatter.TryGetValue("weight", out int weight) && weight == 16);
        Assert.True(document.TryGetFrontMatterValue("date", out string? date) && date == "2024-01-02");
        Assert.Empty(frontMatter.Diagnostics);
        Assert.Equal("Body", Assert.IsType<HeadingBlock>(Assert.Single(document.Blocks)).Text);
    }

    [Fact]
    public void Edited_Toml_Front_Matter_Saves_Unedited_Text_Byte_For_Byte_And_Rewrites_Only_The_Change() {
        var document = MarkdownDocument.Parse(Post);
        var save = new MarkdownSaveOptions { OutputLineEnding = "\n" };
        string frontMatter = Post.Substring(0, Post.IndexOf("\n+++\n", StringComparison.Ordinal) + 5);

        Assert.StartsWith(frontMatter, document.ToMarkdown(save), StringComparison.Ordinal);

        document.DocumentHeader!.Toml!.Set("title", "Final");

        Assert.Equal("Final", document.FindFrontMatterEntry("title")!.Value.AsString());
        Assert.StartsWith(frontMatter.Replace("\"Draft\"", "\"Final\""), document.ToMarkdown(save), StringComparison.Ordinal);
    }

    [Fact]
    public void Invalid_Toml_Front_Matter_Is_Kept_As_Text_With_A_Markdown_Positioned_Error() {
        const string markdown = "+++\ntitle = \"ok\"\ntitle = \"again\"\n+++\n\n# Body\n";

        var document = MarkdownDocument.Parse(markdown);
        FrontMatterBlock frontMatter = document.DocumentHeader!;

        Assert.Null(frontMatter.Toml);
        Assert.Empty(frontMatter.Entries);
        OfficeDiagnostic diagnostic = Assert.Single(frontMatter.Diagnostics);
        Assert.Equal(OfficeDiagnosticCodes.MarkdownFrontMatterInvalidToml, diagnostic.Code);
        Assert.Equal(OfficeDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.Contains(TomlErrorCodes.DuplicateKey, diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("line 3, column 1", diagnostic.Location);
        Assert.Equal("Body", Assert.IsType<HeadingBlock>(Assert.Single(document.Blocks)).Text);
        Assert.StartsWith("+++\ntitle = \"ok\"\ntitle = \"again\"\n+++\n", document.ToMarkdown(new MarkdownSaveOptions { OutputLineEnding = "\n" }), StringComparison.Ordinal);
    }

    [Fact]
    public void Toml_Front_Matter_Built_In_Code_Is_Written_Between_Plus_Fences() {
        TomlDocument toml = TomlDocument.Create();
        toml.Set("title", "Release");
        toml.Set("params.draft", true);
        var document = MarkdownDocument.Create();
        document.Add(FrontMatterBlock.FromToml(toml));
        document.H1("Body");

        string markdown = document.ToMarkdown(new MarkdownSaveOptions { OutputLineEnding = "\n" });
        FrontMatterBlock parsed = MarkdownDocument.Parse(markdown).DocumentHeader!;

        Assert.StartsWith("+++\n", markdown, StringComparison.Ordinal);
        Assert.Equal(MarkdownFrontMatterFormat.Toml, parsed.Format);
        Assert.Equal("Release", parsed.FindEntry("title")!.Value.AsString());
        Assert.True(parsed.Select("params.draft")!.GetValue<bool>());
    }

    [Fact]
    public void Toml_Front_Matter_Entries_Carry_Key_And_Value_Source_Spans() {
        var native = MarkdownNativeDocument.Parse(Post);
        var frontMatter = Assert.IsType<MarkdownNativeFrontMatterBlock>(native.Blocks[0]);

        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 3), frontMatter.OpeningFenceSourceSpan);
        FrontMatterBlock.Entry title = frontMatter.Entries[0];
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 5), title.KeySourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 11, 3, 17), title.ValueSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 8, 5, 17), frontMatter.Entries[2].ValueSourceSpan);
        FrontMatterBlock.Entry author = frontMatter.Entries[4];
        Assert.Equal(new MarkdownSourceSpan(8, 2, 8, 7), author.KeySourceSpan);
        Assert.Null(author.ValueSourceSpan);
        Assert.Equal("+++", native.EnumerateBlockSourceFields("closingFence").Single().Value);
    }
}
