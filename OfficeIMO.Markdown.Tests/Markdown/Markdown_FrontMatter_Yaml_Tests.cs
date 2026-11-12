using System.Text.Json;
using OfficeIMO.Markdown;
using OfficeIMO.Yaml;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_FrontMatter_Yaml_Tests {
    [Fact]
    public void Nested_Front_Matter_Reads_As_Yaml_And_Keeps_Top_Level_Entries() {
        const string markdown = "---\ntitle: Guide\nauthor:\n  name: Ada\n  links: [site, blog]\nmenu:\n  - main\n  - docs\n---\n\n# Guide\n";

        var document = MarkdownDocument.Parse(markdown);
        FrontMatterBlock frontMatter = document.DocumentHeader!;

        Assert.Equal(new[] { "title", "author", "menu" }, frontMatter.Entries.Select(entry => entry.Key).ToArray());
        Assert.Equal("Ada", frontMatter.Select("author.name")!.AsString());
        Assert.Equal("blog", frontMatter.Select("author.links[1]")!.AsString());
        Assert.Equal(2, frontMatter.FindEntry("menu")!.Value.AsSequence().Count);
        Assert.Empty(frontMatter.Diagnostics);
    }

    [Fact]
    public void Edited_Front_Matter_Rewrites_Only_The_Changed_Value() {
        const string markdown = "---\n# Site settings\ntitle: Draft   # working title\ntags: [a, b]\n---\n\n# Body\n";

        var document = MarkdownDocument.Parse(markdown);
        document.DocumentHeader!.Yaml!.Set("title", "Final");
        string output = document.ToMarkdown(new MarkdownSaveOptions { OutputLineEnding = "\n" });

        Assert.StartsWith("---\n# Site settings\ntitle: Final   # working title\ntags: [a, b]\n---\n\n# Body", output, StringComparison.Ordinal);
    }

    [Fact]
    public void Invalid_Front_Matter_Is_Kept_As_Text_With_A_Markdown_Positioned_Error() {
        const string markdown = "---\ntitle: ok\nbad: [unclosed\n---\n\n# Body\n";

        var document = MarkdownDocument.Parse(markdown);
        FrontMatterBlock frontMatter = document.DocumentHeader!;

        Assert.Null(frontMatter.Yaml);
        Assert.Empty(frontMatter.Entries);
        OfficeDiagnostic diagnostic = Assert.Single(frontMatter.Diagnostics);
        Assert.Equal(OfficeDiagnosticCodes.MarkdownFrontMatterInvalidYaml, diagnostic.Code);
        Assert.Equal(OfficeDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.StartsWith("line ", diagnostic.Location, StringComparison.Ordinal);
        Assert.True(int.Parse(diagnostic.Location!.Split(' ', ',')[1]) >= 3, diagnostic.Location);
        Assert.Equal("Body", Assert.IsType<HeadingBlock>(Assert.Single(document.Blocks)).Text);
        Assert.StartsWith("---\ntitle: ok\nbad: [unclosed\n---\n", document.ToMarkdown(new MarkdownSaveOptions { OutputLineEnding = "\n" }), StringComparison.Ordinal);
    }

    [Fact]
    public void Front_Matter_Built_From_Values_Reads_Back_With_The_Same_Types() {
        var document = MarkdownDocument.Create()
            .FrontMatter(new Dictionary<string, object?> {
                ["enabled"] = "yes",
                ["version"] = "1.0",
                ["title"] = "Release: notes # 1",
                ["draft"] = false,
                ["count"] = 3,
                ["tags"] = new[] { "on", "docs" },
                ["author"] = new Dictionary<string, object?> { ["name"] = "Ada", ["since"] = "2024-01-02" }
            })
            .H1("Body");

        FrontMatterBlock parsed = MarkdownDocument.Parse(document.ToMarkdown()).DocumentHeader!;

        Assert.Empty(parsed.Diagnostics);
        Assert.Equal(YamlScalarKind.String, parsed.FindEntry("enabled")!.Value.AsScalar().Kind);
        Assert.Equal("yes", parsed.FindEntry("enabled")!.Value.AsString());
        Assert.Equal(YamlScalarKind.String, parsed.FindEntry("version")!.Value.AsScalar().Kind);
        Assert.Equal("Release: notes # 1", parsed.FindEntry("title")!.Value.AsString());
        Assert.True(parsed.TryGetValue("draft", out bool draft) && !draft);
        Assert.True(parsed.TryGetValue("count", out int count) && count == 3);
        Assert.Equal(YamlScalarKind.String, parsed.Select("tags[0]")!.AsScalar().Kind);
        Assert.Equal("Ada", parsed.Select("author.name")!.AsString());
        Assert.Equal(YamlScalarKind.String, parsed.Select("author.since")!.AsScalar().Kind);
    }

    [Fact]
    public void Committed_Front_Matter_Corpus_Round_Trips_Byte_For_Byte() {
        string[] files = Directory.GetFiles(Path.Combine(AppContext.BaseDirectory, "Yaml", "FrontMatter"), "*.md");
        Assert.NotEmpty(files);
        Assert.Empty(CheckRoundTrips(files));
    }

    /// <summary>
    /// Opt-in: set OFFICEIMO_YAML_CORPUS_CACHE to the large YAML corpus cache (fetched by
    /// OfficeIMO.TestAssets/Yaml/LargeCorpus/Update-YamlLargeCorpus.ps1) to check its Jekyll, Hugo and docs front matter.
    /// </summary>
    [Fact]
    public void Large_Front_Matter_Corpus_Round_Trips_Byte_For_Byte_When_Configured() {
        string? root = Environment.GetEnvironmentVariable("OFFICEIMO_YAML_CORPUS_CACHE");
        if (string.IsNullOrWhiteSpace(root)) return;

        using JsonDocument index = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(root!, "index.json")));
        var files = new List<string>();
        foreach (JsonElement file in index.RootElement.GetProperty("files").EnumerateArray()) {
            if (file.GetProperty("kind").GetString() != "front-matter" || file.GetProperty("excluded").ValueKind == JsonValueKind.String) continue;
            files.Add(Path.Combine(root!, "sources", file.GetProperty("source").GetString()!, file.GetProperty("path").GetString()!.Replace('/', Path.DirectorySeparatorChar)));
        }

        Assert.NotEmpty(files);
        List<string> failures = CheckRoundTrips(files);
        Assert.True(failures.Count == 0, failures.Count + " of " + files.Count + " files failed:\n" + string.Join("\n", failures.Take(25)));
    }

    /// <summary>Parses each file and checks that Markdown output starts with its front matter unchanged.</summary>
    private static List<string> CheckRoundTrips(IEnumerable<string> files) {
        var failures = new List<string>();
        foreach (string path in files) {
            string text = File.ReadAllText(path).TrimStart('﻿').Replace("\r\n", "\n");
            int closing = FindClosingFence(text);
            if (closing < 0) {
                failures.Add(path + ": no --- front matter");
                continue;
            }

            string frontMatter = text.Substring(0, closing);
            MarkdownDocument document;
            try {
                document = MarkdownDocument.Parse(text);
            } catch (Exception exception) {
                failures.Add(path + ": " + exception.GetType().Name + " " + exception.Message);
                continue;
            }

            FrontMatterBlock? header = document.DocumentHeader;
            if (header?.Yaml == null) {
                failures.Add(path + ": " + (header?.Diagnostics.FirstOrDefault()?.Message ?? "no front matter block"));
                continue;
            }

            string output = document.ToMarkdown(new MarkdownSaveOptions { OutputLineEnding = "\n" });
            if (!output.StartsWith(frontMatter, StringComparison.Ordinal)) failures.Add(path + ": front matter changed on save");
        }

        return failures;
    }

    /// <summary>Returns the index just past the closing <c>---</c> line, as the Markdown reader finds it.</summary>
    private static int FindClosingFence(string text) {
        int lineEnd = text.IndexOf('\n');
        if (lineEnd < 0 || text.Substring(0, lineEnd).Trim() != "---") return -1;
        int position = lineEnd + 1;
        while (position < text.Length) {
            int next = text.IndexOf('\n', position);
            string line = next < 0 ? text.Substring(position) : text.Substring(position, next - position);
            if (line.Trim() == "---") return next < 0 ? text.Length : next + 1;
            if (next < 0) break;
            position = next + 1;
        }

        return -1;
    }
}
