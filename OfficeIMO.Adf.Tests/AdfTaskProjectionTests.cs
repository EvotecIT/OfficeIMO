using System;
using OfficeIMO.Adf;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Adf.Tests;

public sealed class AdfTaskProjectionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConversionDoesNotRenderUnsupportedMarkdownToCreateTaskIds(bool withTask) {
        MarkdownDoc markdown = MarkdownDoc.Create().Add(new UnrenderableBlock());
        if (withTask) markdown.Add(new UnorderedListBlock { Items = { ListItem.Task("Ready", done: true) } });

        AdfConversionResult<AdfDocument> result = AdfConverter.FromMarkdown(markdown);

        Assert.Equal(withTask ? 1 : 0, result.Value.Content.Count);
        if (withTask) Assert.False(string.IsNullOrWhiteSpace(result.Value.Content[0].GetStringAttribute("localId")));
        Assert.Contains(result.Report.Diagnostics, issue => issue.Code == "MARKDOWN_UNSUPPORTED_BLOCK");
    }

    [Fact]
    public void TocPlaceholdersDoNotExpandDuringTaskIdentityGeneration() {
        string source = string.Concat(System.Linq.Enumerable.Repeat("[TOC]\n\n", 128))
            + "# " + new string('H', 4096) + "\n\n- [ ] Ready";

        AdfDocument first = AdfConverter.FromMarkdown(source).Value;
        AdfDocument second = AdfConverter.FromMarkdown(source).Value;

        Assert.Equal(first.ToJson(), second.ToJson());
        Assert.Single(first.Content, node => node.Type == "taskList");
        Assert.Single(first.Content, node => node.Type == "heading");
    }

    [Fact]
    public void ConfiguredDepthAllowsUnrelatedDeepContentAlongsideTasks() {
        var outer = new UnorderedListBlock();
        UnorderedListBlock current = outer;
        for (int depth = 0; depth < 33; depth++) {
            var item = ListItem.Text("Level");
            var child = new UnorderedListBlock();
            item.NestedBlocks.Add(child);
            current.Items.Add(item);
            current = child;
        }
        current.Items.Add(ListItem.Text("Deep content"));
        MarkdownDoc markdown = MarkdownDoc.Create().Add(outer)
            .Add(new UnorderedListBlock { Items = { ListItem.Task("Ready", done: false) } });

        Assert.Throws<System.IO.InvalidDataException>(() => AdfConverter.FromMarkdown(markdown));
        var limits = new AdfConversionOptions { MaxDepth = 128 };
        AdfConversionResult<AdfDocument> result = AdfConverter.FromMarkdown(markdown, limits);

        Assert.Equal(2, result.Value.Content.Count);
        Assert.False(string.IsNullOrWhiteSpace(result.Value.Content[1].GetStringAttribute("localId")));
        Assert.Contains("Deep content", result.Value.ToJson(limits));
        Assert.Throws<System.IO.InvalidDataException>(() => result.Value.ToJson());
    }

    [Fact]
    public void TaskIdentitiesDoNotChangeWhenUnrelatedContentChanges() {
        AdfDocument Convert(string text) => AdfConverter.FromMarkdown(MarkdownDoc.Create()
            .Add(new ParagraphBlock(new InlineSequence().Text(text)))
            .Add(new UnorderedListBlock { Items = { ListItem.Task("Ready") } })).Value;

        AdfNode first = Convert("Before").Content[1];
        AdfNode second = Convert("After").Content[1];
        Assert.Equal(first.GetStringAttribute("localId"), second.GetStringAttribute("localId"));
        Assert.Equal(first.Content[0].GetStringAttribute("localId"), second.Content[0].GetStringAttribute("localId"));
    }

    [Fact]
    public void NestedTaskIdentitiesRespectTheCallerNodeBudgetWithoutDuplicatingRoots() {
        var outer = new UnorderedListBlock { Items = { ListItem.Task("Task") } };
        UnorderedListBlock current = outer;
        for (int depth = 1; depth < 8; depth++) {
            var child = new UnorderedListBlock { Items = { ListItem.Task("Task") } };
            current.Items[0].NestedBlocks.Add(child);
            current = child;
        }
        var limits = new AdfConversionOptions { MaxNodes = 64 };
        AdfDocument document = AdfConverter.FromMarkdown(MarkdownDoc.Create().Add(outer), limits).Value;
        Assert.True(document.Validate(new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema, MaxNodes = 64 }).IsValid);
        Assert.False(string.IsNullOrWhiteSpace(document.Content[0].GetStringAttribute("localId")));
        Assert.Contains("taskList", document.ToJson(limits));
    }

    [Fact]
    public void TaskIdentityCallbacksRunOnlyAfterGraphLimitsPass() {
        int calls = 0;
        var options = new AdfConversionOptions {
            MaxTextCharacters = 1,
            LocalIdFactory = path => { calls++; return path; }
        };
        var markdown = MarkdownDoc.Create().Add(new UnorderedListBlock { Items = { ListItem.Task("Oversized") } });
        Assert.Throws<System.IO.InvalidDataException>(() => AdfConverter.FromMarkdown(markdown, options));
        Assert.Equal(0, calls);
    }

    private sealed class UnrenderableBlock : IMarkdownBlock {
        public string RenderMarkdown() => throw new InvalidOperationException("Unsupported block was rendered.");
        public string RenderHtml() => throw new InvalidOperationException("Unsupported block was rendered.");
    }

    [Fact]
    public void TaskProjection_ReportsRegeneratedLocalIds() {
        var item = new AdfNode("taskItem") {
            Content = { AdfNode.TextNode("Ready") }
        }.SetAttribute("localId", "item-1").SetAttribute("state", "DONE");
        var list = new AdfNode("taskList") {
            Content = { item }
        }.SetAttribute("localId", "list-1");
        var document = new AdfDocument(new[] { list });

        Assert.True(document.Validate().IsValid);

        AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
        AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

        Assert.Equal("- [x] Ready", markdown.Value.Replace("\r\n", "\n"));
        Assert.False(markdown.Report.IsLossless);
        AdfConversionDiagnostic diagnostic = Assert.Single(
            markdown.Report.Diagnostics,
            item => item.Code == "ADF_TASK_LOCAL_IDS_REGENERATED");
        Assert.Equal("$.content[0]", diagnostic.Path);
        OfficeConversionFidelityDiagnostic fidelity = Assert.Single(markdown.Report.FidelityDiagnostics,
            item => item.Code == diagnostic.Code);
        Assert.Equal(OfficeConversionLossKind.Omission, fidelity.LossKind);
        Assert.Throws<InvalidOperationException>(markdown.Report.RequireNoLoss);
        AdfNode projectedList = Assert.Single(roundTrip.Value.Content);
        AdfNode projectedItem = Assert.Single(projectedList.Content);
        Assert.NotEqual("list-1", projectedList.GetStringAttribute("localId"));
        Assert.NotEqual("item-1", projectedItem.GetStringAttribute("localId"));
    }
}
