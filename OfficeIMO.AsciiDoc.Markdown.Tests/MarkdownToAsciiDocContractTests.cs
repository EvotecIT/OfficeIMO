namespace OfficeIMO.AsciiDoc.Markdown.Tests;

public sealed class MarkdownToAsciiDocContractTests {
    [Fact]
    public void RepresentativeMarkdown_ConvertsToCanonicalLosslessAsciiDoc() {
        const string markdown =
            "---\nproduct: OfficeIMO\n---\n\n" +
            "# Guide\n\n" +
            "## Start\n\n" +
            "Use **bold**, *italic*, `code`, [intro](#intro), and ![Icon](icon.svg).\n\n" +
            "- one\n- two\n\n" +
            "```csharp\nConsole.WriteLine();\n```\n\n" +
            "| Name | Value |\n| --- | --- |\n| A | B |\n";
        MarkdownDoc document = MarkdownReader.Parse(markdown);

        MarkdownToAsciiDocResult result = document.ToAsciiDocDocumentResult();

        Assert.StartsWith(":product: OfficeIMO\n\n= Guide\n\n== Start\n", result.Source, StringComparison.Ordinal);
        Assert.Contains("*bold*", result.Source, StringComparison.Ordinal);
        Assert.Contains("_italic_", result.Source, StringComparison.Ordinal);
        Assert.Contains("<<intro,intro>>", result.Source, StringComparison.Ordinal);
        Assert.Contains("image:icon.svg[Icon]", result.Source, StringComparison.Ordinal);
        Assert.Contains("[source,csharp]\n----", result.Source, StringComparison.Ordinal);
        Assert.Contains("[cols=2*,%header]\n|===", result.Source, StringComparison.Ordinal);
        Assert.Equal(result.Source, result.Value.ToAsciiDoc());
        Assert.True(result.Value.SyntaxTree.IsLossless);
        Assert.False(result.HasLoss);
    }

    [Fact]
    public void StructuredTableSpans_SurviveAsciiDocMarkdownAsciiDocRoundTrip() {
        const string source = "[cols=2*,%header]\n|===\n|A |B\n2+|wide\n|===\n";
        AsciiDocToMarkdownResult markdown = AsciiDocDocument.Parse(source).Document.ToMarkdownDocumentResult();

        MarkdownToAsciiDocResult roundTrip = markdown.Value.ToAsciiDocDocumentResult();
        AsciiDocTableBlock table = Assert.Single(roundTrip.Value.BlocksOfType<AsciiDocTableBlock>());

        Assert.Equal(2, table.Table.Cells[2].ColumnSpan);
        Assert.True(table.Table.Rows[0].IsHeader);
    }

    [Fact]
    public void DefinitionsCalloutsAndCompoundListChildren_MapBackToNativeBlocks() {
        var document = MarkdownDoc.Create();
        var definitions = new DefinitionListBlock();
        definitions.AddEntry(new DefinitionListEntry(
            new InlineSequence().Text("Term"),
            new IMarkdownBlock[] { new ParagraphBlock(new InlineSequence().Text("Definition")) }));
        document.Add(definitions);
        document.Add(new CalloutBlock("warning", string.Empty, new IMarkdownBlock[] {
            new ParagraphBlock(new InlineSequence().Text("Be careful"))
        }));
        var list = new UnorderedListBlock();
        var item = ListItem.Text("item");
        item.NestedBlocks.Add(new CodeBlock("text", "attached"));
        list.Items.Add(item);
        document.Add(list);

        MarkdownToAsciiDocResult result = document.ToAsciiDocDocumentResult();

        Assert.Contains("Term:: Definition", result.Source, StringComparison.Ordinal);
        Assert.Contains("WARNING: Be careful", result.Source, StringComparison.Ordinal);
        Assert.Contains("* item\n+\n[source,text]", result.Source, StringComparison.Ordinal);
        Assert.Single(result.Value.BlocksOfType<AsciiDocDescriptionListBlock>());
        Assert.Single(result.Value.BlocksOfType<AsciiDocAdmonitionBlock>());
        Assert.Single(result.Value.BlocksOfType<AsciiDocListContinuation>());
    }

    [Fact]
    public void UnsupportedMarkdownBlock_IsVisibleAndDiagnosed() {
        MarkdownDoc document = MarkdownDoc.Create().Hr();

        MarkdownToAsciiDocResult result = document.ToAsciiDocDocumentResult();

        Assert.Contains("[source,markdown]", result.Source, StringComparison.Ordinal);
        Assert.Contains("---", result.Source, StringComparison.Ordinal);
        Assert.Equal(AsciiDocMarkdownConversionOutcome.SourceFallback, Assert.Single(result.Report.Diagnostics).Outcome);
        Assert.True(result.HasLoss);
    }

    [Fact]
    public void RequestedLineEnding_IsUsedThroughoutGeneratedSource() {
        MarkdownDoc document = MarkdownReader.Parse("# Title\n\nParagraph\n");

        MarkdownToAsciiDocResult result = document.ToAsciiDocDocumentResult(new MarkdownToAsciiDocOptions { LineEnding = "\r\n" });

        Assert.DoesNotContain("\n", result.Source.Replace("\r\n", string.Empty), StringComparison.Ordinal);
        Assert.Equal(result.Source, result.Value.ToAsciiDoc());
    }

    [Fact]
    public void CodeContainingListingFence_UsesALongerDelimiterWithoutChangingContent() {
        MarkdownDoc document = MarkdownDoc.Create().Code("text", "before\n----\nafter");

        MarkdownToAsciiDocResult result = document.ToAsciiDocDocumentResult();

        Assert.Contains("-----\nbefore\n----\nafter\n-----", result.Source, StringComparison.Ordinal);
        AsciiDocDelimitedBlock block = Assert.Single(result.Value.BlocksOfType<AsciiDocDelimitedBlock>());
        Assert.Equal("before\n----\nafter\n", block.Content);
        Assert.DoesNotContain(result.Report.Diagnostics, static diagnostic => diagnostic.Feature == "code-delimiter");
    }

    [Fact]
    public void TableColumnCount_IncludesLogicalColumnSpans() {
        TableBlock table = Assert.Single(MarkdownReader.Parse("| H |\n| --- |\n| wide |\n").Blocks.OfType<TableBlock>());
        table.GetCell(0, 0)!.ColumnSpan = 2;

        MarkdownToAsciiDocResult result = MarkdownDoc.Create().Add(table).ToAsciiDocDocumentResult();

        Assert.Contains("[cols=2*", result.Source, StringComparison.Ordinal);
        Assert.Equal(2, Assert.Single(result.Value.BlocksOfType<AsciiDocTableBlock>()).Table.ColumnCount);
    }
}
