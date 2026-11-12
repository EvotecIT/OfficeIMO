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
        MarkdownDocument document = MarkdownDocument.Parse(markdown);

        var result = document.ToAsciiDoc(null, out OfficeIMO.OfficeConversionReport conversionReport1);
        var resultReport = (MarkdownToAsciiDocReport)conversionReport1;

        Assert.StartsWith(":product: OfficeIMO\n\n= Guide\n\n== Start\n", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("*bold*", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("_italic_", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("<<intro,intro>>", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("image:icon.svg[Icon]", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("[source,csharp]\n----", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("[cols=2*,%header]\n|===", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Equal(resultReport.GeneratedSource, result.ToAsciiDoc());
        Assert.True(result.SyntaxTree.IsLossless);
        Assert.False(resultReport.HasLoss);
    }

    [Fact]
    public void NestedFrontMatter_BecomesOneAttributePerScalar() {
        MarkdownDocument document = MarkdownDocument.Parse("---\ntitle: Guide\nauthor:\n  name: Ada\n  email: ada@example.com\ntags: [docs, yaml]\n---\n\nBody\n");

        document.ToAsciiDoc(null, out OfficeIMO.OfficeConversionReport conversionReport);
        string source = ((MarkdownToAsciiDocReport)conversionReport).GeneratedSource;

        Assert.StartsWith(":title: Guide\n:author-name: Ada\n:author-email: ada@example.com\n:tags: docs, yaml\n\n", source, StringComparison.Ordinal);
    }

    [Fact]
    public void StructuredTableSpans_SurviveAsciiDocMarkdownAsciiDocRoundTrip() {
        const string source = "[cols=2*,%header]\n|===\n|A |B\n2+|wide\n|===\n";
        var markdown = AsciiDocDocument.ParseResult(source).Document.ToMarkdown(null);

        var roundTrip = markdown.ToAsciiDoc(null);
        AsciiDocTableBlock table = Assert.Single(roundTrip.BlocksOfType<AsciiDocTableBlock>());

        Assert.Equal(2, table.Table.Cells[2].ColumnSpan);
        Assert.True(table.Table.Rows[0].IsHeader);
    }

    [Fact]
    public void DefinitionsCalloutsAndCompoundListChildren_MapBackToNativeBlocks() {
        var document = MarkdownDocument.Create();
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

        var result = document.ToAsciiDoc(null, out OfficeIMO.OfficeConversionReport conversionReport4);
        var resultReport = (MarkdownToAsciiDocReport)conversionReport4;

        Assert.Contains("Term:: Definition", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("WARNING: Be careful", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("* item\n+\n[source,text]", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Single(result.BlocksOfType<AsciiDocDescriptionListBlock>());
        Assert.Single(result.BlocksOfType<AsciiDocAdmonitionBlock>());
        Assert.Single(result.BlocksOfType<AsciiDocListContinuation>());
    }

    [Fact]
    public void UnsupportedMarkdownBlock_IsVisibleAndDiagnosed() {
        MarkdownDocument document = MarkdownDocument.Create().Hr();

        var result = document.ToAsciiDoc(null, out OfficeIMO.OfficeConversionReport conversionReport5);
        var resultReport = (MarkdownToAsciiDocReport)conversionReport5;

        Assert.Contains("[source,markdown]", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("---", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Equal(AsciiDocMarkdownConversionOutcome.SourceFallback, Assert.Single(resultReport.NativeDiagnostics).Outcome);
        Assert.True(resultReport.HasLoss);
    }

    [Fact]
    public void RequestedLineEnding_IsUsedThroughoutGeneratedSource() {
        MarkdownDocument document = MarkdownDocument.Parse("# Title\n\nParagraph\n");

        var result = document.ToAsciiDoc(new MarkdownAsciiDocExportOptions { LineEnding = "\r\n" }, out OfficeIMO.OfficeConversionReport conversionReport6);
        var resultReport = (MarkdownToAsciiDocReport)conversionReport6;

        Assert.DoesNotContain("\n", resultReport.GeneratedSource.Replace("\r\n", string.Empty), StringComparison.Ordinal);
        Assert.Equal(resultReport.GeneratedSource, result.ToAsciiDoc());
    }

    [Fact]
    public void CodeContainingListingFence_UsesALongerDelimiterWithoutChangingContent() {
        MarkdownDocument document = MarkdownDocument.Create().Code("text", "before\n----\nafter");

        var result = document.ToAsciiDoc(null, out OfficeIMO.OfficeConversionReport conversionReport7);
        var resultReport = (MarkdownToAsciiDocReport)conversionReport7;

        Assert.Contains("-----\nbefore\n----\nafter\n-----", resultReport.GeneratedSource, StringComparison.Ordinal);
        AsciiDocDelimitedBlock block = Assert.Single(result.BlocksOfType<AsciiDocDelimitedBlock>());
        Assert.Equal("before\n----\nafter\n", block.Content);
        Assert.DoesNotContain(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Feature == "code-delimiter");
    }

    [Fact]
    public void TableColumnCount_IncludesLogicalColumnSpans() {
        TableBlock table = Assert.Single(MarkdownDocument.Parse("| H |\n| --- |\n| wide |\n").Blocks.OfType<TableBlock>());
        table.GetCell(0, 0)!.ColumnSpan = 2;

        using var addedDocument = MarkdownDocument.Create();
        addedDocument.Add(table);
        var result = addedDocument.ToAsciiDoc(null, out OfficeIMO.OfficeConversionReport conversionReport8);
        var resultReport = (MarkdownToAsciiDocReport)conversionReport8;

        Assert.Contains("[cols=2*", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Equal(2, Assert.Single(result.BlocksOfType<AsciiDocTableBlock>()).Table.ColumnCount);
    }
}
