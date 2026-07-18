namespace OfficeIMO.AsciiDoc.Markdown.Tests;

public sealed class AsciiDocToMarkdownContractTests {
    [Fact]
    public void RepresentativeDocument_ConvertsThroughTypedMarkdownNodes() {
        const string source =
            "= Guide\n" +
            ":toc: left\n\n" +
            "== Start\n" +
            "Paragraph text\n\n" +
            "* one\n** nested\n\n" +
            "----\ncode();\n----\n" +
            "image::diagram.png[Architecture]\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;

        AsciiDocToMarkdownResult result = document.ToMarkdownDocumentResult();

        Assert.NotNull(result.Value.DocumentHeader);
        Assert.Equal(2, result.Value.Blocks.OfType<HeadingBlock>().Count());
        Assert.Single(result.Value.Blocks.OfType<ParagraphBlock>());
        UnorderedListBlock list = Assert.Single(result.Value.Blocks.OfType<UnorderedListBlock>());
        Assert.Equal(2, list.Items.Count);
        Assert.Equal(1, list.Items[1].Level);
        Assert.Single(result.Value.Blocks.OfType<CodeBlock>());
        Assert.Single(result.Value.Blocks.OfType<ImageBlock>());
        Assert.Empty(result.Report.Diagnostics);
        Assert.Equal(source, document.ToAsciiDoc());
    }

    [Fact]
    public void UnsupportedBlockMacro_IsSourceFallbackWithLocatedDiagnostic() {
        const string source = "diagram::architecture[format=svg]\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;

        AsciiDocToMarkdownResult result = document.ToMarkdownDocumentResult();

        CodeBlock fallback = Assert.Single(result.Value.Blocks.OfType<CodeBlock>());
        Assert.Equal("asciidoc", fallback.Language);
        Assert.Contains("diagram::architecture", fallback.Content);
        AsciiDocMarkdownConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics);
        Assert.Equal(AsciiDocMarkdownConversionOutcome.SourceFallback, diagnostic.Outcome);
        Assert.Equal(0, diagnostic.SourceSpan.Start.Offset);
        Assert.True(result.HasLoss);
    }

    [Fact]
    public void TableSource_ConvertsToTypedMarkdownTable() {
        AsciiDocDocument document = AsciiDocDocument.Parse("|===\n|a |b\n|===\n").Document;

        AsciiDocToMarkdownResult result = document.ToMarkdownDocumentResult(new AsciiDocToMarkdownOptions {
            PreserveUnsupportedAsSource = false
        });

        TableBlock table = Assert.Single(result.Value.Blocks.OfType<TableBlock>());
        Assert.Equal(2, table.RowCells[0].Count);
        Assert.Empty(result.Report.Diagnostics);
    }

    [Fact]
    public void DocumentTitleSectionsAndTableTitle_PreserveHierarchyAndCaptionMetadata() {
        const string source = "= Guide\n\n== Start\n\n.Important values\n|===\n|A |B\n|===\n";

        AsciiDocToMarkdownResult result = AsciiDocDocument.Parse(source).Document.ToMarkdownDocumentResult();

        Assert.Equal(new[] { 1, 2 }, result.Value.Blocks.OfType<HeadingBlock>().Select(static heading => heading.Level));
        TableBlock table = Assert.Single(result.Value.Blocks.OfType<TableBlock>());
        Assert.Equal("Important values", table.Attributes.GetAttribute("caption"));
        Assert.Contains("caption=\"Important values\"", result.Value.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void Comments_AreOmittedByDefaultWithExplicitDiagnostic() {
        AsciiDocDocument document = AsciiDocDocument.Parse("// internal note\nVisible\n").Document;

        AsciiDocToMarkdownResult result = document.ToMarkdownDocumentResult();

        Assert.Single(result.Value.Blocks.OfType<ParagraphBlock>());
        Assert.Equal("ADOCMD031", Assert.Single(result.Report.Diagnostics).Code);
    }

    [Fact]
    public void TypedMarkdownProjection_UsesExistingWordBridge() {
        const string source = "= Guide\n\n== Start\nParagraph text\n\n* one\n* two\n";
        AsciiDocToMarkdownResult conversion = AsciiDocDocument.Parse(source).Document.ToMarkdownDocumentResult();

        using var word = conversion.Value.ToWordDocument();
        string visibleText = string.Join(" ", word.Paragraphs.Select(paragraph => paragraph.Text));

        Assert.Contains("Guide", visibleText, StringComparison.Ordinal);
        Assert.Contains("Start", visibleText, StringComparison.Ordinal);
        Assert.Contains("Paragraph text", visibleText, StringComparison.Ordinal);
        Assert.Contains("one", visibleText, StringComparison.Ordinal);
    }

    [Fact]
    public void DirectPdfAdapter_PreservesProjectionDiagnosticsAndProducesPdf() {
        const string source = "= Guide\n\n== Start\nParagraph with stem:[x^2].\n\n|===\n|Name |Value\n|===\n";
        var result = AsciiDocDocument.Parse(source).Document.ToPdfDocumentResult();
        byte[] bytes = result.ToBytes();

        Assert.True(bytes.Length > 100);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(bytes, 0, 5));
        Assert.Contains(result.Warnings, warning =>
            warning.Converter == "OfficeIMO.AsciiDoc.Pdf" &&
            warning.Code == "ADOCMD103" &&
            warning.Details["stage"] == "semantic-projection");
    }

    [Fact]
    public void InlineSemantics_AttributesLinksImagesAndFormattingUseTypedMarkdownNodes() {
        const string source =
            ":product: OfficeIMO\n" +
            "Use *{product}* with _care_, `code`, <<intro,Introduction>>, image:icon.svg[Icon], and stem:[x^2].\n";

        AsciiDocToMarkdownResult result = AsciiDocDocument.Parse(source).Document.ToMarkdownDocumentResult();
        ParagraphBlock paragraph = Assert.Single(result.Value.Blocks.OfType<ParagraphBlock>());

        Assert.Contains(paragraph.Inlines.Nodes, node => node is BoldSequenceInline);
        Assert.Contains(paragraph.Inlines.Nodes, node => node is ItalicSequenceInline);
        Assert.Equal(2, paragraph.Inlines.Nodes.OfType<CodeSpanInline>().Count());
        Assert.Contains(paragraph.Inlines.Nodes, node => node is LinkInline);
        Assert.Contains(paragraph.Inlines.Nodes, node => node is ImageInline);
        Assert.Contains("OfficeIMO", paragraph.Inlines.Nodes.OfType<BoldSequenceInline>().Single().Inlines.Nodes.OfType<TextRun>().Single().Text);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "ADOCMD103");
    }

    [Fact]
    public void Phase1Blocks_ConvertToDefinitionsCalloutsCompoundListsAndStructuredTables() {
        const string source =
            "Term:: Definition\n\n" +
            "WARNING: Be careful\n\n" +
            "* item\n+\ncontinued\n\n" +
            "[cols=2*,%header]\n" +
            "|===\n|Name |Value\n2+|spanning\n|===\n";

        AsciiDocToMarkdownResult result = AsciiDocDocument.Parse(source).Document.ToMarkdownDocumentResult();

        Assert.Single(result.Value.Blocks.OfType<DefinitionListBlock>());
        Assert.Single(result.Value.Blocks.OfType<CalloutBlock>());
        UnorderedListBlock list = Assert.Single(result.Value.Blocks.OfType<UnorderedListBlock>());
        Assert.Single(list.Items[0].NestedBlocks.OfType<ParagraphBlock>());
        Assert.DoesNotContain(result.Value.Blocks.OfType<ParagraphBlock>(), paragraph =>
            paragraph.Inlines.Nodes.OfType<TextRun>().Any(text => text.Text == "continued"));
        TableBlock table = Assert.Single(result.Value.Blocks.OfType<TableBlock>());
        Assert.Equal(2, table.HeaderCells.Count);
        Assert.Equal(2, table.GetCell(0, 0)!.ColumnSpan);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "ADOCMD041");
    }

    [Fact]
    public void BlockMetadata_MapsToMarkdownAttributesCaptionAndCodeLanguage() {
        const string source = ".Example\n[[sample]]\n[source,csharp,.wide]\n----\ncode\n----\n";

        CodeBlock code = Assert.Single(AsciiDocDocument.Parse(source).Document.ToMarkdownDocumentResult().Value.Blocks.OfType<CodeBlock>());

        Assert.Equal("csharp", code.Language);
        Assert.Equal("Example", code.Caption);
        Assert.Equal("sample", code.Attributes.ElementId);
        Assert.Contains("wide", code.Attributes.Classes);
    }
}
