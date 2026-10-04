namespace OfficeIMO.AsciiDoc.Markdown.Tests;

public sealed class AsciiDocReferencesAndCellsTests {
    [Fact]
    public void AdjacentFootnotesAndQuotedProseReachMarkdownWithoutLoss() {
        const string source = "doublefootnote:[Don't forget *this*.] and link:guide.pdf[Don't miss it].";
        AsciiDocToMarkdownResult conversion = AsciiDocDocument.Parse(source).ToMarkdownDocumentResult();

        Assert.Single(conversion.Value.Blocks.OfType<FootnoteDefinitionBlock>());
        string markdown = conversion.Value.ToMarkdown();
        Assert.Contains("double[^", markdown);
        Assert.Contains("Don't forget **this**.", markdown);
        Assert.Contains("[Don't miss it](guide.pdf)", markdown);
        Assert.False(conversion.Report.HasLoss);
    }

    [Fact]
    public void ReplacingTableContentRefreshesTheTypedModelAndConversion() {
        var document = AsciiDocDocument.Parse("[cols=\"2\"]\n|===\n|Old |Value\n|===\n");
        var block = document.BlocksOfType<AsciiDocTableBlock>().Single();
        block.Content = "|New |Cells\n";
        Assert.Equal(new[] { "New", "Cells" }, block.Table.Cells.Select(cell => cell.Value));
        block.Table.Cells[0].Value = "Edited";
        Assert.Contains("Edited", block.Content);
        Assert.Contains("Edited", document.ToAsciiDoc());
        string markdown = document.ToMarkdownDocumentResult().Value.ToMarkdown();
        Assert.Contains("Edited", markdown);
        Assert.DoesNotContain("Old", markdown);
    }

    [Fact]
    public void InlineCellFootnotesReferencesAndEditsShareTheNativeModel() {
        var document = AsciiDocDocument.Parse("[[section]]\n== A section\n\n[cols=\"1\"]\n|===\n|See xref:section[] footnote:note[*Cell note*].\n|===\n");
        var cell = document.BlocksOfType<AsciiDocTableBlock>().Single().Table.Cells.Single();
        var note = cell.Inlines!.Items.OfType<AsciiDocFootnoteInline>().Single();
        note.AttributeList = "*Edited note*";
        Assert.Contains("Edited note", document.ToAsciiDoc());
        var converted = document.ToMarkdownDocumentResult();
        Assert.Single(converted.Value.Blocks.OfType<FootnoteDefinitionBlock>());
        Assert.Contains("**Edited note**", converted.Value.ToMarkdown());
        Assert.Contains("[A section](#section)", converted.Value.ToMarkdown());
        Assert.DoesNotContain(converted.Report.Diagnostics, issue => issue.Code == "ADOCREF003" || issue.Code == "ADOCMD104");
    }

    [Fact]
    public void ConvertingOneBlockUsesInheritedAttributesAndOnlyItsReferencedDefinitions() {
        var document = AsciiDocDocument.Parse(":value: inherited\n\nFirst footnote:first[{value}].\n\nSecond footnote:second[Elsewhere].\n");
        var paragraph = document.BlocksOfType<AsciiDocParagraph>().First();
        var attributes = document.GetBlockContexts().First(context => context.Block == paragraph).Attributes;
        var converted = paragraph.ToMarkdownDocumentResult(attributes);
        Assert.Contains("inherited", converted.Value.ToMarkdown());
        converted = paragraph.ToMarkdownDocumentResult(attributes, new AsciiDocToMarkdownOptions { References = AsciiDocReferenceCatalog.Create(document) });
        Assert.Single(converted.Value.Blocks.OfType<FootnoteDefinitionBlock>());
        Assert.DoesNotContain("Elsewhere", converted.Value.ToMarkdown());
    }

    [Theory]
    [InlineData("<4> Four\n<5> Five\n", "4. Four", "5. Five")]
    [InlineData("<4> Four\n<8> Eight\n", "\\(4\\) Four", "\\(8\\) Eight")]
    public void CalloutExplanationsKeepTheirNumbers(string source, string first, string second) {
        var converted = AsciiDocDocument.Parse(source).ToMarkdownDocumentResult();
        string markdown = converted.Value.ToMarkdown();
        Assert.Contains(first, markdown);
        Assert.Contains(second, markdown);
        Assert.Contains(converted.Report.Diagnostics, issue => issue.Code == "ADOCMD052");
    }
    [Fact]
    public void ColumnStylesAndRepetitionDriveRichCellBodies() {
        AsciiDocDocument document = AsciiDocDocument.Parse("[cols=\"2*a,1\"]\n|===\n|First\n\n* Item\n\n|Second\n\nParagraph\n\n|Plain\n|===\n");
        AsciiDocTable table = document.BlocksOfType<AsciiDocTableBlock>().Single().Table;
        Assert.Equal(3, table.ColumnCount);
        Assert.Equal(new[] { 'a', 'a', 'd' }, table.Cells.Select(cell => cell.Style));
        Assert.Single(table.Cells[0].Body!.BlocksOfType<AsciiDocListBlock>());
        Assert.Equal(2, table.Cells[1].Body!.BlocksOfType<AsciiDocParagraph>().Count());
        Assert.Null(table.Cells[2].Body);
        Assert.Contains(document.ToMarkdownDocumentResult().Value.Blocks.OfType<TableBlock>().Single().ChildBlocks, block => block is UnorderedListBlock);
    }

    [Fact]
    public void LabeledUrlsPreserveRichLabelsAndSourceOrderedAttributes() {
        const string source = ":host: example.org\n\nhttps://{host}/guide[*Read* the guide]\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source);
        Assert.Equal(source, document.ToAsciiDoc());
        AsciiDocToMarkdownResult conversion = document.ToMarkdownDocumentResult();
        Assert.Contains("[**Read** the guide](https://example.org/guide)", conversion.Value.ToMarkdown());
        Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ADOCMD102");
    }

    [Fact]
    public void AnonymousAndRepeatedNamedFootnotesKeepEditsAndDefinitionAttributeState() {
        const string source = ":version: one\n\nText footnote:shared[Version {version} with *emphasis*.] and footnote:[Anonymous.].\n\n:version: two\n\nAgain footnote:shared[].\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source);
        var note = document.BlocksOfType<AsciiDocParagraph>().First().Inlines.Items.OfType<AsciiDocFootnoteInline>().First();
        note.AttributeList = "Version {version} with *new emphasis*.";
        Assert.Contains("with *new emphasis*", document.ToAsciiDoc());
        AsciiDocToMarkdownResult conversion = document.ToMarkdownDocumentResult();
        Assert.Equal(2, conversion.Value.Blocks.OfType<FootnoteDefinitionBlock>().Count());
        string markdown = conversion.Value.ToMarkdown();
        Assert.Contains("Version one with **new emphasis**.", markdown);
        Assert.Contains("Anonymous.", markdown);
        Assert.DoesNotContain("Version two", markdown);
        Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Feature == "footnote");
    }

    [Fact]
    public void ReferenceCatalogBindsBibliographyLabelsAndReportsDuplicateOrDanglingTargets() {
        const string source = "See <<ref>> and <<missing>>.\n\n* [[[ref,12]]] First reference.\n* [[[ref,13]]] Duplicate.\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source);
        var catalog = AsciiDocReferenceCatalog.Create(document);
        Assert.True(catalog.Targets["ref"].IsBibliography);
        Assert.Equal("12", catalog.Targets["ref"].Label);
        Assert.Contains(catalog.Diagnostics, diagnostic => diagnostic.Code == "ADOCREF001");
        Assert.Contains(catalog.Diagnostics, diagnostic => diagnostic.Code == "ADOCREF003");
        MarkdownDoc converted = document.ToMarkdownDocumentResult().Value;
        string markdown = converted.ToMarkdown();
        Assert.Contains("(#ref)", markdown);
        Assert.Contains(converted.Blocks.OfType<ParagraphBlock>().First().Inlines.Nodes.OfType<LinkInline>(), link => link.Text == "[12]");
        Assert.Equal(source, document.ToAsciiDoc());
    }

    [Fact]
    public void AsciiDocTableCellsKeepMultipleBlocksAndTypedEdits() {
        const string source = "[cols=\"1,1\"]\n|===\na|First paragraph.\n\nSecond paragraph.\n\n* First item\n* Second item\n\n|Other cell\n|===\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source);
        AsciiDocTableCell cell = document.BlocksOfType<AsciiDocTableBlock>().Single().Table.Cells.First();
        Assert.Equal(2, cell.Body!.BlocksOfType<AsciiDocParagraph>().Count());
        cell.Body.BlocksOfType<AsciiDocParagraph>().Last().Text = "Changed paragraph.";
        Assert.Contains("Changed paragraph.", document.ToAsciiDoc());
        TableBlock table = document.ToMarkdownDocumentResult().Value.Blocks.OfType<TableBlock>().Single();
        Assert.Contains(table.ChildBlocks, block => block is UnorderedListBlock);
        Assert.Contains(table.ChildBlocks.OfType<ParagraphBlock>(), paragraph => ((IMarkdownBlock)paragraph).RenderMarkdown().Contains("Changed paragraph."));
        Assert.Contains("First item", ((IMarkdownBlock)table).RenderMarkdown());
    }
}
