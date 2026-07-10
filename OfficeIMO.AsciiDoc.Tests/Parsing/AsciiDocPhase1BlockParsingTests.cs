namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocPhase1BlockParsingTests {
    [Fact]
    public void MixedBlockMetadata_BindsWithoutDuplicatingWriterOwnership() {
        const string source = ".*Code* example\n[[sample,Sample code]]\n[source,csharp,.wide]\n----\nConsole.WriteLine();\n----\n";

        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocDelimitedBlock block = Assert.Single(document.BlocksOfType<AsciiDocDelimitedBlock>());

        Assert.Equal("source", block.Style);
        Assert.NotNull(block.BlockTitle);
        Assert.Equal("*Code* example", block.BlockTitle!.Title);
        Assert.Single(block.BlockTitle.Inlines.Items.OfType<AsciiDocFormattedInline>());
        Assert.NotNull(block.BlockAnchor);
        Assert.Equal("sample", block.BlockAnchor!.Id);
        Assert.Equal("Sample code", block.BlockAnchor.ReferenceText);
        Assert.Equal(new[] { "wide" }, Assert.Single(block.AttributeLists).Attributes.Roles);
        Assert.Equal(source, document.ToAsciiDoc());
        Assert.True(document.SyntaxTree.IsLossless);
    }

    [Theory]
    [InlineData("NOTE", AsciiDocAdmonitionKind.Note)]
    [InlineData("TIP", AsciiDocAdmonitionKind.Tip)]
    [InlineData("IMPORTANT", AsciiDocAdmonitionKind.Important)]
    [InlineData("WARNING", AsciiDocAdmonitionKind.Warning)]
    [InlineData("CAUTION", AsciiDocAdmonitionKind.Caution)]
    public void AdmonitionParagraphs_AreTypedAndEditable(string label, AsciiDocAdmonitionKind kind) {
        string source = label + ": Read *carefully*.\r\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocAdmonitionBlock admonition = Assert.Single(document.BlocksOfType<AsciiDocAdmonitionBlock>());

        Assert.Equal(kind, admonition.Kind);
        Assert.Single(admonition.Inlines.Items.OfType<AsciiDocFormattedInline>());
        Assert.Equal(source, document.ToAsciiDoc());

        admonition.Text = "Changed";
        Assert.Equal(label + ": Changed\r\n", document.ToAsciiDoc());
    }

    [Fact]
    public void DescriptionLists_RetainDepthInlinesAndEdits() {
        const string source = "*Term*:: Definition with {product}\nNested::: `value`\r\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocDescriptionListBlock list = Assert.Single(document.BlocksOfType<AsciiDocDescriptionListBlock>());

        Assert.Equal(2, list.Items.Count);
        Assert.Equal(1, list.Items[0].Depth);
        Assert.Equal(2, list.Items[1].Depth);
        Assert.Single(list.Items[0].TermInlines.Items.OfType<AsciiDocFormattedInline>());
        Assert.Single(list.Items[0].DescriptionInlines.Items.OfType<AsciiDocAttributeReferenceInline>());
        Assert.Single(list.Items[1].DescriptionInlines.Items.OfType<AsciiDocFormattedInline>());
        Assert.Equal(source, document.ToAsciiDoc());

        list.Items[0].Description = "Updated";
        Assert.Equal("*Term*:: Updated\nNested::: `value`\r\n", document.ToAsciiDoc());
    }

    [Fact]
    public void ListContinuations_BindMultipleCompoundBlocksToTheItem() {
        const string source =
            "* item\n" +
            "+\n" +
            "continued paragraph\n" +
            "+\n" +
            "[source,csharp]\n" +
            "----\ncode\n----\n";

        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocListItem item = Assert.Single(document.BlocksOfType<AsciiDocListBlock>()).Items.Single();
        AsciiDocListContinuation[] continuations = document.BlocksOfType<AsciiDocListContinuation>().ToArray();

        Assert.Equal(2, continuations.Length);
        Assert.All(continuations, continuation => Assert.Same(item, continuation.TargetItem));
        Assert.Equal(2, item.AttachedBlocks.Count);
        Assert.IsType<AsciiDocParagraph>(item.AttachedBlocks[0]);
        AsciiDocDelimitedBlock listing = Assert.IsType<AsciiDocDelimitedBlock>(item.AttachedBlocks[1]);
        Assert.Equal("source", listing.Style);
        Assert.Same(listing, continuations[1].AttachedBlock);
        Assert.Equal(source, document.ToAsciiDoc());
    }
}
