namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocLosslessParsingTests {
    [Fact]
    public void Phase0Document_RoundTripsEverySourceCharacter() {
        const string source =
            "= Product Guide\r\n" +
            ":toc: left\n" +
            ":sectnums:\r" +
            "\r\n" +
            "// retained comment\n" +
            "== Overview\r\n" +
            "First line\nsecond line\r\n" +
            "\r\n" +
            "* item one\n" +
            "** nested item\r\n" +
            ". first ordered\n" +
            ".. nested ordered\r\n" +
            "----\r\n" +
            "Console.WriteLine(\"hello\");\n" +
            "----\r\n" +
            "diagram::architecture[format=svg]\n";

        AsciiDocParseResult result = AsciiDocDocument.Parse(source);

        Assert.True(result.IsLossless);
        Assert.Equal(source, result.Document.ToAsciiDoc());
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "ADOC900");
        Assert.Contains(result.Document.Blocks, block => block is AsciiDocHeading heading && heading.IsDocumentTitle);
        Assert.Contains(result.Document.Blocks, block => block is AsciiDocAttributeEntry attribute && attribute.Name == "toc");
        Assert.Contains(result.Document.Blocks, block => block is AsciiDocLineComment);
        Assert.Contains(result.Document.Blocks, block => block is AsciiDocParagraph paragraph && paragraph.Text == "First line\nsecond line");
        Assert.Equal(2, result.Document.Blocks.OfType<AsciiDocListBlock>().Count());
        Assert.Contains(result.Document.Blocks, block => block is AsciiDocDelimitedBlock delimited && delimited.Kind == AsciiDocDelimitedBlockKind.Listing);
        AsciiDocBlockMacro macro = Assert.Single(result.Document.Blocks.OfType<AsciiDocBlockMacro>());
        Assert.Equal("diagram", macro.Name);
        Assert.False(macro.IsKnown);
    }

    [Fact]
    public void RootChildren_AreContiguousAndSliceBackToOriginalSource() {
        const string source = "= Title\n\nParagraph\r\n* item\ncustom::target[x=1]";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;

        Assert.True(document.SyntaxTree.IsLossless);
        Assert.Equal(source, string.Concat(document.SyntaxTree.Root.Children.Select(node => node.OriginalText)));

        int expectedOffset = 0;
        foreach (AsciiDocSyntaxNode node in document.SyntaxTree.Root.Children) {
            Assert.Equal(expectedOffset, node.Span.Start.Offset);
            Assert.Equal(node.OriginalText, node.Span.Slice(source));
            expectedOffset = node.Span.End.Offset;
        }

        Assert.Equal(source.Length, expectedOffset);
        Assert.All(document.SyntaxTree.Root.DescendantsAndSelf(), node => AssertNodeCoverage(source, node));
    }

    [Fact]
    public void TokenNodes_RetainMarkersAndLineEndings() {
        const string source = "== Section\r\n:name: value\nmacro::target[a=b]\r";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;

        AsciiDocSyntaxNode heading = document.SyntaxTree.Root.Children[0];
        Assert.Equal("==", Assert.Single(heading.Children, node => node.Kind == AsciiDocSyntaxKind.HeadingMarker).OriginalText);
        Assert.Equal("\r\n", Assert.Single(heading.Children, node => node.Kind == AsciiDocSyntaxKind.LineEnding).OriginalText);

        AsciiDocSyntaxNode attribute = document.SyntaxTree.Root.Children[1];
        Assert.Equal("name", Assert.Single(attribute.Children, node => node.Kind == AsciiDocSyntaxKind.AttributeName).OriginalText);
        Assert.Equal("value", Assert.Single(attribute.Children, node => node.Kind == AsciiDocSyntaxKind.AttributeValue).OriginalText);

        AsciiDocSyntaxNode macro = document.SyntaxTree.Root.Children[2];
        Assert.Equal("::", Assert.Single(macro.Children, node => node.Kind == AsciiDocSyntaxKind.MacroSeparator).OriginalText);
        Assert.Equal("[a=b]", Assert.Single(macro.Children, node => node.Kind == AsciiDocSyntaxKind.MacroAttributeList).OriginalText);
    }

    [Fact]
    public void UnterminatedDelimitedBlock_IsDiagnosedAndPreserved() {
        const string source = "----\ncode\r\nstill code";

        AsciiDocParseResult result = AsciiDocDocument.Parse(source);

        AsciiDocDelimitedBlock block = Assert.Single(result.Document.Blocks.OfType<AsciiDocDelimitedBlock>());
        Assert.False(block.IsTerminated);
        Assert.Equal("ADOC001", Assert.Single(result.Diagnostics).Code);
        Assert.True(result.HasErrors);
        Assert.True(result.IsLossless);
        Assert.Equal(source, result.Document.ToAsciiDoc());
    }

    [Fact]
    public void EmptySource_IsLossless() {
        AsciiDocParseResult result = AsciiDocDocument.Parse(string.Empty);

        Assert.True(result.IsLossless);
        Assert.Empty(result.Document.Blocks);
        Assert.Equal(string.Empty, result.Document.ToAsciiDoc());
        Assert.Equal(1, result.Document.Source.LineCount);
    }

    [Fact]
    public void SourceWithoutFinalLineEnding_RemainsWithoutOne() {
        const string source = "= Title\n\nLast paragraph";

        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;

        Assert.Equal(source, document.ToAsciiDoc());
        Assert.DoesNotContain("Last paragraph\n", document.ToAsciiDoc());
    }

    [Theory]
    [InlineData("----", AsciiDocDelimitedBlockKind.Listing)]
    [InlineData("....", AsciiDocDelimitedBlockKind.Literal)]
    [InlineData("====", AsciiDocDelimitedBlockKind.Example)]
    [InlineData("****", AsciiDocDelimitedBlockKind.Sidebar)]
    [InlineData("____", AsciiDocDelimitedBlockKind.Quote)]
    [InlineData("++++", AsciiDocDelimitedBlockKind.Passthrough)]
    [InlineData("--", AsciiDocDelimitedBlockKind.Open)]
    [InlineData("|===", AsciiDocDelimitedBlockKind.Table)]
    [InlineData("////", AsciiDocDelimitedBlockKind.Comment)]
    public void CommonDelimitedBlocks_AreTypedAndLossless(string delimiter, AsciiDocDelimitedBlockKind expectedKind) {
        string source = delimiter + "\ncontent\n" + delimiter + "\n";

        AsciiDocParseResult result = AsciiDocDocument.Parse(source);

        AsciiDocDelimitedBlock block = Assert.Single(result.Document.Blocks.OfType<AsciiDocDelimitedBlock>());
        Assert.Equal(expectedKind, block.Kind);
        Assert.True(block.IsTerminated);
        Assert.Equal(source, result.Document.ToAsciiDoc());
    }

    [Theory]
    [InlineData("-----", AsciiDocDelimitedBlockKind.Listing)]
    [InlineData("......", AsciiDocDelimitedBlockKind.Literal)]
    [InlineData("======", AsciiDocDelimitedBlockKind.Example)]
    [InlineData("******", AsciiDocDelimitedBlockKind.Sidebar)]
    [InlineData("______", AsciiDocDelimitedBlockKind.Quote)]
    [InlineData("++++++", AsciiDocDelimitedBlockKind.Passthrough)]
    [InlineData("//////", AsciiDocDelimitedBlockKind.Comment)]
    public void VariableLengthDelimitedBlocks_AreTypedAndLossless(string delimiter, AsciiDocDelimitedBlockKind expectedKind) {
        string source = delimiter + "\ncontent\n" + delimiter + "\n";

        AsciiDocParseResult result = AsciiDocDocument.Parse(source);

        AsciiDocDelimitedBlock block = Assert.Single(result.Document.Blocks.OfType<AsciiDocDelimitedBlock>());
        Assert.Equal(expectedKind, block.Kind);
        Assert.Equal(delimiter, block.Delimiter);
        Assert.True(block.IsTerminated);
        Assert.Equal(source, result.Document.ToAsciiDoc());
    }

    [Fact]
    public void HeadingLevels_DistinguishDocumentTitleFromSections() {
        const string source = "= Document\n== First\n=== Second\n";

        AsciiDocHeading[] headings = AsciiDocDocument.Parse(source).Document.BlocksOfType<AsciiDocHeading>().ToArray();

        Assert.True(headings[0].IsDocumentTitle);
        Assert.Equal(0, headings[0].SectionLevel);
        Assert.False(headings[1].IsDocumentTitle);
        Assert.Equal(1, headings[1].SectionLevel);
        Assert.Equal(2, headings[2].SectionLevel);
    }

    [Fact]
    public void AttributeUnsetForms_AreRecognizedWithoutChangingSource() {
        const string source = ":toc!:\n:!sectnums:\n";

        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocAttributeEntry[] entries = document.BlocksOfType<AsciiDocAttributeEntry>().ToArray();

        Assert.Equal(2, entries.Length);
        Assert.All(entries, entry => Assert.True(entry.IsUnset));
        Assert.Equal("toc", entries[0].Name);
        Assert.Equal("sectnums", entries[1].Name);
        Assert.Equal(source, document.ToAsciiDoc());
    }

    private static void AssertNodeCoverage(string source, AsciiDocSyntaxNode node) {
        Assert.Equal(node.OriginalText, node.Span.Slice(source));
        if (node.Children.Count == 0) return;

        int expectedOffset = node.Span.Start.Offset;
        foreach (AsciiDocSyntaxNode child in node.Children) {
            Assert.Equal(expectedOffset, child.Span.Start.Offset);
            expectedOffset = child.Span.End.Offset;
        }
        Assert.Equal(node.Span.End.Offset, expectedOffset);
    }
}
