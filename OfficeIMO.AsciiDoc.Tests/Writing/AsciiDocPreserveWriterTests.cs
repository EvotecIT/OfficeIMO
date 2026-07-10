namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocPreserveWriterTests {
    [Fact]
    public void EditingHeading_RewritesOnlyThatSourceBlock() {
        const string source = "= Original\r\n:toc: left\n\nParagraph\r\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocHeading heading = Assert.Single(document.BlocksOfType<AsciiDocHeading>());

        heading.Title = "Updated";

        Assert.True(document.IsModified);
        Assert.Equal("= Updated\r\n:toc: left\n\nParagraph\r\n", document.ToAsciiDoc());
        Assert.Equal(":toc: left\n\nParagraph\r\n", document.ToAsciiDoc().Substring("= Updated\r\n".Length));
    }

    [Fact]
    public void AssigningExistingValue_DoesNotDirtyDocument() {
        AsciiDocDocument document = AsciiDocDocument.Parse("= Same\n").Document;
        AsciiDocHeading heading = Assert.Single(document.BlocksOfType<AsciiDocHeading>());

        heading.Title = "Same";

        Assert.False(document.IsModified);
        Assert.Equal("= Same\n", document.ToAsciiDoc());
    }

    [Fact]
    public void EditingParagraph_PreservesSurroundingTrivia() {
        const string source = "== Section\r\n\r\nOld text\r\n\r\n// tail\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocParagraph paragraph = Assert.Single(document.BlocksOfType<AsciiDocParagraph>());

        paragraph.Text = "New first line\nNew second line";

        Assert.Equal("== Section\r\n\r\nNew first line\r\nNew second line\r\n\r\n// tail\n", document.ToAsciiDoc());
    }

    [Fact]
    public void EditingNestedListItem_PreservesOtherItemSpellingsAndLineEndings() {
        const string source = "- hyphen item\r\n** old nested\n* untouched\r\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocListBlock list = Assert.Single(document.BlocksOfType<AsciiDocListBlock>());

        list.Items[1].Text = "new nested";

        Assert.Equal("- hyphen item\r\n** new nested\n* untouched\r\n", document.ToAsciiDoc());
    }

    [Fact]
    public void EditingAttributeAndUnknownMacro_UsesTypedValues() {
        const string source = ":toc: left\r\nwidget::old-target[role=old]\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocAttributeEntry attribute = Assert.Single(document.BlocksOfType<AsciiDocAttributeEntry>());
        AsciiDocBlockMacro macro = Assert.Single(document.BlocksOfType<AsciiDocBlockMacro>());

        attribute.Value = "right";
        macro.Target = "new-target";
        macro.AttributeList = "role=new";

        Assert.Equal(":toc: right\r\nwidget::new-target[role=new]\n", document.ToAsciiDoc());
    }

    [Fact]
    public void EditingDelimitedContent_RetainsOriginalDelimiterLines() {
        const string source = "----\r\nold\n----\r\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocDelimitedBlock block = Assert.Single(document.BlocksOfType<AsciiDocDelimitedBlock>());

        block.Content = "new";

        Assert.Equal("----\r\nnew\r\n----\r\n", document.ToAsciiDoc());
    }

    [Fact]
    public void CanonicalMode_NormalizesRecognizedMarkersAndLineEndings() {
        const string source = "= Title\r\n\r\n- item\ncustom::target[]";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;

        string canonical = document.ToAsciiDoc(new AsciiDocWriterOptions {
            Mode = AsciiDocWriterMode.Canonical,
            LineEnding = "\n"
        });

        Assert.Equal("= Title\n\n* item\ncustom::target[]", canonical);
    }

    [Fact]
    public void SingleLineSemanticProperties_RejectEmbeddedNewlines() {
        AsciiDocDocument document = AsciiDocDocument.Parse("= Title\nmacro::target[]\n").Document;

        Assert.Throws<ArgumentException>(() => document.BlocksOfType<AsciiDocHeading>().Single().Title = "Bad\nTitle");
        Assert.Throws<ArgumentException>(() => document.BlocksOfType<AsciiDocBlockMacro>().Single().Target = "Bad\rTarget");
    }

    [Fact]
    public void AttributeName_RejectsSyntaxMarkersAndWhitespace() {
        AsciiDocAttributeEntry attribute = Assert.Single(AsciiDocDocument.Parse(":toc:\n").Document.BlocksOfType<AsciiDocAttributeEntry>());

        Assert.Throws<ArgumentException>(() => attribute.Name = "bad name");
        Assert.Throws<ArgumentException>(() => attribute.Name = "bad:name");
        Assert.Throws<ArgumentException>(() => attribute.Name = "bad!");
    }
}
