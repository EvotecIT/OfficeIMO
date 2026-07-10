namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocBlockAttributeListTests {
    [Fact]
    public void AttributeList_IsLosslessTypedAndBoundToFollowingBlock() {
        const string source = "[source,java,#sample,.wide,%nowrap,linenums=true,title=\"A, B\"]\r\n----\r\ncode\r\n----\r\n";

        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;

        AsciiDocBlockAttributeList metadata = Assert.Single(document.BlocksOfType<AsciiDocBlockAttributeList>());
        AsciiDocDelimitedBlock target = Assert.Single(document.BlocksOfType<AsciiDocDelimitedBlock>());
        Assert.Same(target, metadata.Target);
        Assert.Same(metadata, Assert.Single(target.AttributeLists));
        Assert.Equal("source", metadata.Attributes.Style);
        Assert.Equal("sample", metadata.Attributes.Id);
        Assert.Equal(new[] { "wide" }, metadata.Attributes.Roles);
        Assert.Equal(new[] { "nowrap" }, metadata.Attributes.Options);
        Assert.Equal("true", metadata.Attributes.GetNamedValue("LINENUMS"));
        Assert.Equal("A, B", metadata.Attributes.GetNamedValue("title"));
        Assert.Equal(source, document.ToAsciiDoc());
        Assert.True(document.SyntaxTree.IsLossless);
    }

    [Fact]
    public void BlankLine_PreventsMetadataBinding() {
        const string source = "[source]\n\n----\ncode\n----\n";

        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;

        AsciiDocBlockAttributeList metadata = Assert.Single(document.BlocksOfType<AsciiDocBlockAttributeList>());
        AsciiDocDelimitedBlock target = Assert.Single(document.BlocksOfType<AsciiDocDelimitedBlock>());
        Assert.Null(metadata.Target);
        Assert.Empty(target.AttributeLists);
        Assert.Equal(source, document.ToAsciiDoc());
    }

    [Fact]
    public void EditingContent_ReparsesSemanticAttributesAndKeepsFollowingSource() {
        const string source = "[source,java]\r\n----\r\ncode\r\n----\r\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;
        AsciiDocBlockAttributeList metadata = Assert.Single(document.BlocksOfType<AsciiDocBlockAttributeList>());

        metadata.Content = "source,csharp,#sample";

        Assert.Equal("source", metadata.Attributes.Style);
        Assert.Equal("sample", metadata.Attributes.Id);
        Assert.Equal("[source,csharp,#sample]\r\n----\r\ncode\r\n----\r\n", document.ToAsciiDoc());
    }
}
