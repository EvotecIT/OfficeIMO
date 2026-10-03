namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocAuthoringTests {
    [Fact]
    public void AddParagraphRejectsMultipleOrEmptyParagraphsWithoutChangingTheDocument() {
        AsciiDocDocument document = AsciiDocDocument.Create().AddParagraph("Original");
        string original = document.ToAsciiDoc();
        Assert.Throws<ArgumentException>(() => document.AddParagraph("First\n\nSecond"));
        Assert.Throws<ArgumentException>(() => document.AddParagraph("\n\n"));
        Assert.Equal(original, document.ToAsciiDoc());
        document.AddParagraph("Wrapped\ntext\n\n");
        Assert.Equal(2, AsciiDocDocument.Parse(document.ToAsciiDoc()).BlocksOfType<AsciiDocParagraph>().Count());
    }
    [Theory]
    [InlineData("= Literal")]
    [InlineData("Term:: definition")]
    [InlineData("include::secret[]")]
    [InlineData("----")]
    public void EditedParagraphKeepsItsBlockKindAfterReopening(string text) {
        AsciiDocDocument document = AsciiDocDocument.Parse("Original\n");
        document.BlocksOfType<AsciiDocParagraph>().Single().Text = text;
        AsciiDocDocument reopened = AsciiDocDocument.Parse(document.ToAsciiDoc());
        Assert.Single(reopened.BlocksOfType<AsciiDocParagraph>());
        Assert.Empty(reopened.BlocksOfType<AsciiDocHeading>());
        Assert.Empty(reopened.BlocksOfType<AsciiDocDescriptionListBlock>());
    }
    [Fact]
    public void CreatedDocumentPreservesDistinctParagraphAndHeadingBlocks() {
        AsciiDocDocument document = AsciiDocDocument.Create().AddHeading(1, "Section").AddParagraph("= Literal").AddParagraph("Next");
        AsciiDocDocument reopened = AsciiDocDocument.Parse(document.ToAsciiDoc());
        Assert.True(document.IsModified);
        Assert.Single(reopened.BlocksOfType<AsciiDocHeading>());
        Assert.Equal(2, reopened.BlocksOfType<AsciiDocParagraph>().Count());
        Assert.Equal(1, reopened.BlocksOfType<AsciiDocHeading>().Single().SectionLevel);
    }

    [Fact]
    public void MovingAndRemovingBlockKeepsBoundMetadataTogether() {
        AsciiDocDocument document = AsciiDocDocument.Parse(".Caption\n[.wide]\n[[target]]\nFirst\n\nSecond\n");
        AsciiDocParagraph first = document.BlocksOfType<AsciiDocParagraph>().First();
        document.Move(first, document.Blocks.Count);
        AsciiDocDocument reopened = AsciiDocDocument.Parse(document.ToAsciiDoc());
        Assert.Equal(new[] { "Second", "First" }, reopened.BlocksOfType<AsciiDocParagraph>().Select(block => block.Text));
        Assert.Equal("Caption", reopened.BlocksOfType<AsciiDocParagraph>().Last().BlockTitle!.Title);
        Assert.Equal("target", reopened.BlocksOfType<AsciiDocParagraph>().Last().BlockAnchor!.Id);
        Assert.True(document.Remove(first));
        Assert.False(document.Remove(first));
        Assert.Equal("Second", AsciiDocDocument.Parse(document.ToAsciiDoc()).BlocksOfType<AsciiDocParagraph>().Single().Text);
        Assert.Empty(document.Blocks.OfType<AsciiDocBlockTitle>());
        Assert.Empty(document.Blocks.OfType<AsciiDocBlockAnchor>());
    }

    [Fact]
    public void RemovingListRemovesItsContinuationAttachments() {
        AsciiDocDocument document = AsciiDocDocument.Parse("* item\n+\n----\nattached\n----\n\nKept\n");
        Assert.True(document.Remove(document.BlocksOfType<AsciiDocListBlock>().Single()));
        Assert.Empty(document.BlocksOfType<AsciiDocDelimitedBlock>());
        Assert.Empty(document.BlocksOfType<AsciiDocListContinuation>());
        Assert.Equal("Kept", AsciiDocDocument.Parse(document.ToAsciiDoc()).BlocksOfType<AsciiDocParagraph>().Single().Text);
    }

    [Fact]
    public void RemovingAttachmentDetachesTheParentListProjection() {
        AsciiDocDocument document = AsciiDocDocument.Parse("* item\n+\n----\nattached\n----\n");
        Assert.True(document.Remove(document.BlocksOfType<AsciiDocDelimitedBlock>().Single()));
        Assert.Empty(document.BlocksOfType<AsciiDocListBlock>().Single().Items.Single().AttachedBlocks);
        Assert.Empty(document.BlocksOfType<AsciiDocListContinuation>());
    }

    [Fact]
    public void InsertionCannotSplitMetadataOrImportMalformedRecovery() {
        AsciiDocDocument document = AsciiDocDocument.Parse(".Caption\nParagraph\n");
        Assert.Throws<ArgumentException>(() => document.Insert(1, "Inserted"));
        Assert.Throws<ArgumentException>(() => document.Add("----\nunterminated"));
        Assert.False(document.IsModified);
        Assert.Equal(".Caption\nParagraph\n", document.ToAsciiDoc());
    }

    [Fact]
    public void CompoundChildEditsReachTheSourcePreservingWriter() {
        AsciiDocDocument document = AsciiDocDocument.Parse("====\nOriginal\n====\n");
        AsciiDocDelimitedBlock compound = document.BlocksOfType<AsciiDocDelimitedBlock>().Single();
        compound.Body!.BlocksOfType<AsciiDocParagraph>().Single().Text = "Edited";
        compound.Body.AddParagraph("Added");
        Assert.True(document.IsModified);
        AsciiDocDocument reopenedBody = AsciiDocDocument.Parse(document.ToAsciiDoc()).BlocksOfType<AsciiDocDelimitedBlock>().Single().Body!;
        Assert.Equal(new[] { "Edited", "Added" }, reopenedBody.BlocksOfType<AsciiDocParagraph>().Select(block => block.Text));
    }
}
