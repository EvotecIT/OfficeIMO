using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void LegacyDoc_CommentReferencesRetainFormattingAndRevisionIdentity(int revisionKind) {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("RevisedCommentMarker");
        paragraph.AddComment("Comment author", "CA", "Comment text");
        Run reference = paragraph._paragraph.Descendants<Run>().Single(run => run.Elements<CommentReference>().Any());
        reference.RunProperties = new RunProperties(
            new FontSize { Val = "28" }, new FontSizeComplexScript { Val = "28" },
            new RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
            new Color { Val = "3355CC" });
        var date = new DateTime(2026, 10, 3, 10, 20, 0, DateTimeKind.Utc);
        if (revisionKind != 0) {
            OpenXmlCompositeElement revision = revisionKind == 1
                ? new InsertedRun { Id = "124", Author = "Reference author", Date = date }
                : new DeletedRun { Id = "124", Author = "Reference author", Date = date };
            reference.InsertBeforeSelf(revision);
            reference.Remove();
            revision.Append(reference);
        }
        using var reopened = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        Run imported = reopened._document!.Body!.Descendants<Run>().Single(run => run.Elements<CommentReference>().Any());
        Assert.Equal("28", imported.RunProperties?.FontSize?.Val?.Value);
        Assert.Equal("Arial", imported.RunProperties?.RunFonts?.Ascii?.Value);
        Assert.Equal("3355CC", imported.RunProperties?.Color?.Val?.Value);
        if (revisionKind == 1) {
            InsertedRun wrapper = Assert.IsType<InsertedRun>(imported.Parent);
            Assert.Equal("Reference author", wrapper.Author?.Value);
            Assert.Equal(date, wrapper.Date?.Value);
        } else if (revisionKind == 2) {
            DeletedRun wrapper = Assert.IsType<DeletedRun>(imported.Parent);
            Assert.Equal("Reference author", wrapper.Author?.Value);
            Assert.Equal(date, wrapper.Date?.Value);
        } else {
            Assert.IsType<Paragraph>(imported.Parent);
        }
        Assert.Equal("Comment text", Assert.Single(reopened.Comments).Text);
        reopened.AcceptRevisions();
        Assert.Equal(revisionKind == 2 ? 0 : 1, reopened._document.Body.Descendants<CommentReference>().Count());
    }
}
