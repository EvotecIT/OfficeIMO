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
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void LegacyDoc_CommentReferenceFollowsPrecedingRevisedNote(bool endnote, bool inserted) {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("MixedMarkerPrefix");
        if (endnote) paragraph.AddEndNote("Endnote body"); else paragraph.AddFootNote("Footnote body");
        Run note = paragraph._paragraph.Descendants<Run>().Single(run => endnote
            ? run.Elements<EndnoteReference>().Any() : run.Elements<FootnoteReference>().Any());
        OpenXmlCompositeElement revision = inserted
            ? new InsertedRun { Id = "125", Author = "Note author" }
            : new DeletedRun { Id = "125", Author = "Note author" };
        note.InsertBeforeSelf(revision);
        note.Remove();
        revision.Append(note);
        paragraph.AddComment("Comment author", "CA", "Following comment");
        Run comment = paragraph._paragraph.Descendants<Run>().Single(run => run.Elements<CommentReference>().Any());
        comment.Remove();
        paragraph._paragraph.Append(comment);
        Assert.Equal(new[] { "note", "comment" }, MarkerOrder(paragraph._paragraph));

        using var reopened = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        Paragraph imported = reopened._document!.Body!.Descendants<Paragraph>().First();
        Assert.Equal(new[] { "note", "comment" }, MarkerOrder(imported));
        Assert.Equal("Following comment", Assert.Single(reopened.Comments).Text);

        static string[] MarkerOrder(Paragraph value) => value.Descendants<Run>()
            .Where(run => run.Elements<FootnoteReference>().Any() || run.Elements<EndnoteReference>().Any() || run.Elements<CommentReference>().Any())
            .Select(run => run.Elements<CommentReference>().Any() ? "comment" : "note").ToArray();
    }
}
