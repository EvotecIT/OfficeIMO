using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordCellCommentTests {
    [Fact]
    public void Legacy_extension_states_and_reply_links_survive_assignment_of_missing_paragraph_identities() {
        using var bytes = new MemoryStream();
        using (var initial = WordDocument.Create()) {
            initial.AddParagraph("First target").AddComment("Old", "", "Resolved root");
            Assert.Single(initial.Comments).MarkResolved();
            initial.AddParagraph("Second target").AddComment("Old", "", "Open root");
            initial.Comments.Single(c => c.Text == "Open root").AddReply("Other", "", "Original reply");
            initial.AddTable(1, 1);
            initial.Save(bytes);
        }
        bytes.Position = 0;
        using (var legacy = WordprocessingDocument.Open(bytes, true)) {
            foreach (Comment comment in legacy.MainDocumentPart!.WordprocessingCommentsPart!.Comments!.Elements<Comment>())
                foreach (Paragraph paragraph in comment.Elements<Paragraph>()) paragraph.ParagraphId = null;
            legacy.MainDocumentPart.WordprocessingCommentsPart.Comments.Save();
        }
        bytes.Position = 0;
        using var document = WordDocument.Load(bytes);
        document.AddCellComments(new[] { new WordCellComment(document.Tables[0].Rows[0].Cells[0], "New", "", "New root") });
        Assert.True(document.Comments.Single(c => c.Text == "Resolved root").IsResolved);
        Assert.Equal("Original reply", Assert.Single(document.Comments.Single(c => c.Text == "Open root").Replies).Text);
        document.Comments.Single(c => c.Text == "New root").DeleteThread();
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        Assert.True(reopened.Comments.Single(c => c.Text == "Resolved root").IsResolved);
        Assert.Equal("Original reply", Assert.Single(reopened.Comments.Single(c => c.Text == "Open root").Replies).Text);
        saved.Position = 0;
        using var artifact = WordprocessingDocument.Open(saved, false);
        foreach (Comment comment in artifact.MainDocumentPart!.WordprocessingCommentsPart!.Comments!.Elements<Comment>()) {
            string? identifier = comment.Elements<Paragraph>().First().ParagraphId?.Value;
            Assert.False(string.IsNullOrEmpty(identifier));
            Assert.Single(artifact.MainDocumentPart.WordprocessingCommentsExPart!.CommentsEx!
                .Elements<DocumentFormat.OpenXml.Office2013.Word.CommentEx>(), c => c.ParaId?.Value == identifier);
        }
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(artifact));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Appending_comments_to_legacy_roots_preserves_thread_resolution_and_removal_isolation(bool useExistingParagraphApi) {
        using var legacyBytes = new MemoryStream();
        using (var initial = WordDocument.Create()) {
            initial.AddParagraph("Legacy target").AddComment("Old", "", "Legacy root");
            initial.AddTable(1, 1);
            initial.Save(legacyBytes);
        }
        legacyBytes.Position = 0;
        using (var legacy = WordprocessingDocument.Open(legacyBytes, true)) {
            var main = legacy.MainDocumentPart!;
            Assert.Single(main.WordprocessingCommentsPart!.Comments!.Elements<Comment>()).Elements<Paragraph>().Single().ParagraphId = null;
            main.WordprocessingCommentsPart.Comments.Save();
            main.DeletePart(main.WordprocessingCommentsExPart!);
        }
        legacyBytes.Position = 0;
        using var document = WordDocument.Load(legacyBytes);
        Assert.Null(Assert.Single(document.Comments).ParaId);
        if (useExistingParagraphApi) document.AddParagraph("New target").AddComment("New", "", "New root");
        else document.AddCellComments(new[] { new WordCellComment(document.Tables[0].Rows[0].Cells[0], "New", "", "New root") });
        WordComment oldRoot = document.Comments.Single(c => c.Text == "Legacy root");
        WordComment newRoot = document.Comments.Single(c => c.Text == "New root");
        Assert.NotNull(oldRoot.ParaId);
        Assert.NotEqual(oldRoot.ParaId, newRoot.ParaId);
        newRoot.AddReply("New", "", "New reply");
        oldRoot.MarkResolved();
        Assert.Empty(oldRoot.Replies);
        Assert.Equal("New reply", Assert.Single(newRoot.Replies).Text);
        Assert.NotEqual(true, newRoot.IsResolved);
        oldRoot.Remove();
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        WordComment retained = reopened.Comments.Single(c => c.Text == "New root");
        Assert.Equal("New reply", Assert.Single(retained.Replies).Text);
        Assert.NotEqual(true, retained.IsResolved);
        saved.Position = 0;
        using var artifact = WordprocessingDocument.Open(saved, false);
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(artifact));
    }

    [Fact]
    public void Cell_comment_batches_preserve_content_dates_anchors_and_existing_threads_when_reopened() {
        using var document = WordDocument.Create();
        WordTable first = document.AddTable(2, 2);
        WordTable second = document.AddTable(1, 1);
        WordTableCell textCell = first.Rows[0].Cells[1];
        textCell.AddParagraph("First", removeExistingParagraphs: true).Bold = true;
        textCell.AddParagraph("Last").ParagraphAlignment = WordParagraphAlignment.Right;
        WordTableCell emptyCell = first.Rows[1].Cells[0];
        emptyCell.Paragraphs[0].ParagraphAlignment = WordParagraphAlignment.Center;
        var date = new DateTime(2001, 1, 1, 0, 0, 42, DateTimeKind.Utc);
        string text = "  Review\n\tthis 😀 & <literal>  ";
        IReadOnlyList<WordComment> roots = document.AddCellComments(new[] {
            new WordCellComment(textCell, " Reviewer ", "R", text, date),
            new WordCellComment(emptyCell, string.Empty, string.Empty, string.Empty, date)
        });
        roots[0].AddReply("Other", "O", "Reply");
        roots[0].MarkResolved();
        document.AddCellComments(new[] { new WordCellComment(second.Rows[0].Cells[0], "Other", "O", "Second table") });

        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        WordComment original = reopened.Comments.Single(c => c.Id == roots[0].Id);
        Assert.Equal(text, original.Text);
        Assert.Equal(" Reviewer ", original.Author);
        Assert.Equal(date, original.DateTime);
        Assert.True(original.IsResolved);
        Assert.Equal("Reply", Assert.Single(original.Replies).Text);
        Assert.Equal(string.Empty, reopened.Comments.Single(c => c.Id == roots[1].Id).Text);
        Assert.Null(reopened.Comments.Single(c => c.Text == "Second table").DateTime);
        Assert.Equal(new[] { "First", "Last" }, reopened.Tables[0].Rows[0].Cells[1].Paragraphs.Select(p => p.Text).Where(text => text.Length > 0));
        Assert.Equal(string.Empty, reopened.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text);

        saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(package));
        TableCell[] cells = package.MainDocumentPart!.Document.Body!.Descendants<TableCell>().ToArray();
        CheckAnchor(cells[1], roots[0].Id!);
        Assert.Equal(2, cells[1].Elements<Paragraph>().Count());
        Assert.Equal(new[] { "First", "Last" }, cells[1].Elements<Paragraph>().Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))));
        CheckAnchor(cells[2], roots[1].Id!);
        Assert.Empty(cells[0].Descendants<CommentReference>());
        Assert.Empty(cells[3].Descendants<CommentReference>());
        Assert.Equal(4, package.MainDocumentPart.WordprocessingCommentsPart!.Comments!.Elements<Comment>()
            .Select(c => c.Id!.Value).Distinct().Count());
    }

    [Theory]
    [InlineData("line\rbreak")]
    [InlineData("layout\u2028break")]
    [InlineData("invalid\0xml")]
    public void An_invalid_later_comment_does_not_apply_an_earlier_valid_comment(string invalidText) {
        using var document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        Assert.Throws<ArgumentException>(() => document.AddCellComments(new[] {
            new WordCellComment(cell, "Author", "", "Valid"),
            new WordCellComment(cell, "Author", "", invalidText)
        }));
        Assert.Empty(document.Comments);
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        Assert.Null(package.MainDocumentPart!.WordprocessingCommentsPart);
        Assert.Empty(package.MainDocumentPart.Document.Descendants<CommentReference>());
    }

    [Fact]
    public void A_batch_rejects_foreign_and_detached_cells_and_observes_enumeration_cancellation_before_mutation() {
        using var document = WordDocument.Create();
        using var other = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        WordTableCell foreign = other.AddTable(1, 1).Rows[0].Cells[0];
        Assert.Throws<ArgumentException>(() => document.AddCellComments(new[] { new WordCellComment(foreign, "A", "", "Text") }));
        WordTableCell detached = document.CreateTable(1, 1).Rows[0].Cells[0];
        Assert.Throws<ArgumentException>(() => document.AddCellComments(new[] { new WordCellComment(detached, "A", "", "Text") }));
        using var cancellation = new CancellationTokenSource();
        Assert.Throws<OperationCanceledException>(() => document.AddCellComments(Definitions(), cancellation.Token));
        Assert.Empty(document.Comments);

        IEnumerable<WordCellComment> Definitions() {
            yield return new WordCellComment(cell, "A", "", "First");
            cancellation.Cancel();
            yield return new WordCellComment(cell, "A", "", "Second");
        }
    }

    [Fact]
    public void Batch_identity_allocation_keeps_large_cell_sets_distinct_across_successive_batches() {
        using var document = WordDocument.Create();
        WordTable table = document.AddTable(64, 16);
        WordCellComment[] definitions = table.Rows.SelectMany(r => r.Cells)
            .Select((cell, index) => new WordCellComment(cell, "A", "", index.ToString())).ToArray();
        IReadOnlyList<WordComment> first = document.AddCellComments(definitions);
        IReadOnlyList<WordComment> second = document.AddCellComments(definitions);
        Assert.Equal(2048, first.Concat(second).Select(c => c.Id).Distinct().Count());
        Assert.Equal(2048, first.Concat(second).Select(c => c.ParaId).Distinct().Count());
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        foreach (TableCell cell in package.MainDocumentPart!.Document.Descendants<TableCell>()) {
            Assert.Equal(2, cell.Descendants<CommentReference>().Count());
        }
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(package));
    }

    [Fact]
    public void Whole_table_comments_keep_paragraph_properties_before_anchor_markers() {
        using var document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        table.Rows[0].Cells[0].Paragraphs[0].ParagraphAlignment = WordParagraphAlignment.Center;
        table.AddComment("A", "", "Whole table");
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(package));
    }

    private static void CheckAnchor(TableCell cell, string identifier) {
        Assert.Equal(identifier, Assert.Single(cell.Descendants<CommentRangeStart>()).Id?.Value);
        Assert.Equal(identifier, Assert.Single(cell.Descendants<CommentRangeEnd>()).Id?.Value);
        Assert.Equal(identifier, Assert.Single(cell.Descendants<CommentReference>()).Id?.Value);
    }
}
