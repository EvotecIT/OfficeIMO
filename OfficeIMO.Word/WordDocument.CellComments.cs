using System.Globalization;
using System.Threading;
using System.Xml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Adds root comments across body table cells, allocating identities and saving comment parts once per batch.</summary>
    /// <remarks>Targets must belong to this document's body. Invalid targets or text and pre-commit cancellation leave the batch unapplied.
    /// Existing cell content is retained. Replies and resolved state are managed through the returned comments.</remarks>
    public IReadOnlyList<WordComment> AddCellComments(IEnumerable<WordCellComment> comments,
        CancellationToken cancellationToken = default) {
        if (comments == null) throw new ArgumentNullException(nameof(comments));
        cancellationToken.ThrowIfCancellationRequested();
        var definitions = new List<WordCellComment>();
        foreach (WordCellComment definition in comments) {
            cancellationToken.ThrowIfCancellationRequested();
            if (definition == null) throw new ArgumentException("A cell comment cannot be null.", nameof(comments));
            if (!ReferenceEquals(definition.Cell.Document, this)
                || !definition.Cell._tableCell.Ancestors<Body>().Any(body => ReferenceEquals(body,
                    _wordprocessingDocument.MainDocumentPart?.Document?.Body)))
                throw new ArgumentException("A comment target must be an attached cell in this document's body.", nameof(comments));
            if (!WordCellComment.CanPreserveText(definition.Text))
                throw new ArgumentException("The comment text cannot be preserved by the plain-text paragraph owner.", nameof(comments));
            XmlConvert.VerifyXmlChars(definition.Author);
            XmlConvert.VerifyXmlChars(definition.Initials);
            definitions.Add(definition);
        }
        if (definitions.Count == 0) return Array.Empty<WordComment>();

        // Inspect existing parts without creating them, so preflight failure has no package side effects.
        var main = _wordprocessingDocument.MainDocumentPart!;
        Comments? existing = main.WordprocessingCommentsPart?.Comments;
        CommentsEx? existingEx = main.WordprocessingCommentsExPart?.CommentsEx;
        long maximumId = 0;
        uint maximumParagraphId = 0;
        foreach (Comment comment in existing?.Elements<Comment>() ?? Enumerable.Empty<Comment>()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!long.TryParse(comment.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long identifier)
                || identifier > int.MaxValue)
                throw new InvalidOperationException("An existing comment identifier is outside the supported integer range.");
            maximumId = Math.Max(maximumId, identifier);
            foreach (Paragraph paragraph in comment.Elements<Paragraph>()) ObserveParagraphId(paragraph.ParagraphId?.Value);
        }
        foreach (CommentEx comment in existingEx?.Elements<CommentEx>() ?? Enumerable.Empty<CommentEx>()) {
            cancellationToken.ThrowIfCancellationRequested();
            ObserveParagraphId(comment.ParaId?.Value);
        }
        if (maximumId > int.MaxValue - definitions.Count || (ulong)maximumParagraphId + (uint)definitions.Count > uint.MaxValue)
            throw new InvalidOperationException("The comment identity range is exhausted.");

        var prepared = new List<(WordComment Comment, Paragraph First, Paragraph Last)>(definitions.Count);
        for (int index = 0; index < definitions.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            WordCellComment definition = definitions[index];
            Paragraph[] paragraphs = definition.Cell._tableCell.Elements<Paragraph>().ToArray();
            if (paragraphs.Length == 0)
                throw new ArgumentException("A comment target must contain a direct paragraph.", nameof(comments));
            string id = (maximumId + index + 1).ToString(CultureInfo.InvariantCulture);
            string paragraphId = (maximumParagraphId + (uint)index + 1).ToString("X8", CultureInfo.InvariantCulture);
            prepared.Add((WordComment.CreateDetached(this, definition, id, paragraphId), paragraphs[0], paragraphs[paragraphs.Length - 1]));
        }
        cancellationToken.ThrowIfCancellationRequested();
        Comments destination = WordComment.GetCommentsPart(this);
        CommentsEx destinationEx = WordComment.GetCommentsExPart(this);
        for (int index = 0; index < prepared.Count; index++) {
            var item = prepared[index];
            item.Comment.AppendTo(destination, destinationEx);
            definitions[index].Cell.ParentTable.InsertComment(item.Comment, item.First, item.Last, item.Last);
        }
        destination.Save();
        destinationEx.Save();
        return prepared.Select(item => item.Comment).ToArray();

        void ObserveParagraphId(string? value) {
            if (uint.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint identifier))
                maximumParagraphId = Math.Max(maximumParagraphId, identifier);
        }
    }
}
