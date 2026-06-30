using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word {
    /// <summary>
    /// A wrapper for Word document comments.
    /// </summary>
    public partial class WordComment {
        private WordComment(WordDocument document, Comment comment, CommentEx? commentEx) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (comment == null) throw new ArgumentNullException(nameof(comment));

            _document = document;
            _comment = comment;
            _commentEx = commentEx;

            IEnumerable<Paragraph> paragraphs = comment.ChildElements.OfType<Paragraph>();
            List<WordParagraph> list = WordSection.ConvertParagraphsToWordParagraphs(document, paragraphs).ToList();
            if (list.Count == 0) {
                throw new InvalidOperationException("A comment must contain at least one paragraph.");
            }

            _paragraph = list.Count > 1 ? list[1] : list[0];
            _list = list;
        }

        internal static string GetNewId(WordDocument document) {
            Comments comments = GetCommentsPart(document);
            int maxId = comments.Descendants<Comment>()
                .Select(e => e.Id?.Value)
                .Select(value => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? parsed : 0)
                .DefaultIfEmpty(0)
                .Max();

            return (maxId + 1).ToString(CultureInfo.InvariantCulture);
        }

        internal static string GetNewId(WordDocument document, Comments comments) {
            if (comments == null) throw new ArgumentNullException(nameof(comments));

            int maxId = comments.Descendants<Comment>()
                .Select(e => e.Id?.Value)
                .Select(value => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? parsed : 0)
                .DefaultIfEmpty(0)
                .Max();

            return (maxId + 1).ToString(CultureInfo.InvariantCulture);
        }

        internal static Comments GetCommentsPart(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            MainDocumentPart mainPart = GetMainDocumentPart(document);
            WordprocessingCommentsPart commentsPart = mainPart.WordprocessingCommentsPart ?? mainPart.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments ??= new Comments();
            return commentsPart.Comments;
        }

        internal static CommentsEx GetCommentsExPart(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            MainDocumentPart mainPart = GetMainDocumentPart(document);
            WordprocessingCommentsExPart commentsExPart = mainPart.WordprocessingCommentsExPart ?? mainPart.AddNewPart<WordprocessingCommentsExPart>();
            commentsExPart.CommentsEx ??= new CommentsEx();
            return commentsExPart.CommentsEx;
        }

        internal static CommentEx? FindCommentExForComment(Comment comment, IReadOnlyList<CommentEx> commentsEx, int fallbackIndex) {
            if (comment == null) throw new ArgumentNullException(nameof(comment));
            if (commentsEx == null) throw new ArgumentNullException(nameof(commentsEx));

            string? paraId = GetCommentParagraphId(comment);
            if (!string.IsNullOrWhiteSpace(paraId)) {
                CommentEx? matched = commentsEx.FirstOrDefault(commentEx =>
                    string.Equals(commentEx.ParaId?.Value, paraId, StringComparison.Ordinal));
                return matched;
            }

            return fallbackIndex >= 0 && fallbackIndex < commentsEx.Count
                ? commentsEx[fallbackIndex]
                : null;
        }

        internal static string? GetCommentParagraphId(Comment comment) {
            if (comment == null) throw new ArgumentNullException(nameof(comment));

            return comment.Elements<Paragraph>()
                .Select(paragraph => paragraph.ParagraphId?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        }

        internal static string GetNewParaId(CommentsEx commentsEx, Comments? comments = null) {
            if (commentsEx == null) throw new ArgumentNullException(nameof(commentsEx));

            var existing = commentsEx
                .Descendants<CommentEx>()
                .Select(c => c.ParaId?.Value)
                .Where(v => !string.IsNullOrEmpty(v))
                .ToList();
            if (comments != null) {
                existing.AddRange(comments
                    .Descendants<Comment>()
                    .Select(GetCommentParagraphId)
                    .Where(v => !string.IsNullOrEmpty(v))
                    .Select(v => v!));
            }

            int max = 0;
            foreach (var v in existing) {
                if (int.TryParse(v, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int num)) {
                    if (num > max) max = num;
                }
            }

            return (max + 1).ToString("X8", CultureInfo.InvariantCulture);
        }

        private static MainDocumentPart GetMainDocumentPart(WordDocument document) {
            return document._wordprocessingDocument?.MainDocumentPart ?? throw new InvalidOperationException("The Word document is not associated with a main document part.");
        }
    }
}
