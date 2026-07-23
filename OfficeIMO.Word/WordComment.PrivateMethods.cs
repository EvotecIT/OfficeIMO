using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
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

        internal static Dictionary<string, CommentEx> IndexCommentExByParagraphId(
            IReadOnlyList<CommentEx> commentsEx) {
            if (commentsEx == null) throw new ArgumentNullException(nameof(commentsEx));

            var indexed = new Dictionary<string, CommentEx>(StringComparer.Ordinal);
            foreach (CommentEx commentEx in commentsEx) {
                string? paraId = commentEx.ParaId?.Value;
                if (!string.IsNullOrWhiteSpace(paraId)) {
                    string key = paraId!;
                    if (!indexed.ContainsKey(key)) indexed.Add(key, commentEx);
                }
            }

            return indexed;
        }

        internal static CommentEx? FindCommentExForComment(Comment comment,
            IReadOnlyList<CommentEx> commentsEx, IReadOnlyDictionary<string, CommentEx> commentsExByParagraphId,
            int fallbackIndex) {
            if (comment == null) throw new ArgumentNullException(nameof(comment));
            if (commentsEx == null) throw new ArgumentNullException(nameof(commentsEx));
            if (commentsExByParagraphId == null) throw new ArgumentNullException(nameof(commentsExByParagraphId));

            string? paraId = GetCommentParagraphId(comment);
            if (!string.IsNullOrWhiteSpace(paraId)) {
                return commentsExByParagraphId.TryGetValue(paraId!, out CommentEx? matched)
                    ? matched
                    : null;
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

            var used = new HashSet<uint>();
            uint max = 0;
            foreach (var v in existing) {
                if (uint.TryParse(v, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint num)) {
                    used.Add(num);
                    if (num > max) max = num;
                }
            }

            uint candidate = max + 1U;
            while (used.Contains(candidate)) {
                candidate++;
            }

            return candidate.ToString("X8", CultureInfo.InvariantCulture);
        }

        private static string EnsureCommentParaId(WordComment comment, Comments comments, CommentsEx commentsEx) {
            if (comment == null) throw new ArgumentNullException(nameof(comment));
            if (comments == null) throw new ArgumentNullException(nameof(comments));
            if (commentsEx == null) throw new ArgumentNullException(nameof(commentsEx));

            string? paraId = comment.ParaId;
            if (string.IsNullOrWhiteSpace(paraId)) {
                paraId = GetNewParaId(commentsEx, comments);
                Paragraph paragraph = comment._comment.Elements<Paragraph>().First();
                paragraph.ParagraphId = paraId;
                comments.Save();
            }

            CommentEx? commentEx = comment.FindCommentEx();
            if (commentEx == null || commentEx.Parent == null) {
                commentEx = new CommentEx();
                commentsEx.AppendChild(commentEx);
            }

            if (!string.Equals(commentEx.ParaId?.Value, paraId, StringComparison.Ordinal)) {
                commentEx.ParaId = paraId;
            }

            comment._commentEx = commentEx;
            commentsEx.Save();
            return paraId!;
        }

        private static MainDocumentPart GetMainDocumentPart(WordDocument document) {
            return document._wordprocessingDocument?.MainDocumentPart ?? throw new InvalidOperationException("The Word document is not associated with a main document part.");
        }
    }
}
