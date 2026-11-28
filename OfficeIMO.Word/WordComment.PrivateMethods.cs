using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    // A wrapper for Word document comments.
    public partial class WordComment {
        private WordComment(WordDocument document, Comment comment, CommentEx? commentEx) {
            ArgumentNullException.ThrowIfNull(document);
            ArgumentNullException.ThrowIfNull(comment);

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

        internal static Comments GetCommentsPart(WordDocument document) {
            ArgumentNullException.ThrowIfNull(document);

            MainDocumentPart mainPart = GetMainDocumentPart(document);
            WordprocessingCommentsPart commentsPart = mainPart.WordprocessingCommentsPart ?? mainPart.AddNewPart<WordprocessingCommentsPart>();
            if (commentsPart.Comments == null) {
                // Comment paragraphs carry w14:paraId, so a new part declares w14 as ignorable like the other story parts.
                var comments = new Comments();
                comments.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                comments.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                comments.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "w14" };
                commentsPart.Comments = comments;
            }
            return commentsPart.Comments;
        }

        internal static CommentsEx GetCommentsExPart(WordDocument document) {
            ArgumentNullException.ThrowIfNull(document);

            MainDocumentPart mainPart = GetMainDocumentPart(document);
            WordprocessingCommentsExPart commentsExPart = mainPart.WordprocessingCommentsExPart ?? mainPart.AddNewPart<WordprocessingCommentsExPart>();
            commentsExPart.CommentsEx ??= new CommentsEx();
            return commentsExPart.CommentsEx;
        }

        internal static CommentEx? FindCommentExForComment(Comment comment, IReadOnlyList<CommentEx> commentsEx, int fallbackIndex) {
            ArgumentNullException.ThrowIfNull(comment);
            ArgumentNullException.ThrowIfNull(commentsEx);

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
            ArgumentNullException.ThrowIfNull(commentsEx);

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
            ArgumentNullException.ThrowIfNull(comment);
            ArgumentNullException.ThrowIfNull(commentsEx);
            ArgumentNullException.ThrowIfNull(commentsExByParagraphId);

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
            ArgumentNullException.ThrowIfNull(comment);

            return comment.Elements<Paragraph>()
                .Select(paragraph => paragraph.ParagraphId?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        }

        private static string EnsureCommentParaId(WordComment comment, WordCommentIdentityState identities, CommentsEx commentsEx) {
            ArgumentNullException.ThrowIfNull(comment);
            ArgumentNullException.ThrowIfNull(identities);
            ArgumentNullException.ThrowIfNull(commentsEx);

            // The identity state normalized legacy comments when it was built, so only a parent whose paragraphs
            // were replaced since then still lacks an identifier.
            string? paraId = comment.ParaId;
            if (string.IsNullOrWhiteSpace(paraId)) {
                paraId = identities.NextParagraphId();
                Paragraph paragraph = comment._comment.Elements<Paragraph>().First();
                paragraph.ParagraphId = paraId;
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
            return paraId!;
        }

        private static MainDocumentPart GetMainDocumentPart(WordDocument document) {
            return document._wordprocessingDocument?.MainDocumentPart ?? throw new InvalidOperationException("The Word document is not associated with a main document part.");
        }
    }
}
