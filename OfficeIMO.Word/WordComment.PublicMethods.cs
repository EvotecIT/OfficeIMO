using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// A wrapper for Word document comments.
    /// </summary>
    public partial class WordComment {

        /// <summary>
        /// Creates a new comment in the specified document.
        /// </summary>
        /// <param name="document">Document to which the comment will be added.</param>
        /// <param name="author">Author of the comment.</param>
        /// <param name="initials">Initials of the author.</param>
        /// <param name="comment">Comment text.</param>
        /// <param name="parent">Optional parent comment when creating a reply.</param>
        /// <returns>The newly created <see cref="WordComment"/>.</returns>
        public static WordComment Create(WordDocument document, string author, string initials, string comment, WordComment? parent = null) {
            var comments = GetCommentsPart(document);
            var commentsEx = GetCommentsExPart(document);
            var paraId = GetNewParaId(commentsEx, comments);
            // Compose a new Comment and add it to the Comments part.
            Paragraph p = new Paragraph(new Run(new Text(comment))) {
                ParagraphId = paraId
            };
            Comment cmt = new Comment() {
                Id = GetNewId(document),
                Author = author,
                Initials = initials,
                Date = System.DateTime.Now
            };
            cmt.AppendChild(p);
            comments.AppendChild(cmt);
            comments.Save();

            CommentEx cmtEx = new CommentEx() { ParaId = paraId };
            if (parent != null) {
                cmtEx.ParaIdParent = EnsureCommentParaId(parent, comments, commentsEx);
            }
            commentsEx.AppendChild(cmtEx);
            commentsEx.Save();

            return new WordComment(document, cmt, cmtEx);
        }

        internal static WordComment Create(WordDocument document, string author, string initials, IReadOnlyList<Paragraph> paragraphs, WordComment? parent = null) {
            if (paragraphs == null) throw new ArgumentNullException(nameof(paragraphs));

            var comments = GetCommentsPart(document);
            var commentsEx = GetCommentsExPart(document);
            var paraId = GetNewParaId(commentsEx, comments);
            Comment cmt = new Comment() {
                Id = GetNewId(document),
                Author = author,
                Initials = initials,
                Date = System.DateTime.Now
            };

            if (paragraphs.Count == 0) {
                cmt.AppendChild(new Paragraph(new Run(new Text(string.Empty))) {
                    ParagraphId = paraId
                });
            } else {
                for (int index = 0; index < paragraphs.Count; index++) {
                    Paragraph paragraph = (Paragraph)paragraphs[index].CloneNode(true);
                    if (index == 0) {
                        paragraph.ParagraphId = paraId;
                    }
                    cmt.AppendChild(paragraph);
                }
            }

            comments.AppendChild(cmt);
            comments.Save();

            CommentEx cmtEx = new CommentEx() { ParaId = paraId };
            if (parent != null) {
                cmtEx.ParaIdParent = EnsureCommentParaId(parent, comments, commentsEx);
            }
            commentsEx.AppendChild(cmtEx);
            commentsEx.Save();

            return new WordComment(document, cmt, cmtEx);
        }


        /// <summary>
        /// Retrieves all comments from the provided document.
        /// </summary>
        /// <param name="document">Word document containing comments.</param>
        /// <returns>List of <see cref="WordComment"/> objects.</returns>
        public static List<WordComment> GetAllComments(WordDocument document) {
            List<WordComment> comments = new List<WordComment>();
            var part = document._wordprocessingDocument.MainDocumentPart;
            if (part?.WordprocessingCommentsPart != null && part.WordprocessingCommentsPart.Comments != null) {
                var commentList = part.WordprocessingCommentsPart.Comments.OfType<Comment>().ToList();
                var commentExList = part.WordprocessingCommentsExPart?.CommentsEx?.OfType<CommentEx>().ToList() ?? new List<CommentEx>();
                Dictionary<string, CommentEx> commentsExByParagraphId = IndexCommentExByParagraphId(commentExList);
                for (int i = 0; i < commentList.Count; i++) {
                    var ce = FindCommentExForComment(commentList[i], commentExList, commentsExByParagraphId, i)
                        ?? new CommentEx();
                    comments.Add(new WordComment(document, commentList[i], ce));
                }
            }
            return comments;
        }

        /// <summary>
        /// Deletes this comment and removes all references from the document.
        /// </summary>
        public void Remove() {
            var commentsPart = _document._wordprocessingDocument.MainDocumentPart?.WordprocessingCommentsPart;
            var commentsExPart = _document._wordprocessingDocument.MainDocumentPart?.WordprocessingCommentsExPart;
            int index = -1;
            if (commentsPart?.Comments != null) {
                var list = commentsPart.Comments.Elements<Comment>().ToList();
                index = list.FindIndex(c => c.Id == _comment.Id);
                var cmt = index >= 0 ? list[index] : null;
                cmt?.Remove();
                commentsPart.Comments.Save();
            }
            if (index >= 0 && commentsExPart?.CommentsEx != null) {
                List<CommentEx> commentExList = commentsExPart.CommentsEx.Elements<CommentEx>().ToList();
                CommentEx? commentEx = FindCommentExForComment(_comment, commentExList, index);
                commentEx?.Remove();
                if (commentEx != null) {
                    commentsExPart.CommentsEx.Save();
                }
            }

            RemoveReferencesFromAllContentParts();
        }

        /// <summary>
        /// Creates a reply to this comment.
        /// </summary>
        /// <param name="author">Author of the reply.</param>
        /// <param name="initials">Initials of the author.</param>
        /// <param name="comment">Reply text.</param>
        /// <returns>Newly created reply comment.</returns>
        public WordComment AddReply(string author, string initials, string comment) {
            return Create(_document, author, initials, comment, this);
        }

        /// <summary>
        /// Marks this comment as resolved in the modern comment metadata part.
        /// </summary>
        public WordComment MarkResolved() {
            return SetResolved(true);
        }

        /// <summary>
        /// Marks this comment as unresolved in the modern comment metadata part.
        /// </summary>
        public WordComment MarkUnresolved() {
            return SetResolved(false);
        }

        /// <summary>
        /// Sets the resolved state for this comment in the modern comment metadata part.
        /// </summary>
        /// <param name="resolved">Whether the comment should be marked resolved.</param>
        public WordComment SetResolved(bool resolved) {
            CommentEx commentEx = EnsureCommentEx();
            commentEx.Done = resolved;
            SaveCommentsEx();
            return this;
        }

        /// <summary>
        /// Deletes this comment and all replies in its thread.
        /// </summary>
        public void DeleteThread() {
            foreach (WordComment reply in Replies.ToList()) {
                reply.DeleteThread();
            }

            Remove();
        }

        private void RemoveReferencesFromAllContentParts() {
            var mainPart = _document._wordprocessingDocument.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart is missing.");
            string commentId = _comment.Id?.Value ?? string.Empty;
            if (commentId.Length == 0) {
                return;
            }

            foreach (var root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                foreach (var start in root.Root.Descendants<CommentRangeStart>()
                    .Where(c => string.Equals(c.Id?.Value, commentId, StringComparison.Ordinal))
                    .ToList()) {
                    start.Remove();
                }

                foreach (var end in root.Root.Descendants<CommentRangeEnd>()
                    .Where(c => string.Equals(c.Id?.Value, commentId, StringComparison.Ordinal))
                    .ToList()) {
                    end.Remove();
                }

                foreach (var reference in root.Root.Descendants<CommentReference>()
                    .Where(c => string.Equals(c.Id?.Value, commentId, StringComparison.Ordinal))
                    .ToList()) {
                    DocumentFormat.OpenXml.OpenXmlElement? parent = reference.Parent;
                    reference.Remove();
                    if (parent is Run run && !run.ChildElements.Any(child => child is not RunProperties)) {
                        run.Remove();
                    }
                }
            }
        }

        private CommentEx? FindCommentEx() {
            var commentsPart = _document._wordprocessingDocument.MainDocumentPart?.WordprocessingCommentsPart;
            var commentsEx = _document._wordprocessingDocument.MainDocumentPart?.WordprocessingCommentsExPart?.CommentsEx;
            if (commentsPart?.Comments == null || commentsEx == null) {
                return _commentEx;
            }

            List<Comment> commentList = commentsPart.Comments.Elements<Comment>().ToList();
            int index = commentList.FindIndex(c => c.Id == _comment.Id);
            if (index < 0) {
                return _commentEx;
            }

            _commentEx = FindCommentExForComment(_comment, commentsEx.Elements<CommentEx>().ToList(), index);
            return _commentEx;
        }

        private CommentEx EnsureCommentEx() {
            var mainPart = _document._wordprocessingDocument.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart is missing.");
            var commentsPart = mainPart.WordprocessingCommentsPart;
            if (commentsPart?.Comments == null) {
                throw new InvalidOperationException("Comments part is missing.");
            }

            int index = commentsPart.Comments.Elements<Comment>().ToList().FindIndex(c => c.Id == _comment.Id);
            if (index < 0) {
                throw new InvalidOperationException("Comment is no longer present in the document.");
            }

            CommentsEx commentsEx = GetCommentsExPart(_document);
            List<CommentEx> commentExList = commentsEx.Elements<CommentEx>().ToList();
            CommentEx? commentEx = FindCommentExForComment(_comment, commentExList, index);
            if (commentEx != null) {
                _commentEx = commentEx;
                return commentEx;
            }

            string? paragraphId = GetCommentParagraphId(_comment);
            if (!string.IsNullOrWhiteSpace(paragraphId)) {
                commentEx = new CommentEx { ParaId = paragraphId };
                commentsEx.AppendChild(commentEx);
                _commentEx = commentEx;
                return commentEx;
            }

            while (commentsEx.Elements<CommentEx>().Count() <= index) {
                commentsEx.AppendChild(new CommentEx { ParaId = GetNewParaId(commentsEx, commentsPart.Comments) });
            }

            commentEx = commentsEx.Elements<CommentEx>().ElementAt(index);
            if (string.IsNullOrWhiteSpace(commentEx.ParaId?.Value)) {
                commentEx.ParaId = _commentEx?.ParaId ?? GetNewParaId(commentsEx, commentsPart.Comments);
            }

            string? ensuredParagraphId = commentEx.ParaId?.Value;
            if (!string.IsNullOrWhiteSpace(ensuredParagraphId) && string.IsNullOrWhiteSpace(GetCommentParagraphId(_comment))) {
                Paragraph paragraph = _comment.Elements<Paragraph>().First();
                paragraph.ParagraphId = ensuredParagraphId;
                commentsPart.Comments.Save();
            }

            _commentEx = commentEx;
            return commentEx;
        }

        private void SaveCommentsEx() {
            var commentsExPart = _document._wordprocessingDocument.MainDocumentPart?.WordprocessingCommentsExPart;
            commentsExPart?.CommentsEx?.Save();
        }
    }
}
