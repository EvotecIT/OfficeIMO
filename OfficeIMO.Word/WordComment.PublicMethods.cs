using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office2013.Word;

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
            // Compose a new Comment and add it to the Comments part.
            Paragraph p = new Paragraph(new Run(new Text(comment)));
            Comment cmt = new Comment() {
                Id = GetNewId(document),
                Author = author,
                Initials = initials,
                Date = System.DateTime.Now
            };
            cmt.AppendChild(p);
            comments.AppendChild(cmt);
            comments.Save();

            var paraId = GetNewParaId(commentsEx);
            CommentEx cmtEx = new CommentEx() { ParaId = paraId };
            if (parent != null) {
                cmtEx.ParaIdParent = parent.ParaId;
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
            if (part.WordprocessingCommentsPart != null && part.WordprocessingCommentsPart.Comments != null) {
                var commentList = part.WordprocessingCommentsPart.Comments.OfType<Comment>().ToList();
                var commentExList = part.WordprocessingCommentsExPart?.CommentsEx?.OfType<CommentEx>().ToList() ?? new List<CommentEx>();
                for (int i = 0; i < commentList.Count; i++) {
                    var ce = i < commentExList.Count ? commentExList[i] : new CommentEx();
                    comments.Add(new WordComment(document, commentList[i], ce));
                }
            }
            return comments;
        }

        /// <summary>
        /// Deletes this comment and removes all references from the document.
        /// </summary>
        public void Delete() {
            var commentsPart = _document._wordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart;
            var commentsExPart = _document._wordprocessingDocument.MainDocumentPart.WordprocessingCommentsExPart;
            int index = -1;
            if (commentsPart?.Comments != null) {
                var list = commentsPart.Comments.Elements<Comment>().ToList();
                index = list.FindIndex(c => c.Id == _comment.Id);
                var cmt = index >= 0 ? list[index] : null;
                cmt?.Remove();
                commentsPart.Comments.Save();
            }
            if (index >= 0 && commentsExPart?.CommentsEx != null && index < commentsExPart.CommentsEx.Count()) {
                commentsExPart.CommentsEx.ChildElements[index].Remove();
                commentsExPart.CommentsEx.Save();
            }

            var body = _document._document.Body;
            foreach (var start in body.Descendants<CommentRangeStart>().Where(c => c.Id == _comment.Id).ToList()) {
                start.Remove();
            }
            foreach (var end in body.Descendants<CommentRangeEnd>().Where(c => c.Id == _comment.Id).ToList()) {
                end.Remove();
            }
            foreach (var reference in body.Descendants<CommentReference>().Where(c => c.Id == _comment.Id).ToList()) {
                reference.Parent?.Remove();
            }
        }

        /// <summary>
        /// Removes this comment from the document. Alias for <see cref="Delete"/>.
        /// </summary>
        public void Remove() {
            Delete();
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
    }
}
