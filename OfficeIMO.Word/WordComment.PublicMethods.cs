using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
        /// <returns>The newly created <see cref="WordComment"/>.</returns>
        public static WordComment Create(WordDocument document, string author, string initials, string comment) {
            var comments = GetCommentsPart(document);
            // Compose a new Comment and add it to the Comments part.
            Paragraph p = new Paragraph(new Run(new Text(comment)));
            Comment cmt = new Comment() {
                Id = GetNewId(document),
                Author = author,
                Initials = initials,
                Date = DateTime.Now
            };
            cmt.AppendChild(p);
            comments.AppendChild(cmt);
            comments.Save();

            return new WordComment(document, cmt);
        }


        /// <summary>
        /// Retrieves all comments from the provided document.
        /// </summary>
        /// <param name="document">Word document containing comments.</param>
        /// <returns>List of <see cref="WordComment"/> objects.</returns>
        public static List<WordComment> GetAllComments(WordDocument document) {
            List<WordComment> comments = new List<WordComment>();
            if (document._wordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart != null && document._wordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart.Comments != null) {
                foreach (var comment in document._wordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart.Comments.OfType<Comment>()) {
                    comments.Add(new WordComment(document, comment));
                }
            }
            return comments;
        }

        /// <summary>
        /// Deletes this comment and removes all references from the document.
        /// </summary>
        public void Delete() {
            var commentsPart = _document._wordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart;
            if (commentsPart?.Comments != null) {
                var cmt = commentsPart.Comments.Elements<Comment>().FirstOrDefault(c => c.Id == _comment.Id);
                cmt?.Remove();
                commentsPart.Comments.Save();
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
    }
}
