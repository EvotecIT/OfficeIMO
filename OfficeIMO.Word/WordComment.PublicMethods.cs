using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordComment {

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
    }
}
