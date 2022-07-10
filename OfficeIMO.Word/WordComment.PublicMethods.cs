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
    }
}
