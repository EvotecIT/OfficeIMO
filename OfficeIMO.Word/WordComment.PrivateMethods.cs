using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordComment {
        private WordComment(WordDocument document, Comment comment) {
            _document = document;
            _comment = comment;

            var paragraph = comment.ChildElements.OfType<Paragraph>();
            List<WordParagraph> list = WordSection.ConvertParagraphsToWordParagraphs(document, paragraph);
            _paragraph = list.Count > 1 ? list[1] : list[0];
            _list = list;
        }

        internal static string GetNewId(WordDocument document) {
            string id = "0";
            var comments = GetCommentsPart(document);
            if (comments.HasChildren) {
                // Obtain an unused ID.
                id = (comments.Descendants<Comment>().Select(e => int.Parse(e.Id.Value)).Max() + 1).ToString();
            }
            return id;
        }

        internal static string GetNewId(WordDocument document, Comments comments) {
            string id = "0";
            if (comments.HasChildren) {
                // Obtain an unused ID.
                id = (comments.Descendants<Comment>().Select(e => int.Parse(e.Id.Value)).Max() + 1).ToString();
            }
            return id;
        }

        internal static Comments GetCommentsPart(WordDocument document) {
            Comments comments = null;
            if (document._wordprocessingDocument.MainDocumentPart != null && document._wordprocessingDocument.MainDocumentPart.GetPartsOfType<WordprocessingCommentsPart>().Any()) {
                if (document._wordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart != null)
                    comments = document._wordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart.Comments;
            } else {
                // No WordprocessingCommentsPart part exists, so add one to the package.
                if (document._wordprocessingDocument.MainDocumentPart != null) {
                    WordprocessingCommentsPart commentPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                    commentPart.Comments = new Comments();
                    comments = commentPart.Comments;
                }
            }

            return comments;
        }
    }
}
