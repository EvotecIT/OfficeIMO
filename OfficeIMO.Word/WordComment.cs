using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordComment {

        private WordParagraph _paragraph;
        private readonly WordDocument _document;
        private readonly Comment _comment;
        private readonly List<WordParagraph> _list;

        public string Text {
            get {
                return _paragraph.Text;
            }
            set {
                _paragraph.Text = value;
            }
        }

        public string Initials {
            get {
                return _comment.Initials;
            }
            set {
                _comment.Initials = value;
            }
        }

        public string Author {
            get {
                return _comment.Author;
            }
            set {
                _comment.Author = value;
            }
        }

        public DateTime DateTime {
            get {
                return _comment.Date;
            }
            set {
                _comment.Date = value;
            }
        }

        private WordComment(WordDocument document, Comment comment) {
            _document = document;
            _comment = comment;

            var paragraph = comment.ChildElements.OfType<Paragraph>();
            List<WordParagraph> list = WordSection.ConvertParagraphsToWordParagraphs(document, paragraph);
            _paragraph = list.Count > 1 ? list[1] : list[0];
            _list = list;
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
