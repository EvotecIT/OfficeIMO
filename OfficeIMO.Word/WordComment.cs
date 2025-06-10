using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordComment : WordElement {
        private WordParagraph _paragraph;
        private readonly WordDocument _document;
        private readonly Comment _comment;
        private readonly List<WordParagraph> _list;

        /// <summary>
        /// ID of a comment
        /// </summary>
        public string Id => _comment.Id;

        /// <summary>
        /// Text content of a comment
        /// </summary>
        public string Text {
            get {
                return _paragraph.Text;
            }
            set {
                _paragraph.Text = value;
            }
        }

        /// <summary>
        /// Initials of a person who created a comment
        /// </summary>
        public string Initials {
            get {
                return _comment.Initials;
            }
            set {
                _comment.Initials = value;
            }
        }

        /// <summary>
        /// Full name of a person who created a comment
        /// </summary>
        public string Author {
            get {
                return _comment.Author;
            }
            set {
                _comment.Author = value;
            }
        }

        /// <summary>
        /// DateTime when the comment was created
        /// </summary>
        public DateTime DateTime {
            get => _comment.Date;
            set => _comment.Date = value;
        }
    }
}
