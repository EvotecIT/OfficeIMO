using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office2013.Word;

namespace OfficeIMO.Word {
    /// <summary>
    /// A wrapper for Word document comments.
    /// </summary>
    public partial class WordComment : WordElement {
        private WordParagraph _paragraph;
        private readonly WordDocument _document;
        private readonly Comment _comment;
        private readonly CommentEx _commentEx;
        private readonly List<WordParagraph> _list;

        /// <summary>
        /// ID of a comment
        /// </summary>
        public string? Id => _comment.Id;

        /// <summary>
        /// Identifier used to link threaded replies.
        /// </summary>
        public string? ParaId => _commentEx?.ParaId;

        /// <summary>
        /// Identifier of parent comment if this comment is a reply.
        /// </summary>
        public string? ParentParaId => _commentEx?.ParaIdParent;

        /// <summary>
        /// Parent comment instance if available.
        /// </summary>
        public WordComment? ParentComment => _document.Comments.FirstOrDefault(c => c.ParaId == ParentParaId);

        /// <summary>
        /// Replies for this comment.
        /// </summary>
        public List<WordComment> Replies => _document.Comments.Where(c => c.ParentParaId == ParaId).ToList();

        /// <summary>
        /// Text content of a comment
        /// </summary>
        public string? Text {
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
        public string? Initials {
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
        public string? Author {
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
        public DateTime? DateTime {
            get => _comment.Date?.Value;
            set => _comment.Date = value;
        }
    }
}
