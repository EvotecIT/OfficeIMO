using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// A wrapper for Word document comments.
    /// </summary>
    public partial class WordComment : WordElement {
        private readonly WordParagraph _paragraph;
        private readonly WordDocument _document;
        private readonly Comment _comment;
        private CommentEx? _commentEx;
        private readonly List<WordParagraph> _list;

        /// <summary>
        /// ID of a comment
        /// </summary>
        public string? Id => _comment.Id;

        /// <summary>
        /// Identifier used to link threaded replies.
        /// </summary>
        public string? ParaId => _commentEx?.ParaId?.Value ?? GetCommentParagraphId(_comment);

        /// <summary>
        /// Identifier of parent comment if this comment is a reply.
        /// </summary>
        public string? ParentParaId => _commentEx?.ParaIdParent?.Value;

        /// <summary>
        /// Whether the comment is marked resolved. A null value means the document has no resolved-state metadata for this comment.
        /// </summary>
        public bool? IsResolved => _commentEx?.Done?.Value;

        /// <summary>
        /// Parent comment instance if available.
        /// </summary>
        public WordComment? ParentComment => _document.Comments.FirstOrDefault(c => c.ParaId == ParentParaId);

        /// <summary>
        /// Replies for this comment.
        /// </summary>
        public List<WordComment> Replies => _document.Comments.Where(c => c.ParentParaId == ParaId).ToList();

        /// <summary>
        /// Paragraph and run views for converters that need to preserve rich comment content.
        /// </summary>
        internal IReadOnlyList<WordParagraph> Paragraphs => _list;

        /// <summary>
        /// Text content of a comment
        /// </summary>
        public string? Text {
            get {
                return string.Concat(_list.Select(paragraph => paragraph.Text));
            }
            set {
                _paragraph.Text = value ?? string.Empty;
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
