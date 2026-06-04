using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a native Visio comment stored in the document comments part.
    /// </summary>
    public sealed class VisioComment {
        /// <summary>
        /// Initializes a new comment with the current UTC creation time.
        /// </summary>
        /// <param name="text">Comment text.</param>
        public VisioComment(string text) {
            Text = text ?? throw new ArgumentNullException(nameof(text));
            CreatedAt = DateTimeOffset.UtcNow;
        }

        /// <summary>
        /// Gets the comment identifier, unique within its page once attached to a page.
        /// </summary>
        public int Id { get; internal set; }

        /// <summary>
        /// Gets or sets the plain text comment content.
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// Gets or sets the display name of the comment author.
        /// </summary>
        public string? AuthorName { get; set; }

        /// <summary>
        /// Gets or sets the author initials shown by Visio.
        /// </summary>
        public string? AuthorInitials { get; set; }

        /// <summary>
        /// Gets or sets the optional Visio author resolution identifier.
        /// </summary>
        public string? AuthorResolutionId { get; set; }

        /// <summary>
        /// Gets or sets the semantic shape or connector identifier this comment targets.
        /// When null, the comment applies to the page.
        /// </summary>
        public string? ShapeId { get; set; }

        /// <summary>
        /// Gets or sets when the comment was created.
        /// </summary>
        public DateTimeOffset CreatedAt { get; set; }

        /// <summary>
        /// Gets or sets when the comment was last edited.
        /// </summary>
        public DateTimeOffset? EditedAt { get; set; }

        /// <summary>
        /// Gets or sets whether the comment is marked as done.
        /// </summary>
        public bool Done { get; set; }

        /// <summary>
        /// Gets or sets the native AutoCommentType value. Visio ignores this value, but it is preserved when present.
        /// </summary>
        public int? AutoCommentType { get; set; }

        /// <summary>
        /// Updates the comment text and records the edit timestamp.
        /// </summary>
        /// <param name="text">New comment text.</param>
        /// <param name="editedAt">Optional edit timestamp. Uses current UTC time when not provided.</param>
        public void UpdateText(string text, DateTimeOffset? editedAt = null) {
            Text = text ?? throw new ArgumentNullException(nameof(text));
            EditedAt = editedAt ?? DateTimeOffset.UtcNow;
        }

        /// <summary>
        /// Marks the comment as done and records the edit timestamp.
        /// </summary>
        /// <param name="editedAt">Optional edit timestamp. Uses current UTC time when not provided.</param>
        public void Resolve(DateTimeOffset? editedAt = null) {
            Done = true;
            EditedAt = editedAt ?? DateTimeOffset.UtcNow;
        }

        /// <summary>
        /// Reopens a resolved comment and records the edit timestamp.
        /// </summary>
        /// <param name="editedAt">Optional edit timestamp. Uses current UTC time when not provided.</param>
        public void Reopen(DateTimeOffset? editedAt = null) {
            Done = false;
            EditedAt = editedAt ?? DateTimeOffset.UtcNow;
        }
    }
}
