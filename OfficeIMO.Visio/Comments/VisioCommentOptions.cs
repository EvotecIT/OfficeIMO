using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Options used when creating native Visio comments.
    /// </summary>
    public sealed class VisioCommentOptions {
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
        /// Gets or sets the creation timestamp. Defaults to the current UTC time.
        /// </summary>
        public DateTimeOffset? CreatedAt { get; set; }

        /// <summary>
        /// Gets or sets the last edit timestamp.
        /// </summary>
        public DateTimeOffset? EditedAt { get; set; }

        /// <summary>
        /// Gets or sets whether the comment should be marked as done.
        /// </summary>
        public bool Done { get; set; }

        /// <summary>
        /// Gets or sets the native AutoCommentType value. Visio ignores this value.
        /// </summary>
        public int? AutoCommentType { get; set; }
    }
}
