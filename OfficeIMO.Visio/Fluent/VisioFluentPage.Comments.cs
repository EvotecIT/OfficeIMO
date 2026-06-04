using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio.Fluent {
    public partial class VisioFluentPage {
        /// <summary>
        /// Adds a page-level native Visio comment.
        /// </summary>
        public VisioFluentPage Comment(string text, string? authorName = null, string? authorInitials = null, VisioCommentOptions? options = null) {
            Page.AddComment(text, authorName, authorInitials, options);
            return this;
        }

        /// <summary>
        /// Adds a native Visio comment to a shape or connector by identifier.
        /// </summary>
        public VisioFluentPage CommentShape(string shapeId, string text, string? authorName = null, string? authorInitials = null, VisioCommentOptions? options = null) {
            Page.AddCommentToShape(shapeId, text, authorName, authorInitials, options);
            return this;
        }

        /// <summary>
        /// Updates a native Visio comment by its page-scoped identifier.
        /// </summary>
        public VisioFluentPage UpdateComment(int commentId, string text, DateTimeOffset? editedAt = null) {
            Page.UpdateCommentText(commentId, text, editedAt);
            return this;
        }

        /// <summary>
        /// Marks a native Visio comment as done by its page-scoped identifier.
        /// </summary>
        public VisioFluentPage ResolveComment(int commentId, DateTimeOffset? editedAt = null) {
            Page.ResolveComment(commentId, editedAt);
            return this;
        }

        /// <summary>
        /// Reopens a native Visio comment by its page-scoped identifier.
        /// </summary>
        public VisioFluentPage ReopenComment(int commentId, DateTimeOffset? editedAt = null) {
            Page.ReopenComment(commentId, editedAt);
            return this;
        }

        /// <summary>
        /// Removes a native Visio comment by its page-scoped identifier.
        /// </summary>
        public VisioFluentPage RemoveComment(int commentId) {
            Page.RemoveComment(commentId);
            return this;
        }

        /// <summary>
        /// Configures all native Visio comments on the page.
        /// </summary>
        public VisioFluentPage Comments(Action<IList<VisioComment>> configure) {
            if (configure == null) {
                throw new ArgumentNullException(nameof(configure));
            }

            configure(Page.Comments);
            return this;
        }

        /// <summary>
        /// Configures native Visio comments attached to a shape or connector identifier.
        /// </summary>
        public VisioFluentPage ShapeComments(string shapeId, Action<IReadOnlyList<VisioComment>> configure) {
            if (configure == null) {
                throw new ArgumentNullException(nameof(configure));
            }

            configure(Page.CommentsForShape(shapeId));
            return this;
        }
    }
}
