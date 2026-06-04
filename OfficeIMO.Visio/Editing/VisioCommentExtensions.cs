using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Editing helpers for native Visio comments.
    /// </summary>
    public static class VisioCommentExtensions {
        /// <summary>
        /// Adds a page-level native Visio comment.
        /// </summary>
        public static VisioComment AddComment(this VisioPage page, string text, string? authorName = null, string? authorInitials = null, VisioCommentOptions? options = null) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            VisioComment comment = CreateComment(page, text, authorName, authorInitials, options);
            page.Comments.Add(comment);
            return comment;
        }

        /// <summary>
        /// Adds a native Visio comment to a shape on the page.
        /// </summary>
        public static VisioComment AddComment(this VisioPage page, VisioShape target, string text, string? authorName = null, string? authorInitials = null, VisioCommentOptions? options = null) {
            if (target == null) {
                throw new ArgumentNullException(nameof(target));
            }

            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            EnsureTargetBelongsToPage(page, target.Id);
            VisioComment comment = CreateComment(page, text, authorName, authorInitials, options);
            comment.ShapeId = target.Id;
            page.Comments.Add(comment);
            return comment;
        }

        /// <summary>
        /// Adds a native Visio comment to a shape or connector by identifier.
        /// </summary>
        public static VisioComment AddCommentToShape(this VisioPage page, string shapeId, string text, string? authorName = null, string? authorInitials = null, VisioCommentOptions? options = null) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            EnsureTargetBelongsToPage(page, shapeId);
            VisioComment comment = CreateComment(page, text, authorName, authorInitials, options);
            comment.ShapeId = shapeId;
            page.Comments.Add(comment);
            return comment;
        }

        /// <summary>
        /// Returns comments that target the provided shape.
        /// </summary>
        public static IReadOnlyList<VisioComment> CommentsForShape(this VisioPage page, VisioShape target) {
            if (target == null) {
                throw new ArgumentNullException(nameof(target));
            }

            return page.CommentsForShape(target.Id);
        }

        /// <summary>
        /// Returns comments that target the provided shape or connector identifier.
        /// </summary>
        public static IReadOnlyList<VisioComment> CommentsForShape(this VisioPage page, string shapeId) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (string.IsNullOrWhiteSpace(shapeId)) {
                throw new ArgumentException("Shape id cannot be empty.", nameof(shapeId));
            }

            return page.Comments
                .Where(comment => string.Equals(comment.ShapeId, shapeId, StringComparison.Ordinal))
                .ToList();
        }

        /// <summary>
        /// Finds a native Visio comment by its page-scoped identifier.
        /// </summary>
        public static VisioComment? FindComment(this VisioPage page, int commentId) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            return page.Comments.FirstOrDefault(comment => comment.Id == commentId);
        }

        /// <summary>
        /// Returns native Visio comments that are not marked as done.
        /// </summary>
        public static IReadOnlyList<VisioComment> UnresolvedComments(this VisioPage page) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            return page.Comments.Where(comment => !comment.Done).ToList();
        }

        /// <summary>
        /// Returns native Visio comments that are marked as done.
        /// </summary>
        public static IReadOnlyList<VisioComment> ResolvedComments(this VisioPage page) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            return page.Comments.Where(comment => comment.Done).ToList();
        }

        /// <summary>
        /// Updates a native Visio comment's text by its page-scoped identifier.
        /// </summary>
        public static VisioComment UpdateCommentText(this VisioPage page, int commentId, string text, DateTimeOffset? editedAt = null) {
            VisioComment comment = RequireComment(page, commentId);
            comment.UpdateText(text, editedAt);
            return comment;
        }

        /// <summary>
        /// Marks a native Visio comment as done by its page-scoped identifier.
        /// </summary>
        public static VisioComment ResolveComment(this VisioPage page, int commentId, DateTimeOffset? editedAt = null) {
            VisioComment comment = RequireComment(page, commentId);
            comment.Resolve(editedAt);
            return comment;
        }

        /// <summary>
        /// Reopens a native Visio comment by its page-scoped identifier.
        /// </summary>
        public static VisioComment ReopenComment(this VisioPage page, int commentId, DateTimeOffset? editedAt = null) {
            VisioComment comment = RequireComment(page, commentId);
            comment.Reopen(editedAt);
            return comment;
        }

        /// <summary>
        /// Removes a native Visio comment from the page.
        /// </summary>
        public static bool RemoveComment(this VisioPage page, VisioComment comment) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (comment == null) {
                throw new ArgumentNullException(nameof(comment));
            }

            return page.Comments.Remove(comment);
        }

        /// <summary>
        /// Removes a native Visio comment by its page-scoped identifier.
        /// </summary>
        public static bool RemoveComment(this VisioPage page, int commentId) {
            VisioComment? comment = page.FindComment(commentId);
            return comment != null && page.Comments.Remove(comment);
        }

        /// <summary>
        /// Removes all native Visio comments targeting a shape or connector identifier.
        /// </summary>
        public static int RemoveCommentsForShape(this VisioPage page, string shapeId) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (string.IsNullOrWhiteSpace(shapeId)) {
                throw new ArgumentException("Shape id cannot be empty.", nameof(shapeId));
            }

            List<VisioComment> comments = page.Comments
                .Where(comment => string.Equals(comment.ShapeId, shapeId, StringComparison.Ordinal))
                .ToList();
            foreach (VisioComment comment in comments) {
                page.Comments.Remove(comment);
            }

            return comments.Count;
        }

        private static VisioComment RequireComment(VisioPage page, int commentId) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            VisioComment? comment = page.FindComment(commentId);
            if (comment == null) {
                throw new InvalidOperationException("The comment does not exist on the page.");
            }

            return comment;
        }

        private static VisioComment CreateComment(VisioPage page, string text, string? authorName, string? authorInitials, VisioCommentOptions? options) {
            if (text == null) {
                throw new ArgumentNullException(nameof(text));
            }

            VisioCommentOptions effectiveOptions = options ?? new VisioCommentOptions();
            VisioComment comment = new(text) {
                Id = GetNextCommentId(page),
                AuthorName = authorName ?? effectiveOptions.AuthorName,
                AuthorInitials = authorInitials ?? effectiveOptions.AuthorInitials,
                AuthorResolutionId = effectiveOptions.AuthorResolutionId,
                CreatedAt = effectiveOptions.CreatedAt ?? DateTimeOffset.UtcNow,
                EditedAt = effectiveOptions.EditedAt,
                Done = effectiveOptions.Done,
                AutoCommentType = effectiveOptions.AutoCommentType
            };

            return comment;
        }

        private static int GetNextCommentId(VisioPage page) {
            int nextId = 1;
            HashSet<int> usedIds = new(page.Comments.Select(comment => comment.Id));
            while (usedIds.Contains(nextId)) {
                nextId++;
            }

            return nextId;
        }

        private static void EnsureTargetBelongsToPage(VisioPage page, string shapeId) {
            if (string.IsNullOrWhiteSpace(shapeId)) {
                throw new ArgumentException("Shape id cannot be empty.", nameof(shapeId));
            }

            bool foundShape = page.FindShapeById(shapeId) != null;
            bool foundConnector = page.Connectors.Any(connector => string.Equals(connector.Id, shapeId, StringComparison.Ordinal));
            if (!foundShape && !foundConnector) {
                throw new InvalidOperationException("The comment target must belong to the page.");
            }
        }
    }
}
