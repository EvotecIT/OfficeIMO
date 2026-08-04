using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Visio {
    /// <summary>Stable author identity used by Visio comment workflows.</summary>
    public sealed class VisioCommentAuthor {
        /// <summary>Creates a comment author.</summary>
        public VisioCommentAuthor(string name, string initials,
            string? resolutionId = null) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Author name cannot be empty.", nameof(name));
            if (string.IsNullOrWhiteSpace(initials)) throw new ArgumentException("Author initials cannot be empty.", nameof(initials));
            Name = name;
            Initials = initials;
            ResolutionId = resolutionId;
        }

        /// <summary>Display name.</summary>
        public string Name { get; }

        /// <summary>Display initials.</summary>
        public string Initials { get; }

        /// <summary>Optional producer identity used to resolve the author.</summary>
        public string? ResolutionId { get; }
    }

    /// <summary>
    /// OfficeIMO thread projection over native Visio comments. Each comment remains a
    /// native CommentEntry; parent/thread identifiers are retained as extension metadata.
    /// </summary>
    public sealed class VisioCommentThread {
        internal VisioCommentThread(string id, VisioComment root,
            IReadOnlyList<VisioComment> comments) {
            Id = id;
            Root = root;
            Comments = new ReadOnlyCollection<VisioComment>(
                new List<VisioComment>(comments));
        }

        /// <summary>Stable thread identifier.</summary>
        public string Id { get; }

        /// <summary>Root native comment.</summary>
        public VisioComment Root { get; }

        /// <summary>Root and replies ordered by creation time and comment id.</summary>
        public IReadOnlyList<VisioComment> Comments { get; }

        /// <summary>Whether every comment in the thread is resolved.</summary>
        public bool Done {
            get {
                foreach (VisioComment comment in Comments) {
                    if (!comment.Done) return false;
                }
                return true;
            }
        }
    }
}
