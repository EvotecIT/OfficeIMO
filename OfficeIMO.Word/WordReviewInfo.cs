namespace OfficeIMO.Word {
    /// <summary>
    /// Identifies the document part where review metadata was found.
    /// </summary>
    public enum WordReviewLocationKind {
        /// <summary>The review item is in the main document body.</summary>
        Body,
        /// <summary>The review item is in a header part.</summary>
        Header,
        /// <summary>The review item is in a footer part.</summary>
        Footer,
        /// <summary>The review item is in a footnote part.</summary>
        Footnote,
        /// <summary>The review item is in an endnote part.</summary>
        Endnote
    }

    /// <summary>
    /// Identifies the tracked-change kind represented by a revision item.
    /// </summary>
    public enum WordReviewRevisionType {
        /// <summary>Inserted content.</summary>
        Insertion,
        /// <summary>Deleted content.</summary>
        Deletion,
        /// <summary>Moved-from content.</summary>
        MoveFrom,
        /// <summary>Moved-to content.</summary>
        MoveTo,
        /// <summary>Paragraph formatting change.</summary>
        ParagraphFormatting,
        /// <summary>Run formatting change.</summary>
        RunFormatting,
        /// <summary>Table formatting change.</summary>
        TableFormatting,
        /// <summary>Table row formatting change.</summary>
        TableRowFormatting,
        /// <summary>Table cell formatting change.</summary>
        TableCellFormatting,
        /// <summary>Section formatting change.</summary>
        SectionFormatting,
        /// <summary>A tracked-change element OfficeIMO can count but not classify more precisely.</summary>
        Unknown
    }

    /// <summary>
    /// Read-only review metadata discovered in a Word document.
    /// </summary>
    public sealed class WordReviewInfo {
        internal WordReviewInfo(
            IReadOnlyList<WordCommentInfo> comments,
            IReadOnlyList<WordRevisionInfo> revisions,
            IReadOnlyList<string> unsupportedMetadata) {
            Comments = comments.ToArray();
            Revisions = revisions.ToArray();
            UnsupportedMetadata = unsupportedMetadata.ToArray();
        }

        /// <summary>Gets classic comments and parsed comment thread metadata.</summary>
        public IReadOnlyList<WordCommentInfo> Comments { get; }

        /// <summary>Gets tracked revisions discovered in document content parts.</summary>
        public IReadOnlyList<WordRevisionInfo> Revisions { get; }

        /// <summary>Gets modern review package metadata that is detected but not yet fully parsed.</summary>
        public IReadOnlyList<string> UnsupportedMetadata { get; }

        /// <summary>Gets whether any review metadata was discovered.</summary>
        public bool HasReviewMetadata => Comments.Count > 0 || Revisions.Count > 0 || UnsupportedMetadata.Count > 0;

        /// <summary>Gets the number of parsed top-level and reply comments.</summary>
        public int CommentCount => Comments.Count;

        /// <summary>Gets the number of parsed tracked revisions.</summary>
        public int RevisionCount => Revisions.Count;

        /// <summary>Gets the number of comments marked as resolved where the metadata is available.</summary>
        public int ResolvedCommentCount => Comments.Count(comment => comment.IsResolved == true);

        /// <summary>Gets the number of comments that are known unresolved where the metadata is available.</summary>
        public int UnresolvedCommentCount => Comments.Count(comment => comment.IsResolved == false);

        /// <summary>Gets the number of comments that are replies to another comment.</summary>
        public int ReplyCount => Comments.Count(comment => comment.IsReply);

        /// <summary>
        /// Returns revisions authored by the specified person using case-insensitive matching.
        /// </summary>
        /// <param name="author">Author name to match.</param>
        public IReadOnlyList<WordRevisionInfo> GetRevisionsByAuthor(string author) {
            if (string.IsNullOrWhiteSpace(author)) {
                return Array.Empty<WordRevisionInfo>();
            }

            return Revisions
                .Where(revision => string.Equals(revision.Author, author, StringComparison.OrdinalIgnoreCase))
                .ToArray();
        }
    }

    /// <summary>
    /// Read-only metadata for a Word comment.
    /// </summary>
    public sealed class WordCommentInfo {
        internal WordCommentInfo(
            int index,
            string? id,
            string? author,
            string? initials,
            DateTime? dateTime,
            string text,
            string? paraId,
            string? parentParaId,
            bool? isResolved,
            string targetText,
            WordReviewLocationKind? targetLocationKind,
            string? targetPartUri,
            bool isInTable,
            bool isInContentControl,
            bool isInTextBox,
            int documentOrder) {
            Index = index;
            Id = id;
            Author = author;
            Initials = initials;
            DateTime = dateTime;
            Text = text;
            ParaId = paraId;
            ParentParaId = parentParaId;
            IsResolved = isResolved;
            TargetText = targetText;
            TargetLocationKind = targetLocationKind;
            TargetPartUri = targetPartUri;
            IsInTable = isInTable;
            IsInContentControl = isInContentControl;
            IsInTextBox = isInTextBox;
            DocumentOrder = documentOrder;
        }

        /// <summary>Gets the deterministic index in document comment order.</summary>
        public int Index { get; }

        /// <summary>Gets the Word comment id.</summary>
        public string? Id { get; }

        /// <summary>Gets the comment author.</summary>
        public string? Author { get; }

        /// <summary>Gets the author's initials.</summary>
        public string? Initials { get; }

        /// <summary>Gets the comment timestamp when present.</summary>
        public DateTime? DateTime { get; }

        /// <summary>Gets the plain text stored in the comment body.</summary>
        public string Text { get; }

        /// <summary>Gets the paragraph id used by modern comment metadata.</summary>
        public string? ParaId { get; }

        /// <summary>Gets the parent paragraph id when this comment is a reply.</summary>
        public string? ParentParaId { get; }

        /// <summary>Gets whether the comment is a reply to another comment.</summary>
        public bool IsReply => !string.IsNullOrWhiteSpace(ParentParaId);

        /// <summary>Gets whether the comment is marked resolved; null means the document did not expose this metadata.</summary>
        public bool? IsResolved { get; }

        /// <summary>Gets the plain text targeted by the comment range where it can be found.</summary>
        public string TargetText { get; }

        /// <summary>Gets the target part category when the comment reference can be found.</summary>
        public WordReviewLocationKind? TargetLocationKind { get; }

        /// <summary>Gets the package part URI that contains the comment target.</summary>
        public string? TargetPartUri { get; }

        /// <summary>Gets whether the comment target is inside a table.</summary>
        public bool IsInTable { get; }

        /// <summary>Gets whether the comment target is inside a content control.</summary>
        public bool IsInContentControl { get; }

        /// <summary>Gets whether the comment target is inside a text box.</summary>
        public bool IsInTextBox { get; }

        internal int DocumentOrder { get; }
    }

    /// <summary>
    /// Read-only metadata for a tracked revision.
    /// </summary>
    public sealed class WordRevisionInfo {
        internal WordRevisionInfo(
            int index,
            WordReviewRevisionType revisionType,
            string elementName,
            string? id,
            string? author,
            DateTime? dateTime,
            string affectedText,
            string locationText,
            WordReviewLocationKind locationKind,
            string partUri,
            bool isInTable,
            bool isInContentControl,
            bool isInTextBox,
            int documentOrder) {
            Index = index;
            RevisionType = revisionType;
            ElementName = elementName;
            Id = id;
            Author = author;
            DateTime = dateTime;
            AffectedText = affectedText;
            LocationText = locationText;
            LocationKind = locationKind;
            PartUri = partUri;
            IsInTable = isInTable;
            IsInContentControl = isInContentControl;
            IsInTextBox = isInTextBox;
            DocumentOrder = documentOrder;
        }

        /// <summary>Gets the deterministic index in revision scan order.</summary>
        public int Index { get; }

        /// <summary>Gets the tracked-change classification.</summary>
        public WordReviewRevisionType RevisionType { get; }

        /// <summary>Gets the source Open XML element local name for diagnostics.</summary>
        public string ElementName { get; }

        /// <summary>Gets the revision id when present.</summary>
        public string? Id { get; }

        /// <summary>Gets the revision author when present.</summary>
        public string? Author { get; }

        /// <summary>Gets the revision timestamp when present.</summary>
        public DateTime? DateTime { get; }

        /// <summary>Gets the text directly affected by the revision when it can be represented as plain text.</summary>
        public string AffectedText { get; }

        /// <summary>Gets nearby paragraph or parent text to help callers locate formatting-only revisions.</summary>
        public string LocationText { get; }

        /// <summary>Gets the document part category that contains the revision.</summary>
        public WordReviewLocationKind LocationKind { get; }

        /// <summary>Gets the package part URI that contains the revision.</summary>
        public string PartUri { get; }

        /// <summary>Gets whether the revision is inside a table.</summary>
        public bool IsInTable { get; }

        /// <summary>Gets whether the revision is inside a content control.</summary>
        public bool IsInContentControl { get; }

        /// <summary>Gets whether the revision is inside a text box.</summary>
        public bool IsInTextBox { get; }

        internal int DocumentOrder { get; }
    }
}
