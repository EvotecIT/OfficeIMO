namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a legacy XLS cell comment parsed from NOTE/OBJ/TXO records.
    /// </summary>
    public sealed class LegacyXlsComment {
        /// <summary>
        /// Creates parsed legacy XLS comment metadata.
        /// </summary>
        public LegacyXlsComment(
            int row,
            int column,
            string text,
            string author,
            ushort objectId,
            bool visible,
            IReadOnlyList<LegacyXlsCommentFormattingRun>? formattingRuns = null,
            ushort? objectType = null,
            ushort? objectFlags = null,
            LegacyXlsDrawingAnchor? anchor = null) {
            Row = row;
            Column = column;
            Text = text;
            Author = author;
            ObjectId = objectId;
            Visible = visible;
            FormattingRuns = formattingRuns ?? Array.Empty<LegacyXlsCommentFormattingRun>();
            ObjectType = objectType;
            ObjectTypeKind = LegacyXlsDrawingObjectMetadata.TryGetObjectTypeKind(objectType);
            ObjectTypeName = ObjectTypeKind?.ToString() ?? (objectType.HasValue ? $"ObjectType:0x{objectType.Value:X4}" : null);
            ObjectFlags = objectFlags;
            ObjectFlagNames = objectFlags.HasValue ? LegacyXlsDrawingObjectMetadata.GetObjectFlagNames(objectFlags.Value) : Array.Empty<string>();
            Anchor = anchor;
        }

        /// <summary>Gets the 1-based row containing the comment.</summary>
        public int Row { get; }

        /// <summary>Gets the 1-based column containing the comment.</summary>
        public int Column { get; }

        /// <summary>Gets the plain comment text.</summary>
        public string Text { get; }

        /// <summary>Gets the comment author.</summary>
        public string Author { get; }

        /// <summary>Gets the BIFF object id associated with the comment.</summary>
        public ushort ObjectId { get; }

        /// <summary>Gets whether the source comment was configured as always visible.</summary>
        public bool Visible { get; }

        /// <summary>Gets rich text run boundaries and font indexes from the source TxO records.</summary>
        public IReadOnlyList<LegacyXlsCommentFormattingRun> FormattingRuns { get; }

        /// <summary>Gets the decoded OBJ common-object type identifier, when present.</summary>
        public ushort? ObjectType { get; }

        /// <summary>Gets the decoded OBJ common-object type, when the identifier is known.</summary>
        public LegacyXlsDrawingObjectType? ObjectTypeKind { get; }

        /// <summary>Gets a stable display name for the decoded OBJ common-object type, or a hexadecimal fallback for unknown types.</summary>
        public string? ObjectTypeName { get; }

        /// <summary>Gets the decoded OBJ common-object flag bitfield, when present.</summary>
        public ushort? ObjectFlags { get; }

        /// <summary>Gets stable names for the defined common-object flags set on the comment OBJ record.</summary>
        public IReadOnlyList<string> ObjectFlagNames { get; }

        /// <summary>Gets the OfficeArt client anchor associated with the comment object, when present.</summary>
        public LegacyXlsDrawingAnchor? Anchor { get; }

        /// <summary>Gets whether the comment has preserved OfficeArt client-anchor geometry.</summary>
        public bool HasAnchor => Anchor != null;

        /// <summary>Gets whether the comment object is locked.</summary>
        public bool IsObjectLocked => HasObjectFlag(0x0001);

        /// <summary>Gets whether the comment object is intended to be printed.</summary>
        public bool IsObjectPrintable => HasObjectFlag(0x0010);

        private bool HasObjectFlag(ushort mask) {
            return ObjectFlags.HasValue && (ObjectFlags.Value & mask) != 0;
        }
    }
}
