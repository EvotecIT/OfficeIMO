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
            ushort? objectFlags = null) {
            Row = row;
            Column = column;
            Text = text;
            Author = author;
            ObjectId = objectId;
            Visible = visible;
            FormattingRuns = formattingRuns ?? Array.Empty<LegacyXlsCommentFormattingRun>();
            ObjectType = objectType;
            ObjectTypeKind = TryGetObjectTypeKind(objectType);
            ObjectTypeName = ObjectTypeKind?.ToString() ?? (objectType.HasValue ? $"ObjectType:0x{objectType.Value:X4}" : null);
            ObjectFlags = objectFlags;
            ObjectFlagNames = objectFlags.HasValue ? GetObjectFlagNames(objectFlags.Value) : Array.Empty<string>();
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

        /// <summary>Gets whether the comment object is locked.</summary>
        public bool IsObjectLocked => HasObjectFlag(0x0001);

        /// <summary>Gets whether the comment object is intended to be printed.</summary>
        public bool IsObjectPrintable => HasObjectFlag(0x0010);

        private bool HasObjectFlag(ushort mask) {
            return ObjectFlags.HasValue && (ObjectFlags.Value & mask) != 0;
        }

        private static LegacyXlsDrawingObjectType? TryGetObjectTypeKind(ushort? objectType) {
            if (!objectType.HasValue) {
                return null;
            }

            return objectType.Value switch {
                0x0000 => LegacyXlsDrawingObjectType.Group,
                0x0001 => LegacyXlsDrawingObjectType.Line,
                0x0002 => LegacyXlsDrawingObjectType.Rectangle,
                0x0003 => LegacyXlsDrawingObjectType.Oval,
                0x0004 => LegacyXlsDrawingObjectType.Arc,
                0x0005 => LegacyXlsDrawingObjectType.Chart,
                0x0006 => LegacyXlsDrawingObjectType.Text,
                0x0007 => LegacyXlsDrawingObjectType.Button,
                0x0008 => LegacyXlsDrawingObjectType.Picture,
                0x0009 => LegacyXlsDrawingObjectType.Polygon,
                0x000B => LegacyXlsDrawingObjectType.Checkbox,
                0x000C => LegacyXlsDrawingObjectType.RadioButton,
                0x000D => LegacyXlsDrawingObjectType.EditBox,
                0x000E => LegacyXlsDrawingObjectType.Label,
                0x000F => LegacyXlsDrawingObjectType.DialogBox,
                0x0010 => LegacyXlsDrawingObjectType.SpinControl,
                0x0011 => LegacyXlsDrawingObjectType.Scrollbar,
                0x0012 => LegacyXlsDrawingObjectType.List,
                0x0013 => LegacyXlsDrawingObjectType.GroupBox,
                0x0014 => LegacyXlsDrawingObjectType.DropdownList,
                0x0019 => LegacyXlsDrawingObjectType.Note,
                0x001E => LegacyXlsDrawingObjectType.OfficeArtObject,
                _ => null
            };
        }

        private static IReadOnlyList<string> GetObjectFlagNames(ushort flags) {
            var names = new List<string>();
            if ((flags & 0x0001) != 0) names.Add("Locked");
            if ((flags & 0x0004) != 0) names.Add("DefaultSize");
            if ((flags & 0x0008) != 0) names.Add("Published");
            if ((flags & 0x0010) != 0) names.Add("Printable");
            if ((flags & 0x0080) != 0) names.Add("Disabled");
            if ((flags & 0x0100) != 0) names.Add("UiObject");
            if ((flags & 0x0200) != 0) names.Add("RecalculateOnLoad");
            if ((flags & 0x1000) != 0) names.Add("AlwaysRecalculate");
            return names;
        }
    }
}
