namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a preserve-only drawing or object BIFF record discovered during legacy XLS import.
    /// </summary>
    public sealed class LegacyXlsDrawingRecord {
        /// <summary>
        /// Creates drawing BIFF record metadata.
        /// </summary>
        public LegacyXlsDrawingRecord(
            LegacyXlsDrawingRecordKind kind,
            string recordName,
            string? sheetName,
            int recordOffset,
            ushort recordType,
            int payloadLength,
            ushort? objectType = null,
            ushort? objectId = null,
            ushort? escherRecordType = null,
            ushort? escherRecordInstance = null,
            byte? escherRecordVersion = null,
            uint? escherPayloadLength = null,
            LegacyXlsDrawingObjectType? objectTypeKind = null,
            LegacyXlsDrawingEscherRecordType? escherRecordTypeKind = null,
            ushort? objectFlags = null,
            IReadOnlyList<LegacyXlsDrawingBlipStoreEntry>? blipStoreEntries = null) {
            if (payloadLength < 0) {
                throw new ArgumentOutOfRangeException(nameof(payloadLength));
            }

            Kind = kind;
            RecordName = recordName ?? throw new ArgumentNullException(nameof(recordName));
            SheetName = sheetName;
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
            ObjectType = objectType;
            ObjectId = objectId;
            ObjectTypeKind = objectTypeKind ?? TryGetObjectTypeKind(objectType);
            ObjectTypeName = ObjectTypeKind?.ToString() ?? (objectType.HasValue ? $"ObjectType:0x{objectType.Value:X4}" : null);
            ObjectFlags = objectFlags;
            ObjectFlagNames = objectFlags.HasValue ? GetObjectFlagNames(objectFlags.Value) : Array.Empty<string>();
            EscherRecordType = escherRecordType;
            EscherRecordTypeKind = escherRecordTypeKind ?? TryGetEscherRecordTypeKind(escherRecordType);
            EscherRecordTypeName = EscherRecordTypeKind?.ToString() ?? (escherRecordType.HasValue ? $"EscherRecordType:0x{escherRecordType.Value:X4}" : null);
            EscherRecordInstance = escherRecordInstance;
            EscherRecordVersion = escherRecordVersion;
            EscherPayloadLength = escherPayloadLength;
            BlipStoreEntries = blipStoreEntries?.ToArray() ?? Array.Empty<LegacyXlsDrawingBlipStoreEntry>();
        }

        /// <summary>Gets the shallow drawing record category.</summary>
        public LegacyXlsDrawingRecordKind Kind { get; }

        /// <summary>Gets the BIFF record name.</summary>
        public string RecordName { get; }

        /// <summary>Gets the worksheet or chart sheet name associated with the record, when known.</summary>
        public string? SheetName { get; }

        /// <summary>Gets the byte offset of the BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type identifier.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the BIFF record payload length in bytes.</summary>
        public int PayloadLength { get; }

        /// <summary>Gets the decoded OBJ common-object type identifier, when present.</summary>
        public ushort? ObjectType { get; }

        /// <summary>Gets the decoded OBJ common-object type, when the identifier is known.</summary>
        public LegacyXlsDrawingObjectType? ObjectTypeKind { get; }

        /// <summary>Gets a stable display name for the decoded OBJ common-object type, or a hexadecimal fallback for unknown types.</summary>
        public string? ObjectTypeName { get; }

        /// <summary>Gets the decoded OBJ object identifier, when present.</summary>
        public ushort? ObjectId { get; }

        /// <summary>Gets the decoded OBJ common-object flag bitfield, when present.</summary>
        public ushort? ObjectFlags { get; }

        /// <summary>Gets stable names for the defined common-object flags set on this OBJ record.</summary>
        public IReadOnlyList<string> ObjectFlagNames { get; }

        /// <summary>Gets whether the object is locked.</summary>
        public bool IsObjectLocked => HasObjectFlag(0x0001);

        /// <summary>Gets whether the application is expected to choose the object size.</summary>
        public bool UsesDefaultObjectSize => HasObjectFlag(0x0004);

        /// <summary>Gets whether this chart object is expected to be published with the sheet.</summary>
        public bool IsObjectPublished => HasObjectFlag(0x0008);

        /// <summary>Gets whether the object is intended to be printed.</summary>
        public bool IsObjectPrintable => HasObjectFlag(0x0010);

        /// <summary>Gets whether the object is disabled.</summary>
        public bool IsObjectDisabled => HasObjectFlag(0x0080);

        /// <summary>Gets whether this is an application-inserted UI object.</summary>
        public bool IsUiObject => HasObjectFlag(0x0100);

        /// <summary>Gets whether the object is expected to recalculate from its linked range on load.</summary>
        public bool RecalculatesObjectOnLoad => HasObjectFlag(0x0200);

        /// <summary>Gets whether the object is expected to recalculate whenever its linked range changes.</summary>
        public bool AlwaysRecalculatesObject => HasObjectFlag(0x1000);

        /// <summary>Gets the top-level Escher record type from MsoDrawing payloads, when present.</summary>
        public ushort? EscherRecordType { get; }

        /// <summary>Gets the top-level Escher OfficeArt record type, when the identifier is known.</summary>
        public LegacyXlsDrawingEscherRecordType? EscherRecordTypeKind { get; }

        /// <summary>Gets a stable display name for the top-level Escher record type, or a hexadecimal fallback for unknown types.</summary>
        public string? EscherRecordTypeName { get; }

        /// <summary>Gets the top-level Escher record instance from MsoDrawing payloads, when present.</summary>
        public ushort? EscherRecordInstance { get; }

        /// <summary>Gets the top-level Escher record version from MsoDrawing payloads, when present.</summary>
        public byte? EscherRecordVersion { get; }

        /// <summary>Gets the declared top-level Escher payload length from MsoDrawing payloads, when present.</summary>
        public uint? EscherPayloadLength { get; }

        /// <summary>Gets preserve-only OfficeArt FBSE image-store entries discovered under this drawing record.</summary>
        public IReadOnlyList<LegacyXlsDrawingBlipStoreEntry> BlipStoreEntries { get; }

        /// <summary>Gets whether this drawing record contains any discovered image-store entries.</summary>
        public bool HasBlipStoreEntries => BlipStoreEntries.Count > 0;

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

        private bool HasObjectFlag(ushort mask) {
            return ObjectFlags.HasValue && (ObjectFlags.Value & mask) != 0;
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

        private static LegacyXlsDrawingEscherRecordType? TryGetEscherRecordTypeKind(ushort? recordType) {
            if (!recordType.HasValue) {
                return null;
            }

            return recordType.Value switch {
                0xF000 => LegacyXlsDrawingEscherRecordType.OfficeArtDggContainer,
                0xF001 => LegacyXlsDrawingEscherRecordType.OfficeArtBStoreContainer,
                0xF002 => LegacyXlsDrawingEscherRecordType.OfficeArtDgContainer,
                0xF003 => LegacyXlsDrawingEscherRecordType.OfficeArtSpgrContainer,
                0xF004 => LegacyXlsDrawingEscherRecordType.OfficeArtSpContainer,
                0xF005 => LegacyXlsDrawingEscherRecordType.OfficeArtSolverContainer,
                0xF006 => LegacyXlsDrawingEscherRecordType.OfficeArtFDGGBlock,
                0xF007 => LegacyXlsDrawingEscherRecordType.OfficeArtFBSE,
                0xF008 => LegacyXlsDrawingEscherRecordType.OfficeArtFDG,
                0xF009 => LegacyXlsDrawingEscherRecordType.OfficeArtFSPGR,
                0xF00A => LegacyXlsDrawingEscherRecordType.OfficeArtFSP,
                0xF00B => LegacyXlsDrawingEscherRecordType.OfficeArtFOPT,
                0xF00D => LegacyXlsDrawingEscherRecordType.OfficeArtFClientTextbox,
                0xF00F => LegacyXlsDrawingEscherRecordType.OfficeArtChildAnchor,
                0xF010 => LegacyXlsDrawingEscherRecordType.OfficeArtFClientAnchor,
                0xF011 => LegacyXlsDrawingEscherRecordType.OfficeArtFClientData,
                _ => null
            };
        }
    }
}
