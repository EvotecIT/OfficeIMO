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
            IReadOnlyList<LegacyXlsDrawingBlipStoreEntry>? blipStoreEntries = null,
            IReadOnlyList<LegacyXlsDrawingShape>? shapeEntries = null,
            IReadOnlyList<LegacyXlsDrawingAnchor>? anchorEntries = null,
            IReadOnlyList<LegacyXlsDrawingChildAnchor>? childAnchorEntries = null,
            IReadOnlyList<LegacyXlsDrawingOfficeArtRecord>? officeArtRecords = null,
            IReadOnlyList<LegacyXlsDrawingGroupBlock>? drawingGroupBlocks = null,
            IReadOnlyList<LegacyXlsDrawingGroupInfo>? drawingGroupInfos = null,
            IReadOnlyList<LegacyXlsDrawingShapeProperty>? shapeProperties = null,
            IReadOnlyList<LegacyXlsDrawingObjectSubRecord>? objectSubRecords = null,
            LegacyXlsDrawingFutureRecordHeader? futureRecordHeader = null,
            LegacyXlsDrawingTextObject? textObject = null,
            LegacyXlsHeaderFooterPicture? headerFooterPicture = null,
            bool officeArtPayloadFullyTraversed = false) {
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
            ObjectTypeKind = objectTypeKind ?? LegacyXlsDrawingObjectMetadata.TryGetObjectTypeKind(objectType);
            ObjectTypeName = ObjectTypeKind?.ToString() ?? (objectType.HasValue ? $"ObjectType:0x{objectType.Value:X4}" : null);
            ObjectFlags = objectFlags;
            ObjectFlagNames = objectFlags.HasValue ? LegacyXlsDrawingObjectMetadata.GetObjectFlagNames(objectFlags.Value) : Array.Empty<string>();
            EscherRecordType = escherRecordType;
            EscherRecordTypeKind = escherRecordTypeKind ?? LegacyXlsDrawingEscherRecordTypeDecoder.TryGetKind(escherRecordType);
            EscherRecordTypeName = escherRecordType.HasValue ? LegacyXlsDrawingEscherRecordTypeDecoder.GetName(escherRecordType.Value) : null;
            EscherRecordInstance = escherRecordInstance;
            EscherRecordVersion = escherRecordVersion;
            EscherPayloadLength = escherPayloadLength;
            BlipStoreEntries = blipStoreEntries?.ToArray() ?? Array.Empty<LegacyXlsDrawingBlipStoreEntry>();
            ShapeEntries = shapeEntries?.ToArray() ?? Array.Empty<LegacyXlsDrawingShape>();
            AnchorEntries = anchorEntries?.ToArray() ?? Array.Empty<LegacyXlsDrawingAnchor>();
            ChildAnchorEntries = childAnchorEntries?.ToArray() ?? Array.Empty<LegacyXlsDrawingChildAnchor>();
            OfficeArtRecords = officeArtRecords?.ToArray() ?? Array.Empty<LegacyXlsDrawingOfficeArtRecord>();
            DrawingGroupBlocks = drawingGroupBlocks?.ToArray() ?? Array.Empty<LegacyXlsDrawingGroupBlock>();
            DrawingGroupInfos = drawingGroupInfos?.ToArray() ?? Array.Empty<LegacyXlsDrawingGroupInfo>();
            ShapeProperties = shapeProperties?.ToArray() ?? Array.Empty<LegacyXlsDrawingShapeProperty>();
            ObjectSubRecords = objectSubRecords?.ToArray() ?? Array.Empty<LegacyXlsDrawingObjectSubRecord>();
            FutureRecordHeader = futureRecordHeader;
            TextObject = textObject;
            HeaderFooterPicture = headerFooterPicture;
            OfficeArtPayloadFullyTraversed = officeArtPayloadFullyTraversed;
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

        /// <summary>Gets preserve-only subrecord metadata discovered inside this OBJ record.</summary>
        public IReadOnlyList<LegacyXlsDrawingObjectSubRecord> ObjectSubRecords { get; }

        /// <summary>Gets whether this OBJ record contains discovered subrecords.</summary>
        public bool HasObjectSubRecords => ObjectSubRecords.Count > 0;

        /// <summary>Gets whether this OBJ record has complete common-object metadata that can be reported without treating the record as unsupported.</summary>
        public bool HasSupportedObjectMetadata =>
            Kind == LegacyXlsDrawingRecordKind.Object
            && ObjectType.HasValue
            && ObjectId.HasValue
            && ObjectFlags.HasValue
            && ObjectTypeKind.HasValue
            && ObjectSubRecords.Count > 0
            && ObjectSubRecords[0].SubRecordType == 0x0015
            && ObjectSubRecords[0].HasSupportedPayload
            && HasSupportedObjectSubRecordTerminator
            && ObjectSubRecords.All(subRecord => subRecord.HasSupportedPayload);

        private bool HasSupportedObjectSubRecordTerminator =>
            ObjectSubRecords.Any(subRecord => subRecord.SubRecordType == 0x0000 && subRecord.IsComplete)
            || ObjectSubRecords[ObjectSubRecords.Count - 1].RequiresContinuation;

        /// <summary>Gets the decoded future-record stream header, when this drawing record uses that wrapper.</summary>
        public LegacyXlsDrawingFutureRecordHeader? FutureRecordHeader { get; }

        /// <summary>Gets whether this drawing record has a decoded future-record stream header.</summary>
        public bool HasFutureRecordHeader => FutureRecordHeader != null;

        /// <summary>Gets whether this future drawing stream has complete metadata that can be reported without treating the record as unsupported.</summary>
        public bool HasSupportedFutureDrawingStreamMetadata =>
            (Kind == LegacyXlsDrawingRecordKind.ShapePropertiesStream
                || Kind == LegacyXlsDrawingRecordKind.TextPropertiesStream
                || Kind == LegacyXlsDrawingRecordKind.RichTextStream)
            && FutureRecordHeader != null
            && FutureRecordHeader.WrappedRecordType == RecordType
            && FutureRecordHeader.HasCompleteRangeReference;

        /// <summary>Gets whether this drawing record is decoded enough to be represented as supported import metadata.</summary>
        public bool HasSupportedDrawingMetadata =>
            HasSupportedObjectMetadata
            || HasSupportedOfficeArtMetadata
            || HasSupportedPartialOfficeArtContainerMetadata
            || HasSupportedOfficeArtClientTextboxMetadata
            || HasSupportedFutureDrawingStreamMetadata
            || HasSupportedHeaderFooterPictureMetadata;

        /// <summary>Gets decoded TxO text-object header metadata, when this record is a TxO record.</summary>
        public LegacyXlsDrawingTextObject? TextObject { get; }

        /// <summary>Gets whether this record has decoded TxO text-object header metadata.</summary>
        public bool HasTextObject => TextObject != null;

        /// <summary>Gets decoded HFPicture wrapper metadata, when this record is a header/footer picture.</summary>
        public LegacyXlsHeaderFooterPicture? HeaderFooterPicture { get; }

        /// <summary>Gets whether this record has decoded HFPicture wrapper metadata.</summary>
        public bool HasHeaderFooterPicture => HeaderFooterPicture != null;

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

        /// <summary>Gets preserve-only OfficeArt shape entries discovered under this drawing record.</summary>
        public IReadOnlyList<LegacyXlsDrawingShape> ShapeEntries { get; }

        /// <summary>Gets whether this drawing record contains any discovered shape entries.</summary>
        public bool HasShapeEntries => ShapeEntries.Count > 0;

        /// <summary>Gets preserve-only OfficeArt client anchors discovered under this drawing record.</summary>
        public IReadOnlyList<LegacyXlsDrawingAnchor> AnchorEntries { get; }

        /// <summary>Gets whether this drawing record contains any discovered client anchors.</summary>
        public bool HasAnchorEntries => AnchorEntries.Count > 0;

        /// <summary>Gets preserve-only OfficeArt child anchors discovered under this drawing record.</summary>
        public IReadOnlyList<LegacyXlsDrawingChildAnchor> ChildAnchorEntries { get; }

        /// <summary>Gets whether this drawing record contains any discovered child anchors.</summary>
        public bool HasChildAnchorEntries => ChildAnchorEntries.Count > 0;

        /// <summary>Gets preserve-only OfficeArt record headers discovered while traversing this drawing record.</summary>
        public IReadOnlyList<LegacyXlsDrawingOfficeArtRecord> OfficeArtRecords { get; }

        /// <summary>Gets whether this drawing record contains discovered OfficeArt record headers.</summary>
        public bool HasOfficeArtRecords => OfficeArtRecords.Count > 0;

        /// <summary>Gets whether the OfficeArt payload was traversed without truncation or trailing bytes.</summary>
        public bool OfficeArtPayloadFullyTraversed { get; }

        /// <summary>Gets whether this MsoDrawing record has complete OfficeArt metadata that can be reported without treating the record as unsupported.</summary>
        public bool HasSupportedOfficeArtMetadata =>
            Kind == LegacyXlsDrawingRecordKind.Drawing
            && EscherRecordType.HasValue
            && EscherPayloadLength.HasValue
            && OfficeArtPayloadFullyTraversed
            && OfficeArtRecords.Count > 0
            && (ShapeEntries.Count > 0
                || AnchorEntries.Count > 0
                || ChildAnchorEntries.Count > 0
                || ShapeProperties.Count > 0
                || DrawingGroupInfos.Count > 0
                || BlipStoreEntries.Count > 0);

        /// <summary>Gets whether this split MsoDrawing record has enough OfficeArt container metadata to report as supported.</summary>
        public bool HasSupportedPartialOfficeArtContainerMetadata =>
            Kind == LegacyXlsDrawingRecordKind.Drawing
            && (EscherRecordTypeKind == LegacyXlsDrawingEscherRecordType.OfficeArtDgContainer
                || EscherRecordTypeKind == LegacyXlsDrawingEscherRecordType.OfficeArtSpContainer)
            && EscherPayloadLength.HasValue
            && !OfficeArtPayloadFullyTraversed
            && OfficeArtRecords.Count > 0
            && OfficeArtRecords[0].RecordTypeKind == EscherRecordTypeKind
            && OfficeArtRecords[0].IsContainer
            && (ShapeEntries.Count > 0
                || AnchorEntries.Count > 0
                || ChildAnchorEntries.Count > 0
                || ShapeProperties.Count > 0
                || DrawingGroupInfos.Count > 0);

        /// <summary>Gets whether this MsoDrawing record contains a complete OfficeArt client-textbox marker.</summary>
        public bool HasSupportedOfficeArtClientTextboxMetadata =>
            Kind == LegacyXlsDrawingRecordKind.Drawing
            && EscherRecordTypeKind == LegacyXlsDrawingEscherRecordType.OfficeArtFClientTextbox
            && EscherPayloadLength == 0
            && OfficeArtPayloadFullyTraversed
            && OfficeArtRecords.Count == 1;

        /// <summary>Gets whether this HFPicture record has complete header/footer picture OfficeArt metadata.</summary>
        public bool HasSupportedHeaderFooterPictureMetadata =>
            Kind == LegacyXlsDrawingRecordKind.HeaderFooterPicture
            && HeaderFooterPicture != null
            && HeaderFooterPicture.HasMatchingFutureRecordHeader
            && HeaderFooterPicture.HasValidDrawingKind
            && !HeaderFooterPicture.IsContinuation
            && EscherRecordType.HasValue
            && EscherPayloadLength.HasValue
            && OfficeArtPayloadFullyTraversed
            && OfficeArtRecords.Count > 0
            && ((HeaderFooterPicture.IsDrawing
                    && EscherRecordTypeKind == LegacyXlsDrawingEscherRecordType.OfficeArtDgContainer
                    && (ShapeEntries.Count > 0
                        || AnchorEntries.Count > 0
                        || ChildAnchorEntries.Count > 0
                        || ShapeProperties.Count > 0
                        || DrawingGroupInfos.Count > 0))
                || (HeaderFooterPicture.IsDrawingGroup
                    && EscherRecordTypeKind == LegacyXlsDrawingEscherRecordType.OfficeArtDggContainer
                    && (DrawingGroupBlocks.Count > 0
                        || BlipStoreEntries.Count > 0)));

        /// <summary>Gets preserve-only document-wide OfficeArtFDGGBlock drawing metadata discovered under this record.</summary>
        public IReadOnlyList<LegacyXlsDrawingGroupBlock> DrawingGroupBlocks { get; }

        /// <summary>Gets whether this drawing record contains document-wide OfficeArtFDGGBlock metadata.</summary>
        public bool HasDrawingGroupBlocks => DrawingGroupBlocks.Count > 0;

        /// <summary>Gets preserve-only per-drawing OfficeArtFDG metadata discovered under this record.</summary>
        public IReadOnlyList<LegacyXlsDrawingGroupInfo> DrawingGroupInfos { get; }

        /// <summary>Gets whether this drawing record contains per-drawing OfficeArtFDG metadata.</summary>
        public bool HasDrawingGroupInfos => DrawingGroupInfos.Count > 0;

        /// <summary>Gets preserve-only OfficeArtFOPT shape properties discovered under this drawing record.</summary>
        public IReadOnlyList<LegacyXlsDrawingShapeProperty> ShapeProperties { get; }

        /// <summary>Gets whether this drawing record contains discovered OfficeArtFOPT shape properties.</summary>
        public bool HasShapeProperties => ShapeProperties.Count > 0;

        /// <summary>
        /// Creates a copy of this record with updated TxO text-object metadata.
        /// </summary>
        internal LegacyXlsDrawingRecord WithTextObject(LegacyXlsDrawingTextObject textObject) {
            if (textObject == null) {
                throw new ArgumentNullException(nameof(textObject));
            }

            return new LegacyXlsDrawingRecord(
                Kind,
                RecordName,
                SheetName,
                RecordOffset,
                RecordType,
                PayloadLength,
                ObjectType,
                ObjectId,
                EscherRecordType,
                EscherRecordInstance,
                EscherRecordVersion,
                EscherPayloadLength,
                ObjectTypeKind,
                EscherRecordTypeKind,
                ObjectFlags,
                BlipStoreEntries,
                ShapeEntries,
                AnchorEntries,
                ChildAnchorEntries,
                OfficeArtRecords,
                DrawingGroupBlocks,
                DrawingGroupInfos,
                ShapeProperties,
                ObjectSubRecords,
                FutureRecordHeader,
                textObject,
                HeaderFooterPicture,
                OfficeArtPayloadFullyTraversed);
        }

        private bool HasObjectFlag(ushort mask) {
            return ObjectFlags.HasValue && (ObjectFlags.Value & mask) != 0;
        }

    }
}
