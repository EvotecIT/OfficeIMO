namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes the HFPicture wrapper used for legacy XLS header and footer pictures.
    /// </summary>
    public sealed class LegacyXlsHeaderFooterPicture {
        /// <summary>
        /// Creates header/footer picture metadata.
        /// </summary>
        public LegacyXlsHeaderFooterPicture(
            ushort wrappedRecordType,
            ushort futureRecordFlags,
            byte flags,
            byte reserved,
            int drawingByteCount) {
            if (drawingByteCount < 0) {
                throw new ArgumentOutOfRangeException(nameof(drawingByteCount));
            }

            WrappedRecordType = wrappedRecordType;
            FutureRecordFlags = futureRecordFlags;
            Flags = flags;
            Reserved = reserved;
            DrawingByteCount = drawingByteCount;
        }

        /// <summary>Gets the BIFF record type stored in the HFPicture future-record header.</summary>
        public ushort WrappedRecordType { get; }

        /// <summary>Gets the future-record flags from the HFPicture header.</summary>
        public ushort FutureRecordFlags { get; }

        /// <summary>Gets whether the future-record header matches the HFPicture BIFF record type.</summary>
        public bool HasMatchingFutureRecordHeader => WrappedRecordType == 0x0866;

        /// <summary>Gets the HFPicture drawing flags byte.</summary>
        public byte Flags { get; }

        /// <summary>Gets whether the payload declares an OfficeArtDgContainer.</summary>
        public bool IsDrawing => (Flags & 0x01) != 0;

        /// <summary>Gets whether the payload declares an OfficeArtDggContainer.</summary>
        public bool IsDrawingGroup => (Flags & 0x02) != 0;

        /// <summary>Gets whether this record continues a previous HFPicture record.</summary>
        public bool IsContinuation => (Flags & 0x04) != 0;

        /// <summary>Gets whether the drawing and drawing-group flags form a valid exclusive pair.</summary>
        public bool HasValidDrawingKind => IsDrawing != IsDrawingGroup;

        /// <summary>Gets the reserved byte that follows the HFPicture drawing flags.</summary>
        public byte Reserved { get; }

        /// <summary>Gets whether the reserved byte is clear.</summary>
        public bool HasClearReservedByte => Reserved == 0;

        /// <summary>Gets the embedded OfficeArt payload byte count after the HFPicture wrapper.</summary>
        public int DrawingByteCount { get; }

        /// <summary>Gets a stable state name for the wrapped drawing kind.</summary>
        public string DrawingKindName {
            get {
                if (IsDrawing && !IsDrawingGroup) {
                    return "Drawing";
                }

                if (IsDrawingGroup && !IsDrawing) {
                    return "DrawingGroup";
                }

                return "Invalid";
            }
        }

        /// <summary>Gets a stable state name for HFPicture continuation handling.</summary>
        public string ContinuationState => IsContinuation ? "Continuation" : "First";

        /// <summary>Gets a stable state name for the decoded header and flag combination.</summary>
        public string HeaderState {
            get {
                if (!HasMatchingFutureRecordHeader) {
                    return "MismatchedFutureRecordHeader";
                }

                if (!HasValidDrawingKind) {
                    return "InvalidDrawingFlags";
                }

                return HasClearReservedByte ? "Complete" : "ReservedNonZero";
            }
        }
    }
}
