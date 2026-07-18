using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Decodes SlideAtom layout metadata and PlaceholderAtom identity.</summary>
    internal static class LegacyPptLayoutReader {
        internal static bool TryReadSlideAtom(LegacyPptRecord? record,
            out LegacyPptSlideAtomData data) {
            if (record == null || record.PayloadLength < 24) {
                data = default;
                return false;
            }

            uint rawLayoutType = record.ReadUInt32(0);
            LegacyPptSlideLayoutType? layout = IsDefinedLayout(rawLayoutType)
                ? (LegacyPptSlideLayoutType)rawLayoutType
                : null;
            var placeholderTypes = new LegacyPptPlaceholderKind[8];
            bool hasInvalidPlaceholderType = false;
            for (int index = 0; index < placeholderTypes.Length; index++) {
                byte value = record.ReadByte(4 + index);
                if (Enum.IsDefined(typeof(LegacyPptPlaceholderKind), value)) {
                    placeholderTypes[index] = (LegacyPptPlaceholderKind)value;
                } else {
                    placeholderTypes[index] = LegacyPptPlaceholderKind.None;
                    hasInvalidPlaceholderType = true;
                }
            }

            ushort flags = record.ReadUInt16(20);
            data = new LegacyPptSlideAtomData(rawLayoutType, layout, placeholderTypes,
                record.ReadUInt32(12), record.ReadUInt32(16), flags,
                hasInvalidPlaceholderType, record.PayloadLength != 24);
            return true;
        }

        internal static LegacyPptPlaceholder? ReadPlaceholder(LegacyPptRecord? record,
            out LegacyPptPlaceholderReadStatus status) {
            if (record == null) {
                status = LegacyPptPlaceholderReadStatus.Absent;
                return null;
            }
            if (record.PayloadLength != 8) {
                status = LegacyPptPlaceholderReadStatus.Invalid;
                return null;
            }

            int position = record.ReadInt32(0);
            if (position == -1) {
                status = LegacyPptPlaceholderReadStatus.NotPlaceholder;
                return null;
            }
            byte rawKind = record.ReadByte(4);
            byte rawSize = record.ReadByte(5);
            if (position < 0 || rawKind == 0
                || !Enum.IsDefined(typeof(LegacyPptPlaceholderKind), rawKind)
                || !Enum.IsDefined(typeof(LegacyPptPlaceholderSize), rawSize)) {
                status = LegacyPptPlaceholderReadStatus.Invalid;
                return null;
            }

            status = LegacyPptPlaceholderReadStatus.Decoded;
            return new LegacyPptPlaceholder(position, (LegacyPptPlaceholderKind)rawKind,
                (LegacyPptPlaceholderSize)rawSize);
        }

        private static bool IsDefinedLayout(uint value) => value == 0 || value == 1
            || value == 2 || value == 7 || value == 8 || value == 9 || value == 10
            || value == 11 || value == 13 || value == 14 || value == 15 || value == 16
            || value == 17 || value == 18;
    }

    internal readonly struct LegacyPptSlideAtomData {
        internal LegacyPptSlideAtomData(uint rawLayoutType, LegacyPptSlideLayoutType? layout,
            IReadOnlyList<LegacyPptPlaceholderKind> placeholderTypes, uint masterId,
            uint notesId, ushort flags, bool hasInvalidPlaceholderType,
            bool hasInvalidLength) {
            RawLayoutType = rawLayoutType;
            Layout = layout;
            PlaceholderTypes = placeholderTypes?.ToArray()
                ?? throw new ArgumentNullException(nameof(placeholderTypes));
            MasterId = masterId;
            NotesId = notesId;
            Flags = flags;
            HasInvalidPlaceholderType = hasInvalidPlaceholderType;
            HasInvalidLength = hasInvalidLength;
        }

        internal uint RawLayoutType { get; }
        internal LegacyPptSlideLayoutType? Layout { get; }
        internal IReadOnlyList<LegacyPptPlaceholderKind> PlaceholderTypes { get; }
        internal uint MasterId { get; }
        internal uint NotesId { get; }
        internal ushort Flags { get; }
        internal bool HasInvalidPlaceholderType { get; }
        internal bool HasInvalidLength { get; }
        internal bool HasReservedFlags => (Flags & 0xFFF8) != 0;
        internal bool FollowsMasterObjects => (Flags & 0x0001) != 0;
        internal bool FollowsMasterColorScheme => (Flags & 0x0002) != 0;
        internal bool FollowsMasterBackground => (Flags & 0x0004) != 0;
    }

    internal enum LegacyPptPlaceholderReadStatus {
        Absent,
        NotPlaceholder,
        Decoded,
        Invalid
    }
}
