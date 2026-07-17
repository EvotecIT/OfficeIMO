namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Reads the fixed fields required from the Current User stream.</summary>
    internal readonly struct LegacyPptCurrentUserAtom {
        private const ushort RecordCurrentUser = 0x0FF6;
        private const int RecordHeaderLength = 8;
        private const int FixedPayloadLength = 20;
        internal const uint EncryptedHeaderToken = 0xF3D1C4DF;
        internal const uint UnencryptedHeaderToken = 0xE391C05F;

        private LegacyPptCurrentUserAtom(uint headerToken, uint currentEditOffset,
            bool hasFourBytePayloadOverstatement) {
            HeaderToken = headerToken;
            CurrentEditOffset = currentEditOffset;
            HasFourBytePayloadOverstatement = hasFourBytePayloadOverstatement;
        }

        internal uint HeaderToken { get; }

        internal uint CurrentEditOffset { get; }

        /// <summary>
        /// Gets whether recLen is four bytes larger than the available payload. Microsoft
        /// PowerPoint for Mac emits this form when the optional user-name fields are empty.
        /// </summary>
        internal bool HasFourBytePayloadOverstatement { get; }

        internal static LegacyPptCurrentUserAtom Read(byte[] stream) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            int minimumLength = RecordHeaderLength + FixedPayloadLength;
            if (stream.Length < minimumLength) {
                throw new InvalidDataException("The CurrentUserAtom fixed fields are truncated.");
            }

            ushort versionAndInstance = ReadUInt16(stream, 0);
            ushort recordType = ReadUInt16(stream, 2);
            uint declaredPayloadLength = ReadUInt32(stream, 4);
            if (versionAndInstance != 0 || recordType != RecordCurrentUser
                || declaredPayloadLength < FixedPayloadLength) {
                throw new InvalidDataException("The Current User stream does not contain a valid CurrentUserAtom.");
            }

            long declaredRecordLength = RecordHeaderLength + declaredPayloadLength;
            bool hasFourBytePayloadOverstatement = declaredPayloadLength == 28
                && stream.Length == 32
                && declaredRecordLength == stream.Length + 4L
                && ReadUInt16(stream, RecordHeaderLength + 12) == 0;
            if (declaredRecordLength > stream.Length && !hasFourBytePayloadOverstatement) {
                throw new InvalidDataException(
                    $"The CurrentUserAtom declares {declaredPayloadLength} payload bytes, but the Current User stream contains {stream.Length} bytes.");
            }
            if (ReadUInt32(stream, RecordHeaderLength) != FixedPayloadLength) {
                throw new InvalidDataException("The CurrentUserAtom fixed-field size is invalid.");
            }

            return new LegacyPptCurrentUserAtom(
                ReadUInt32(stream, RecordHeaderLength + 4),
                ReadUInt32(stream, RecordHeaderLength + 8),
                hasFourBytePayloadOverstatement);
        }

        private static ushort ReadUInt16(byte[] bytes, int offset) => unchecked((ushort)(bytes[offset]
            | (bytes[offset + 1] << 8)));

        private static uint ReadUInt32(byte[] bytes, int offset) => unchecked((uint)(bytes[offset]
            | (bytes[offset + 1] << 8)
            | (bytes[offset + 2] << 16)
            | (bytes[offset + 3] << 24)));
    }
}
