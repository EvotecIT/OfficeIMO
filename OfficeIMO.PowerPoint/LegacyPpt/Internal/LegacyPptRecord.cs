using System.Text;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    internal sealed class LegacyPptRecord {
        private readonly byte[] _source;
        private readonly List<LegacyPptRecord> _children;

        internal LegacyPptRecord(byte[] source, int offset, byte version, ushort instance, ushort type,
            int payloadOffset, int payloadLength, List<LegacyPptRecord>? children = null) {
            _source = source;
            Offset = offset;
            Version = version;
            Instance = instance;
            Type = type;
            PayloadOffset = payloadOffset;
            PayloadLength = payloadLength;
            _children = children ?? new List<LegacyPptRecord>();
        }

        internal int Offset { get; }

        internal byte Version { get; }

        internal ushort Instance { get; }

        internal ushort Type { get; }

        internal int PayloadOffset { get; }

        internal int PayloadLength { get; }

        internal int EndOffset => checked(PayloadOffset + PayloadLength);

        internal IReadOnlyList<LegacyPptRecord> Children => _children;

        internal ushort ReadUInt16(int relativeOffset) {
            EnsureAvailable(relativeOffset, 2);
            return unchecked((ushort)(_source[PayloadOffset + relativeOffset]
                | (_source[PayloadOffset + relativeOffset + 1] << 8)));
        }

        internal short ReadInt16(int relativeOffset) => unchecked((short)ReadUInt16(relativeOffset));

        internal uint ReadUInt32(int relativeOffset) {
            EnsureAvailable(relativeOffset, 4);
            int offset = PayloadOffset + relativeOffset;
            return unchecked((uint)(_source[offset]
                | (_source[offset + 1] << 8)
                | (_source[offset + 2] << 16)
                | (_source[offset + 3] << 24)));
        }

        internal int ReadInt32(int relativeOffset) => unchecked((int)ReadUInt32(relativeOffset));

        internal byte ReadByte(int relativeOffset) {
            EnsureAvailable(relativeOffset, 1);
            return _source[PayloadOffset + relativeOffset];
        }

        internal string ReadUtf16Text() => ReadUtf16Text(0, PayloadLength);

        internal string ReadUtf16Text(int relativeOffset, int byteCount) {
            if ((byteCount & 1) != 0) {
                throw new InvalidDataException($"Record 0x{Type:X4} at 0x{Offset:X} has an odd UTF-16 payload length.");
            }
            EnsureAvailable(relativeOffset, byteCount);
            return Encoding.Unicode.GetString(_source, PayloadOffset + relativeOffset, byteCount);
        }

        internal string ReadLowByteUnicodeText() {
            var characters = new char[PayloadLength];
            for (int index = 0; index < PayloadLength; index++) {
                characters[index] = (char)_source[PayloadOffset + index];
            }
            return new string(characters);
        }

        internal byte[] CopyRecordBytes() {
            byte[] bytes = new byte[checked(8 + PayloadLength)];
            Buffer.BlockCopy(_source, Offset, bytes, 0, bytes.Length);
            return bytes;
        }

        internal IEnumerable<LegacyPptRecord> DescendantsAndSelf() {
            yield return this;
            foreach (LegacyPptRecord child in _children) {
                foreach (LegacyPptRecord descendant in child.DescendantsAndSelf()) {
                    yield return descendant;
                }
            }
        }

        private void EnsureAvailable(int relativeOffset, int count) {
            if (relativeOffset < 0 || count < 0 || relativeOffset > PayloadLength - count) {
                throw new InvalidDataException(
                    $"Record 0x{Type:X4} at 0x{Offset:X} is too short for a {count}-byte field at payload offset {relativeOffset}.");
            }
        }
    }
}
