namespace OfficeIMO.OneNote;

internal sealed class FssHttpStreamObject {
    internal FssHttpStreamObject(int type, bool compound, long headerOffset, long dataOffset, ulong dataLength, IReadOnlyList<FssHttpStreamObject> children) {
        Type = type;
        Compound = compound;
        HeaderOffset = headerOffset;
        DataOffset = dataOffset;
        DataLength = dataLength;
        Children = children;
    }

    internal int Type { get; }
    internal bool Compound { get; }
    internal long HeaderOffset { get; }
    internal long DataOffset { get; }
    internal ulong DataLength { get; }
    internal IReadOnlyList<FssHttpStreamObject> Children { get; }
}

internal static class FssHttpStreamObjectReader {
    internal static FssHttpStreamObject ReadPackaging(Stream stream, OneNoteReaderOptions options) {
        if (stream.Length < OneNoteFormatConstants.PackageStoreFixedPrefixLength + 4) {
            throw new OneNoteFormatException("ONENOTE_PACKAGE_TRUNCATED", "The package-store file is truncated.", stream.Length);
        }
        stream.Position = OneNoteFormatConstants.PackageStoreFixedPrefixLength - 4;
        var state = new ReaderState(stream, options);
        FssHttpStreamObject root = state.ReadObject(0);
        if (root.Type != 0x7A || !root.Compound) {
            throw new OneNoteFormatException("ONENOTE_PACKAGE_START", "The package does not contain the required packaging stream object.", root.HeaderOffset);
        }
        while (stream.Position < stream.Length) {
            long offset = stream.Position;
            if (stream.ReadByte() != 0) {
                throw new OneNoteFormatException("ONENOTE_PACKAGE_TRAILING_DATA", "Non-zero bytes follow the packaging stream object.", offset);
            }
        }
        return root;
    }

    internal static byte[] ReadData(Stream stream, FssHttpStreamObject item, ulong maxBytes, string name) {
        if (item.DataLength > maxBytes || item.DataLength > int.MaxValue) {
            throw new OneNoteFormatException("ONENOTE_PACKAGE_ITEM_SIZE", "The " + name + " exceeds its configured materialization limit.", item.DataOffset);
        }
        long original = stream.Position;
        try {
            stream.Position = item.DataOffset;
            var bytes = new byte[(int)item.DataLength];
            ReadExactly(stream, bytes, 0, bytes.Length, item.DataOffset);
            return bytes;
        } finally {
            stream.Position = original;
        }
    }

    private sealed class ReaderState {
        private readonly Stream _stream;
        private readonly OneNoteReaderOptions _options;
        private int _objectCount;

        internal ReaderState(Stream stream, OneNoteReaderOptions options) {
            _stream = stream;
            _options = options;
        }

        internal FssHttpStreamObject ReadObject(int depth) {
            if (depth >= _options.MaxStreamObjectDepth) {
                throw new OneNoteFormatException("ONENOTE_PACKAGE_DEPTH", "The compound stream-object depth limit was exceeded.", _stream.Position);
            }
            if (_objectCount++ >= _options.MaxStreamObjects) {
                throw new OneNoteFormatException("ONENOTE_PACKAGE_OBJECT_LIMIT", "The stream-object count limit was exceeded.", _stream.Position);
            }

            Header header = ReadHeader();
            if (header.IsEnd) {
                throw new OneNoteFormatException("ONENOTE_PACKAGE_UNEXPECTED_END", "An unexpected stream-object end header was encountered.", header.Offset);
            }
            long dataOffset = _stream.Position;
            EnsureRange(header.Length, dataOffset);
            _stream.Position += checked((long)header.Length);
            IReadOnlyList<FssHttpStreamObject> children = Array.Empty<FssHttpStreamObject>();
            if (header.Compound) {
                var list = new List<FssHttpStreamObject>();
                while (true) {
                    long childOffset = _stream.Position;
                    Header next = ReadHeader();
                    if (next.IsEnd) {
                        if (next.Type != header.Type) {
                            throw new OneNoteFormatException("ONENOTE_PACKAGE_END_TYPE", "A compound stream object has a mismatched end header.", next.Offset);
                        }
                        break;
                    }
                    _stream.Position = childOffset;
                    list.Add(ReadObject(depth + 1));
                }
                children = list.AsReadOnly();
            }
            return new FssHttpStreamObject(header.Type, header.Compound, header.Offset, dataOffset, header.Length, children);
        }

        private Header ReadHeader() {
            long offset = _stream.Position;
            int firstValue = _stream.ReadByte();
            if (firstValue < 0) throw new OneNoteFormatException("ONENOTE_PACKAGE_HEADER_EOF", "The file ended while reading a stream-object header.", offset);
            byte first = (byte)firstValue;
            int headerType = first & 0x03;
            if (headerType == 0x01) return new Header(offset, first >> 2, false, true, 0);

            int secondValue = _stream.ReadByte();
            if (secondValue < 0) throw new OneNoteFormatException("ONENOTE_PACKAGE_HEADER_EOF", "The file ended while reading a stream-object header.", offset);
            ushort raw16 = (ushort)(first | (secondValue << 8));
            if (headerType == 0x03) return new Header(offset, raw16 >> 2, false, true, 0);
            if (headerType == 0x00) {
                bool compound = (raw16 & 0x04) != 0;
                int type = (raw16 >> 3) & 0x3F;
                ulong length = (uint)(raw16 >> 9);
                return new Header(offset, type, compound, false, length);
            }
            if (headerType != 0x02) {
                throw new OneNoteFormatException("ONENOTE_PACKAGE_HEADER_TYPE", "An invalid stream-object header type was encountered.", offset);
            }

            int third = _stream.ReadByte();
            int fourth = _stream.ReadByte();
            if (third < 0 || fourth < 0) throw new OneNoteFormatException("ONENOTE_PACKAGE_HEADER_EOF", "The file ended while reading a 32-bit stream-object header.", offset);
            uint raw32 = (uint)(raw16 | (third << 16) | (fourth << 24));
            bool isCompound = (raw32 & 0x04U) != 0;
            int objectType = (int)((raw32 >> 3) & 0x3FFFU);
            ulong objectLength = raw32 >> 17;
            if (objectLength == 0x7FFFU) objectLength = ReadCompactUInt64(_stream, offset);
            return new Header(offset, objectType, isCompound, false, objectLength);
        }

        private void EnsureRange(ulong count, long offset) {
            if (count > long.MaxValue || offset > _stream.Length || (ulong)(_stream.Length - offset) < count) {
                throw new OneNoteFormatException("ONENOTE_PACKAGE_BOUNDS", "A stream-object payload lies outside the package.", offset);
            }
        }
    }

    private readonly struct Header {
        internal Header(long offset, int type, bool compound, bool isEnd, ulong length) {
            Offset = offset;
            Type = type;
            Compound = compound;
            IsEnd = isEnd;
            Length = length;
        }
        internal long Offset { get; }
        internal int Type { get; }
        internal bool Compound { get; }
        internal bool IsEnd { get; }
        internal ulong Length { get; }
    }

    internal static ulong ReadCompactUInt64(Stream stream, long offset) {
        int firstValue = stream.ReadByte();
        if (firstValue < 0) throw new OneNoteFormatException("ONENOTE_PACKAGE_COMPACT_UINT", "The file ended while reading a compact integer.", offset);
        byte first = (byte)firstValue;
        if (first == 0) return 0;
        if (first == 0x80) {
            var bytes = new byte[8];
            ReadExactly(stream, bytes, 0, 8, offset + 1);
            return OneNoteBinary.ReadUInt64(bytes, 0);
        }
        int encodedLength = 1;
        byte mask = 1;
        while ((first & mask) == 0 && encodedLength < 8) {
            encodedLength++;
            mask <<= 1;
        }
        if (encodedLength > 7) throw new OneNoteFormatException("ONENOTE_PACKAGE_COMPACT_UINT", "A compact integer uses an invalid type prefix.", offset);
        ulong raw = first;
        for (int index = 1; index < encodedLength; index++) {
            int value = stream.ReadByte();
            if (value < 0) throw new OneNoteFormatException("ONENOTE_PACKAGE_COMPACT_UINT", "The file ended while reading a compact integer.", offset);
            raw |= (ulong)(byte)value << (index * 8);
        }
        return raw >> encodedLength;
    }

    internal static void ReadExactly(Stream stream, byte[] buffer, int offset, int count, long errorOffset) {
        int total = 0;
        while (total < count) {
            int read = stream.Read(buffer, offset + total, count - total);
            if (read <= 0) throw new OneNoteFormatException("ONENOTE_PACKAGE_TRUNCATED", "The file ended while reading package data.", errorOffset + total);
            total += read;
        }
    }
}
