// Adapted from the Evotec ImagePlayground HEIF metadata implementation.
// Copyright (c) 2022 Evotec. MIT license; see Licenses/ImagePlayground-LICENSE.txt.
// OfficeIMO adaptation provides bounded byte/stream APIs, cancellation and output preflight.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeHeifMetadataReader {
    private sealed partial class Parser {
        private bool TryFindMetaBox(byte[] data, out Box metaBox) {
            foreach (Box box in EnumerateBoxes(data, 0, data.Length)) {
                CheckWork();
                if (box.Type == "meta") {
                    metaBox = box;
                    return true;
                }
            }

            metaBox = default;
            return false;
        }

        private bool TryFindItemInfoBox(byte[] data, out Box itemInfoBox) {
            if (!TryFindMetaBox(data, out Box metaBox)) {
                itemInfoBox = default;
                return false;
            }

            int metaChildrenStart = metaBox.DataOffset + 4;
            if (metaChildrenStart > metaBox.EndOffset) {
                itemInfoBox = default;
                return false;
            }

            foreach (Box childBox in EnumerateBoxes(data, metaChildrenStart, metaBox.EndOffset)) {
                CheckWork();
                if (childBox.Type == "iinf") {
                    itemInfoBox = childBox;
                    return true;
                }
            }

            itemInfoBox = default;
            return false;
        }

        private IEnumerable<Box> EnumerateBoxes(byte[] data, int startOffset, int endOffset) {
            int offset = startOffset;
            while (offset < endOffset) {
                CheckWork();
                if (!TryReadBox(data, offset, endOffset, out Box box)) {
                    throw new FormatException("Truncated or invalid HEIF box collection.");
                }

                yield return box;
                offset = box.EndOffset;
            }
        }

        private bool TryReadBox(byte[] data, int offset, int endOffset, out Box box) {
            box = default;
            if (offset < 0 || offset + 8 > endOffset) {
                return false;
            }

            ulong size = ReadUInt32(data, offset);
            string type = ReadAscii(data, offset + 4, 4);
            int headerSize = 8;

            if (size == 1) {
                if (offset + 16 > endOffset) {
                    return false;
                }

                size = ReadUInt64(data, offset + 8);
                headerSize = 16;
            } else if (size == 0) {
                size = (ulong)(endOffset - offset);
            }

            if (size < (ulong)headerSize || size > int.MaxValue || size > (ulong)(endOffset - offset)) {
                return false;
            }

            int boxEnd = offset + (int)size;
            box = new Box(offset, boxEnd, offset + headerSize, type);
            return true;
        }

        private bool TryReadFullBoxHeader(byte[] data, int offset, int endOffset, out byte version, out uint flags, out int nextOffset) {
            version = 0;
            flags = 0;
            nextOffset = offset;

            if (offset + 4 > endOffset) {
                return false;
            }

            version = data[offset];
            flags = (uint)((data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3]);
            nextOffset = offset + 4;
            return true;
        }

        private bool IsSupportedFieldSize(int size) =>
            size == 0 || size == 4 || size == 8;

        private bool TryReadVariableUInt(byte[] data, int offset, int endOffset, int size, out ulong value, out int nextOffset) {
            value = 0;
            nextOffset = offset;

            if (size == 0) {
                return true;
            }

            if (offset + size > endOffset) {
                return false;
            }

            value = size == 4
                ? ReadUInt32(data, offset)
                : ReadUInt64(data, offset);
            nextOffset = offset + size;
            return true;
        }

        private bool TryReadUInt16(byte[] data, int offset, int endOffset, out ushort value) {
            value = 0;
            if (offset + 2 > endOffset) {
                return false;
            }

            value = (ushort)((data[offset] << 8) | data[offset + 1]);
            return true;
        }

        private bool TryReadUInt32(byte[] data, int offset, int endOffset, out uint value) {
            value = 0;
            if (offset + 4 > endOffset) {
                return false;
            }

            value = ReadUInt32(data, offset);
            return true;
        }

        private uint ReadUInt32(byte[] data, int offset) =>
            ((uint)data[offset] << 24) |
            ((uint)data[offset + 1] << 16) |
            ((uint)data[offset + 2] << 8) |
            data[offset + 3];

        private ulong ReadUInt64(byte[] data, int offset) =>
            ((ulong)ReadUInt32(data, offset) << 32) | ReadUInt32(data, offset + 4);

        private string ReadAscii(byte[] data, int offset, int length) =>
            System.Text.Encoding.ASCII.GetString(data, offset, length);

        private string ReadNullTerminatedString(byte[] data, int offset, int endOffset, out int nextOffset) {
            int terminatorOffset = offset;
            while (terminatorOffset < endOffset && data[terminatorOffset] != 0) {
                if ((terminatorOffset & 4095) == 0) {
                    _cancellationToken.ThrowIfCancellationRequested();
                }
                terminatorOffset++;
            }

            if (terminatorOffset == endOffset) {
                throw new FormatException("Truncated HEIF null-terminated text field.");
            }
            int stringLength = terminatorOffset - offset;
            if (stringLength > OfficeExifProfileCodec.MaximumProfileBytes) {
                throw new FormatException("HEIF text exceeds metadata limits.");
            }
            ReserveBytes(stringLength * 2L + 24L);
            nextOffset = terminatorOffset < endOffset
                ? terminatorOffset + 1
                : terminatorOffset;

            return terminatorOffset == offset
                ? string.Empty
                : System.Text.Encoding.UTF8.GetString(data, offset, terminatorOffset - offset);
        }

        private bool LooksLikeTiffHeader(byte[] data, int offset) {
            if (offset + 4 > data.Length) {
                return false;
            }

            bool littleEndian = data[offset] == (byte)'I' && data[offset + 1] == (byte)'I' && data[offset + 2] == 42 && data[offset + 3] == 0;
            bool bigEndian = data[offset] == (byte)'M' && data[offset + 1] == (byte)'M' && data[offset + 2] == 0 && data[offset + 3] == 42;
            return littleEndian || bigEndian;
        }

        private byte[] CopyTail(byte[] data, int offset) {
            ReserveBytes(data.LongLength - offset + 24L);
            var tail = new byte[data.Length - offset];
            Buffer.BlockCopy(data, offset, tail, 0, tail.Length);
            return tail;
        }

        private byte[] CopyRange(byte[] data, int offset, int length) {
            if (length > OfficeExifProfileCodec.MaximumProfileBytes) {
                throw new FormatException("HEIF property exceeds metadata limits.");
            }
            ReserveBytes(length + 24L);
            var range = new byte[length];
            Buffer.BlockCopy(data, offset, range, 0, range.Length);
            return range;
        }

        private byte[] CreateExifItemData(byte[] tiffPayload) {
            if (checked(_sourceBytes + _allocatedBytes + tiffPayload.LongLength + 34L) > OfficeRasterGuards.MaximumDecodedBytes) {
                throw new FormatException("HEIF replacement metadata exceeds the managed working-set limit.");
            }
            var exifItemData = new byte[10 + tiffPayload.Length];
            WriteVariableUInt(exifItemData, 0, 4, 6);
            exifItemData[4] = (byte)'E';
            exifItemData[5] = (byte)'x';
            exifItemData[6] = (byte)'i';
            exifItemData[7] = (byte)'f';
            exifItemData[8] = 0;
            exifItemData[9] = 0;
            Buffer.BlockCopy(tiffPayload, 0, exifItemData, 10, tiffPayload.Length);
            return exifItemData;
        }

        private bool CanWriteVariableUInt(ulong value, int size) =>
            size == 0
                ? value == 0
                : size == 8 || (size == 4 && value <= uint.MaxValue);

        private void WriteVariableUInt(byte[] data, int offset, int size, ulong value) {
            if (size == 0) {
                return;
            }

            if (size == 8) {
                data[offset] = (byte)(value >> 56);
                data[offset + 1] = (byte)(value >> 48);
                data[offset + 2] = (byte)(value >> 40);
                data[offset + 3] = (byte)(value >> 32);
                data[offset + 4] = (byte)(value >> 24);
                data[offset + 5] = (byte)(value >> 16);
                data[offset + 6] = (byte)(value >> 8);
                data[offset + 7] = (byte)value;
                return;
            }

            data[offset] = (byte)(value >> 24);
            data[offset + 1] = (byte)(value >> 16);
            data[offset + 2] = (byte)(value >> 8);
            data[offset + 3] = (byte)value;
        }

    }
}
