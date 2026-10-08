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
        private bool TryFindExifItemId(byte[] data, Box itemInfoBox, out uint itemId, bool requireSupportedPayload = false) {
            itemId = 0;

            int offset = itemInfoBox.DataOffset;
            if (!TryReadFullBoxHeader(data, offset, itemInfoBox.EndOffset, out byte version, out _, out offset)) {
                return false;
            }

            uint entryCount;
            if (version == 0) {
                if (!TryReadUInt16(data, offset, itemInfoBox.EndOffset, out ushort count)) {
                    return false;
                }

                entryCount = count;
                offset += 2;
            } else {
                if (!TryReadUInt32(data, offset, itemInfoBox.EndOffset, out entryCount)) {
                    return false;
                }

                offset += 4;
            }

            if (entryCount > 4096) {
                throw new FormatException("HEIF item count exceeds structure limits.");
            }
            for (uint index = 0; index < entryCount; index++) {
                CheckWork();
                if (!TryReadBox(data, offset, itemInfoBox.EndOffset, out Box entryBox)) {
                    return false;
                }

                if (entryBox.Type == "infe" && TryReadItemInfoEntry(data, entryBox, out uint entryItemId, out string itemType, out _, out ushort protection, out _, out _, out _) && itemType == ExifItemType) {
                    if (requireSupportedPayload && protection != 0) {
                        return false;
                    }
                    itemId = entryItemId;
                    return true;
                }

                offset = entryBox.EndOffset;
            }

            return false;
        }

        private bool TryFindXmpItemId(byte[] data, Box itemInfoBox, out uint itemId, bool requireSupportedPayload = false) {
            itemId = 0;

            int offset = itemInfoBox.DataOffset;
            if (!TryReadFullBoxHeader(data, offset, itemInfoBox.EndOffset, out byte version, out _, out offset)) {
                return false;
            }

            uint entryCount;
            if (version == 0) {
                if (!TryReadUInt16(data, offset, itemInfoBox.EndOffset, out ushort count)) {
                    return false;
                }

                entryCount = count;
                offset += 2;
            } else {
                if (!TryReadUInt32(data, offset, itemInfoBox.EndOffset, out entryCount)) {
                    return false;
                }

                offset += 4;
            }

            if (entryCount > 4096) {
                throw new FormatException("HEIF item count exceeds structure limits.");
            }
            for (uint index = 0; index < entryCount; index++) {
                CheckWork();
                if (!TryReadBox(data, offset, itemInfoBox.EndOffset, out Box entryBox)) {
                    return false;
                }

                if (entryBox.Type == "infe" &&
                    TryReadItemInfoEntry(data, entryBox, out uint entryItemId, out _, out _, out ushort protection, out _, out string? mimeType, out string? contentEncoding) &&
                    IsXmpMimeType(mimeType)) {
                    if (requireSupportedPayload && (protection != 0 || !string.IsNullOrEmpty(contentEncoding))) {
                        return false;
                    }
                    itemId = entryItemId;
                    return true;
                }

                offset = entryBox.EndOffset;
            }

            return false;
        }

        private bool TryReadItemInfos(byte[] data, Box itemInfoBox, List<HeifItemInfoBuilder> items) {
            int offset = itemInfoBox.DataOffset;
            if (!TryReadFullBoxHeader(data, offset, itemInfoBox.EndOffset, out byte version, out _, out offset)) {
                return false;
            }

            uint entryCount;
            if (version == 0) {
                if (!TryReadUInt16(data, offset, itemInfoBox.EndOffset, out ushort count)) {
                    return false;
                }

                entryCount = count;
                offset += 2;
            } else {
                if (!TryReadUInt32(data, offset, itemInfoBox.EndOffset, out entryCount)) {
                    return false;
                }

                offset += 4;
            }

            if (entryCount > 4096) {
                throw new FormatException("HEIF item count exceeds structure limits.");
            }
            for (uint index = 0; index < entryCount; index++) {
                CheckWork();
                if (!TryReadBox(data, offset, itemInfoBox.EndOffset, out Box entryBox)) {
                    return false;
                }

                if (entryBox.Type == "infe" && TryReadItemInfoEntry(data, entryBox, out uint itemId, out string itemType, out string itemName, out ushort itemProtectionIndex, out bool isHidden, out string? mimeType, out string? contentEncoding)) {
                    items.Add(new HeifItemInfoBuilder(itemId, itemType, itemName, itemProtectionIndex, isHidden, mimeType, contentEncoding));
                }

                offset = entryBox.EndOffset;
            }

            return true;
        }

        private bool TryReadItemInfoEntry(byte[] data, Box entryBox, out uint itemId, out string itemType, out string itemName, out ushort itemProtectionIndex, out bool isHidden, out string? mimeType, out string? contentEncoding) {
            itemId = 0;
            itemType = string.Empty;
            itemName = string.Empty;
            itemProtectionIndex = 0;
            isHidden = false;
            mimeType = null;
            contentEncoding = null;

            int offset = entryBox.DataOffset;
            if (!TryReadFullBoxHeader(data, offset, entryBox.EndOffset, out byte version, out uint flags, out offset)) {
                return false;
            }

            isHidden = (flags & 1) == 1;

            if (version == 2) {
                if (!TryReadUInt16(data, offset, entryBox.EndOffset, out ushort shortItemId)) {
                    return false;
                }

                itemId = shortItemId;
                offset += 2;
            } else if (version == 3) {
                if (!TryReadUInt32(data, offset, entryBox.EndOffset, out itemId)) {
                    return false;
                }

                offset += 4;
            } else {
                return false;
            }

            if (!TryReadUInt16(data, offset, entryBox.EndOffset, out itemProtectionIndex)) {
                return false;
            }

            offset += 2;
            if (offset + 4 > entryBox.EndOffset) {
                return false;
            }

            itemType = ReadAscii(data, offset, 4);
            offset += 4;
            itemName = ReadNullTerminatedString(data, offset, entryBox.EndOffset, out offset);
            if (itemType == "mime" && offset < entryBox.EndOffset) {
                mimeType = ReadNullTerminatedString(data, offset, entryBox.EndOffset, out offset);
                if (offset < entryBox.EndOffset) {
                    contentEncoding = ReadNullTerminatedString(data, offset, entryBox.EndOffset, out _);
                }
            }

            return true;
        }

        private bool TryReadFileType(byte[] data, out string majorBrand, out uint minorVersion, out List<string>? compatibleBrands) {
            majorBrand = string.Empty;
            minorVersion = 0;
            compatibleBrands = null;

            foreach (Box box in EnumerateBoxes(data, 0, data.Length)) {
                CheckWork();
                if (box.Type != "ftyp") {
                    continue;
                }

                if (box.DataOffset + 8 > box.EndOffset) {
                    return false;
                }

                majorBrand = ReadAscii(data, box.DataOffset, 4);
                minorVersion = ReadUInt32(data, box.DataOffset + 4);
                compatibleBrands = new List<string>();
                for (int offset = box.DataOffset + 8; offset + 4 <= box.EndOffset; offset += 4) {
                    CheckWork();
                    compatibleBrands.Add(ReadAscii(data, offset, 4));
                }

                return true;
            }

            return false;
        }

        private bool IsXmpMimeType(string? mimeType) =>
            string.Equals(mimeType, XmpMimeType, StringComparison.OrdinalIgnoreCase);

        private bool TryReadPrimaryItemId(byte[] data, Box primaryItemBox, out uint primaryItemId) {
            primaryItemId = 0;

            int offset = primaryItemBox.DataOffset;
            if (!TryReadFullBoxHeader(data, offset, primaryItemBox.EndOffset, out byte version, out _, out offset)) {
                return false;
            }

            if (version == 0) {
                if (!TryReadUInt16(data, offset, primaryItemBox.EndOffset, out ushort shortItemId)) {
                    return false;
                }

                primaryItemId = shortItemId;
                return true;
            }

            return TryReadUInt32(data, offset, primaryItemBox.EndOffset, out primaryItemId);
        }

    }
}
